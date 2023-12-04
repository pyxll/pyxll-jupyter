"""
Start an IPython Qt console or notebook connected to the python session
running in Excel.

This requires sys.executable to be set, and so it's recommended
that the following is added to the pyxll.cfg file:

[PYTHON]
executable = <path to your python installation>/pythonw.exe
"""
from .magic import ExcelMagics
from ipykernel.kernelapp import IPKernelApp
from ipykernel.embed import embed_kernel
from pyxll import schedule_call, get_config
import pyxll
import importlib.util
import subprocess
import threading
import logging
import ctypes
import atexit
import queue
import zmq
import sys
import os
import re

if zmq.pyzmq_version_info()[0] >= 17:
    from tornado.ioloop import IOLoop
else:
    from zmq.eventloop.ioloop import IOLoop


# Command line args for the different Jupyter subcommands
_subcommand_jupyter_args = {
    # subcommand
    "notebook": [(
        # requirement
        "jupyter-client >= 7.0.0", [
            # args
            "--ServerApp.kernel_manager_class=pyxll_jupyter.kernel_managers.nbkm.ExternalMappingKernelManager",
            "--KernelProvisionerFactory.default_provisioner_name=pyxll-provisioner"
        ]), (
        # requirement
        "jupyter-client >= 6.0.0", [
            # args (KernelProvisionerFactory not available before version 7.0)
            "--NotebookApp.kernel_manager_class=pyxll_jupyter.kernel_managers.extipy.ExternalMappingKernelManager"
        ])
    ],

    # subcommand
    "lab": [(
        # requirement
        "jupyterlab >= 4.0.0", [
            # args
            "--ServerApp.kernel_manager_class=pyxll_jupyter.kernel_managers.labkm.ExternalMappingKernelManager",
            "--KernelProvisionerFactory.default_provisioner_name=pyxll-provisioner"
        ])
    ]
}

# Query args to add to the URL to open the client
_subcommand_query_params = {
    "notebook": [],
    "lab": []
}

_log = logging.getLogger(__name__)
_all_jupyter_processes = []

try:
    # pywintypes needs to be imported before win32api for some Python installs.
    # See https://github.com/mhammond/pywin32/issues/1399
    import pywintypes
    import win32api
except ImportError:
    pywintypes = None
    win32api = None


if getattr(sys, "_ipython_kernel_running", None) is None:
    sys._ipython_kernel_running = False

if getattr(sys, "_ipython_app", None) is None:
    sys._ipython_app = False


def _which(program):
    """find an exe's full path by looking at the PATH environment variable"""
    def is_exe(fpath):
        return os.path.isfile(fpath) and os.access(fpath, os.X_OK)

    fpath, fname = os.path.split(program)
    if fpath:
        if is_exe(program):
            return program
    else:
        for path in os.environ["PATH"].split(os.pathsep):
            path = path.strip('"')
            exe_file = os.path.join(path, program)
            if is_exe(exe_file):
                return exe_file

    return None


def _get_excel_path():
    """Run the path of the Excel executable"""
    buffer_len = 260
    buffer = ctypes.create_unicode_buffer(buffer_len)
    num_chars = ctypes.windll.kernel32.GetModuleFileNameW(None, buffer, buffer_len)
    while num_chars >= 0 and buffer_len <= num_chars:
        print(f"num_chars = {num_chars}; buffer_len = {buffer_len}")
        buffer_len += 260
        buffer = ctypes.create_unicode_buffer(buffer_len)
        num_chars = ctypes.windll.kernel32.GetModuleFileNameW(None, buffer, buffer_len)

    if num_chars <= 0:
        error = ctypes.windll.kernel32.GetLastError()
        raise RuntimeError(f"Error getting Excel executable path: {error}")

    return buffer.value


def _get_connection_dir(app):
    """Gets the connection dir to use for the IPKernelApp"""
    connection_dir = None

    # Check the pyxll config first
    cfg = get_config()
    if cfg.has_option("JUPYTER", "runtime_dir"):
        connection_dir = cfg.get("JUPYTER", "runtime_dir")
        if not os.path.abspath(connection_dir):
            connection_dir = os.path.join(os.path.dirname(pyxll.__file__), connection_dir)

    # If not set in the pyxll config use the default from the kernel
    if not connection_dir:
        connection_dir = app.connection_dir

    # If Excel is installed as a UWP then AppData will appear as a different folder when the
    # child Python process is run so use a different path.
    excel_path = _get_excel_path()
    if "WindowsApps" in re.split(r"[/\\]+", excel_path):
        _log.debug("Excel looks like a UWP app.")
        if "AppData" in re.split(r"[/\\]+", connection_dir):
            connection_dir = os.path.join(os.path.dirname(pyxll.__file__), ".jupyter", "runtime")
            _log.warning("Jupyter's runtime directory is in AppData but Excel is installed as a UWP. ")
            _log.warning(f"{connection_dir} will be used instead.")
            _log.warning("Set 'runtime_dir' in the '[JUPYTER]' section of your pyxll.cfg to change this directory.")

    return connection_dir


class PushStdout:
    """Context manage to temporarily replace stdout/stderr."""

    def __init__(self, stdout, stderr):
        self.__stdout = stdout
        self.__stderr = stderr

    def __enter__(self):
        self.__orig_stdout = sys.stdout
        self.__orig_stderr = sys.stderr
        sys.stdout = self.__stdout
        sys.stderr = self.__stderr

    def __exit__(self, exc_type, exc_val, exc_tb):
        sys.stdout = self.__orig_stdout
        sys.stderr = self.__orig_stderr


def start_kernel():
    """starts the ipython kernel and returns the ipython app"""
    if sys._ipython_app and sys._ipython_kernel_running:
        return sys._ipython_app

    # The stdout/stderrs used by IPython. These get set after the kernel has started.
    ipy_stdout = sys.stdout
    ipy_stderr = sys.stderr

    # patch IPKernelApp.start so that it doesn't block
    def _IPKernelApp_start(self):
        nonlocal ipy_stdout, ipy_stderr

        if self.poller is not None:
            self.poller.start()
        self.kernel.start()

        # set up a timer to periodically poll the zmq ioloop
        self.loop = IOLoop.current()

        def poll_ioloop():
            try:
                # Use the IPython stdout/stderr while running the kernel
                with PushStdout(ipy_stdout, ipy_stderr):
                    # If the kernel has been closed then run the event loop until it gets to the
                    # stop event added by IPKernelApp.shutdown_request
                    if self.kernel.shell.exit_now:
                        _log.debug("IPython kernel stopping (%s)" % self.connection_file)
                        self.loop.start()
                        sys._ipython_kernel_running = False
                        return

                    # otherwise call the event loop but stop immediately if there are no pending events
                    self.loop.add_timeout(0, lambda: self.loop.add_callback(self.loop.stop))
                    self.loop.start()
            except:
                _log.error("Error polling Jupyter loop", exc_info=True)

            schedule_call(poll_ioloop, delay=0.1)

        sys._ipython_kernel_running = True
        schedule_call(poll_ioloop, delay=0.1)

    IPKernelApp.start = _IPKernelApp_start

    # IPython expects sys.__stdout__ to be set, and keep the original values to
    # be used after IPython has set its own.
    sys.__stdout__ = sys_stdout = sys.stdout
    sys.__stderr__ = sys_stderr = sys.stderr

    # Get or create the IPKernelApp instance and set the 'connection_dir' property
    if IPKernelApp.initialized():
        ipy = IPKernelApp.instance()
    else:
        ipy = IPKernelApp.instance()
        ipy.connection_dir = _get_connection_dir(ipy)
        ipy.initialize([])

    # call the API embed function, which will use the monkey-patched method above
    embed_kernel(local_ns={})

    # register the magic functions
    ipy.shell.register_magics(ExcelMagics)

    # Keep a reference to the kernel even if this module is reloaded
    sys._ipython_app = ipy

    # Restore sys stdout/stderr and keep track of the IPython versions
    ipy_stdout = sys.stdout
    ipy_stderr = sys.stderr
    sys.stdout = sys_stdout
    sys.stderr = sys_stderr

    # patch user_global_ns so that it always references the user_ns dict
    setattr(ipy.shell.__class__, 'user_global_ns', property(lambda self: self.user_ns))

    # patch ipapp so anything else trying to get a terminal app (e.g. ipdb) gets our IPKernalApp.
    from IPython.terminal.ipapp import TerminalIPythonApp
    TerminalIPythonApp.instance = lambda: ipy

    # Use the inline matplotlib backend
    mpl = ipy.shell.find_magic("matplotlib")
    if mpl:
        try:
            mpl("inline")
        except ImportError:
            pass

    return ipy


def _check_requirement(requirement):
    """Return True if a required package version is installed, False otherwise"""
    if sys.version_info[:2] >= (3, 10):
        from packaging.requirements import Requirement
        import importlib.metadata

        try:
            requirement = Requirement(requirement)
            version = importlib.metadata.version(requirement.name)
            return version in requirement.specifier
        except importlib.metadata.PackageNotFoundError:
            return False

    else:
        import pkg_resources

        try:
            pkg_resources.require(requirement)
            return True
        except (pkg_resources.VersionConflict, pkg_resources.DistributionNotFound):
            return


def _get_jupyter_args(subcommand):
    """Get the args for a Jupyter subcommand."""
    if subcommand not in _subcommand_jupyter_args:
        raise RuntimeError(f"Unexpected Juputer subcommand '{subcommand}'")

    # Find the args for the installed Jupyter version
    requirement = "unknown"
    for requirement, args in _subcommand_jupyter_args[subcommand]:
        if requirement is None or _check_requirement(requirement):
            return args

    raise RuntimeError(f"Requirements for Jupyter {subcommand} not satisfied ({requirement}).")


def _find_jupyter_script(subcommand="notebook"):
    """Returns the path to 'jupyter-notebook-script.py' or 'jupyter-lab-script.py" used to start
    the Jupyter notebook server. Returns None if the script can't be found.
    """
    # Look for it using importlib first
    spec = importlib.util.find_spec(f"jupyter-{subcommand}-script")
    if spec is not None and spec.origin and os.path.exists(spec.origin):
        return os.path.abspath(spec.origin)

    # If that doesn't work look in the Scripts folder
    if sys.executable and os.path.basename(sys.executable).lower() in ("python.exe", "pythonw.exe"):
        path = os.path.join(os.path.dirname(sys.executable), "Scripts", f"jupyter-{subcommand}-script.py")
        if os.path.exists(path):
            return os.path.abspath(path)

    return None


def _find_entry_point(dist, group, name):
    """Finds an entry point but doesn't load it."""
    try:
        from importlib.metadata import distribution

        def get_entry_info(spec, group, name):
            try:
                dist_name, _, _ = spec.partition('==')
                matches = (
                    entry_point
                    for entry_point in distribution(dist_name).entry_points
                    if entry_point.group == group and entry_point.name == name
                )
                return next(matches)
            except StopIteration:
                return None
    except ImportError:
        get_entry_info = None

    if get_entry_info is None:
        try:
            from pkg_resources import get_entry_info
        except ImportError:
            get_entry_info = None

    if get_entry_info is None:
        _log.warning("Unable to find 'get_entry_info' function.")
        return

    return get_entry_info(dist, group, name)


def _run_entry_point(dist, group, name):
    """Loads and runs an entry point"""
    ep = _find_entry_point(dist, group, name)
    cls = ep.load()
    return getattr(cls(), ep.attr)()


def _get_jupyter_python_script(subcommand="notebook"):
    """Try to determine a Python command that will start the Jupyter notebook.
    Returns None if the entry point wasn't found or couldn't be turned
    into a Python command.
    """
    # We only find the entry point here and don't try loading it.
    # That's because the jupyterlab entry point displays a console when imported
    # as it uses subprocess to get some npm information.
    dist = "jupyterlab" if subcommand == "lab" else "notebook"
    try:
        ep = _find_entry_point(dist, "console_scripts", f"jupyter-{subcommand}")
        if ep is None:
            _log.debug(f"Entry point {dist}.console_scripts.jupyter-{subcommand} not found.")
            return
    except:
        _log.debug(f"Error loading {dist}.console_scripts.jupyter-{subcommand} entry point.", exc_info=True)
        return

    try:
        return "; ".join((
            "import sys",
            f"from {__name__} import _run_entry_point",
            f"sys.exit(_run_entry_point('{dist}', '{ep.group}', '{ep.name}'))"
        ))
    except:
        _log.debug(f"Unexpected error getting jupyter-{subcommand} entry point", exc_info=True)


def _find_jupyter_cmd(subcommand="notebook"):
    """Find the 'jupyter-notebook' or 'jupyter-lab' executable or bat file.
    Returns None if it can't be found.
    """
    # Look in the python folder and in the scripts folder
    if sys.executable and os.path.basename(sys.executable).lower() in ("python.exe", "pythonw.exe"):
        for ext in (".exe", ".bat"):
            for path in (os.path.dirname(sys.executable), os.path.join(os.path.dirname(sys.executable), "Scripts")):
                jupyter_cmd = os.path.join(path, f"jupyter-{subcommand}{ext}")
                if os.path.exists(jupyter_cmd):
                    return os.path.abspath(jupyter_cmd)

    # If it wasn't found look for it on the system path
    for ext in (".exe", ".bat"):
        jupyter_cmd = _which(f"jupyter-{subcommand}{ext}")
        if jupyter_cmd is not None and os.path.exists(jupyter_cmd):
            return os.path.abspath(jupyter_cmd)

    return None


def launch_jupyter(initial_path=None,
                   notebook_path=None,
                   subcommand="notebook",
                   timeout=30,
                   no_browser=False):
    """Start the IPython kernel and launch a Jupyter notebook server as a child process.

    :param initial_path: Directory to start Jupyter in
    :param notebook_path: Path of notebook to open.
    :param timeout: Timeout in seconds to wait for the Jupyter process to start.
    :param no_browser: Don't open a web browser if False.
    :return: (Popen2 instance, URL string)
    """
    notebook = None
    if notebook_path is not None:
        if initial_path is not None:
            raise RuntimeError("'notebook_path' and 'initial_path' cannot be set together.")
        initial_path = os.path.dirname(notebook_path)
        notebook = os.path.basename(notebook_path)

    # Start the kernel and open Jupyter in a new tab
    app = start_kernel()
    connection_file = os.path.abspath(app.abs_connection_file)
    _log.debug(f"Kernel started with connection file '{connection_file}'")

    cmd = []
    pythonpath = list(sys.path)

    if sys.executable and os.path.basename(sys.executable).lower() in ("python.exe", "pythonw.exe"):
        python = os.path.join(os.path.dirname(sys.executable), "python.exe")
        if os.path.exists(python):
            jupyter_script = _find_jupyter_script(subcommand=subcommand)
            if jupyter_script:
                module, _ = os.path.splitext(os.path.basename(jupyter_script))
                pythonpath.insert(0, os.path.dirname(jupyter_script))
                cmd.extend([python, "-m", module])
                _log.debug("Using Jupyter script '%s'" % jupyter_script)

        if not cmd:
            python_cmd = _get_jupyter_python_script(subcommand=subcommand)
            if python_cmd:
                cmd.extend([python, "-c", python_cmd])
                _log.debug("Using Jupyter command '%s'" % python_cmd)

    if not cmd:
        jupyter_cmd = _find_jupyter_cmd(subcommand=subcommand)
        if not jupyter_cmd:
            raise RuntimeError("jupyter-notebook command not found")
        cmd.append(jupyter_cmd)
        _log.debug("Using Jupyter command '%s'" % jupyter_cmd)

    # Use the current python path when launching
    env = dict(os.environ)
    env["PYTHONPATH"] = ";".join(pythonpath)

    # Workaround for bug in win32api
    # https://github.com/mhammond/pywin32/issues/1399
    if pywintypes is not None:
        path = env.get("PATH", "").split(";")
        path.insert(0, os.path.dirname(pywintypes.__file__))
        env["PATH"] = ";".join(path)

    # Set PYXLL_IPYTHON_CONNECTION_FILE so the manager knows what to connect to
    env["PYXLL_IPYTHON_CONNECTION_FILE"] = connection_file

    if no_browser:
        cmd.append("--no-browser")

    cmd.extend(_get_jupyter_args(subcommand))
    cmd.append("-y")

    # run jupyter in it's own process
    si = subprocess.STARTUPINFO()
    si.wShowWindow |= subprocess.SW_HIDE

    # Run the exe directly if possible without a cmd child process.
    # This avoids a permission error when trying to run 'cmd' as a child process.
    shell = True
    popen_kwargs = {}
    exe = cmd[0]
    if "." in exe and exe.rsplit(".", 1)[-1].lower() == "exe":
        shell = False
        popen_kwargs["executable"] = exe
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW | 1
        cmd = [os.path.basename(exe)] + cmd[1:]

    if initial_path:
        _log.debug(f"Starting Jupyter in '{initial_path}'.")

    proc = subprocess.Popen(cmd,
                            cwd=initial_path,
                            env=env,
                            shell=shell,
                            stdout=subprocess.PIPE,
                            stderr=subprocess.STDOUT,
                            startupinfo=si,
                            **popen_kwargs)

    if proc.poll() is not None:
        raise Exception("Command '%s' failed to start" % " ".join(cmd))

    # Add it to the list of processes to be killed when Excel exits
    _all_jupyter_processes.append(proc)

    # Monitor the output of the process in a background thread
    def thread_func(proc, url_queue, killed_event):
        encoding = sys.getfilesystemencoding()
        next_line_is_url = False
        matched_url = None

        while proc.poll() is None:
            line = proc.stdout.readline().decode(encoding, "replace").strip()
            if not line:
                continue
            if line.startswith("DEBUG"):
                _log.debug(line)
                continue
            _log.info(line)

            if matched_url is None:
                if next_line_is_url:
                    match = re.search(r"(?:^|\s)(https?://.*)$", line, re.IGNORECASE)
                    if match:
                        matched_url = match.group(1).strip()
                        _log.info("Found Jupyter notebook server running on '%s'" % matched_url)
                        next_line_is_url = False
                        url_queue.put(matched_url)
                        continue

                if re.search(r"(^|\s)Jupyter (.+) is running at:", line, re.IGNORECASE):
                    next_line_is_url = True
                    continue

        if matched_url is None and not killed_event.is_set():
            _log.error("Jupyter process ended without printing a URL.")
            url_queue.put(None)

    url_queue = queue.Queue()
    killed_event = threading.Event()
    thread = threading.Thread(target=thread_func, args=(proc, url_queue, killed_event))
    thread.daemon = True
    thread.start()

    # Wait for the URL to be logged
    try:
        url = url_queue.get(timeout=timeout)
    except queue.Empty:
        _log.error("Timed-out waiting for the Jupyter notebook URL.")
        url = None

    if url is None:
        if proc.poll() is None:
            _log.debug("Killing Jupyter notebook process...")
            killed_event.set()
            kill_process(proc)
            _all_jupyter_processes.remove(proc)

        if thread.is_alive():
            _log.debug("Waiting for background thread to complete...")
            thread.join(timeout=1)
            if thread.is_alive():
                _log.warning("Timed out waiting for background thread.")

        raise RuntimeError("Timed-out waiting for the Jupyter notebook URL.")

    root, params = url.split("?", 1) if "?" in url else (url, "")
    params = params.split("&")
    params.extend(_subcommand_query_params[subcommand])

    # Update the URL to point to the notebook
    if notebook is not None:
        root = root.rstrip("/") + "/notebooks/" + notebook

    # Return the proc and url
    url = root + (("?" + "&".join(params)) if params else "")
    return proc, url


def kill_process(proc):
    """Kill a process and its children.

    :param proc: Popen process object
    """
    if proc.poll() is not None:
        return

    # Use CreateToolhelp32Snapshot to find child processes
    _TH32CS_SNAPPROCESS = 0x00000002
    _PROCESS_TERMINATE = 0x0001

    class _PROCESSENTRY32(ctypes.Structure):
        _fields_ = [("dwSize", ctypes.c_ulong),
                    ("cntUsage", ctypes.c_ulong),
                    ("th32ProcessID", ctypes.c_ulong),
                    ("th32DefaultHeapID", ctypes.c_void_p),
                    ("th32ModuleID", ctypes.c_ulong),
                    ("cntThreads", ctypes.c_ulong),
                    ("th32ParentProcessID", ctypes.c_ulong),
                    ("pcPriClassBase", ctypes.c_ulong),
                    ("dwFlags", ctypes.c_ulong),
                    ("szExeFile", ctypes.c_char * 260)]

    kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)
    CreateToolhelp32Snapshot = kernel32.CreateToolhelp32Snapshot
    Process32First = kernel32.Process32First
    Process32Next = kernel32.Process32Next
    OpenProcess = kernel32.OpenProcess
    TerminateProcess = kernel32.TerminateProcess
    CloseHandle = kernel32.CloseHandle

    def get_running_procs():
        """Return return a dict of ppid -> set of child pids for children of the current process.
        If a process has no children there is an empty set in the dict.
        """
        processes = {}
        snapshot = CreateToolhelp32Snapshot(_TH32CS_SNAPPROCESS, 0)
        try:
            entry = _PROCESSENTRY32()
            entry.dwSize = ctypes.sizeof(_PROCESSENTRY32)
            if not Process32First(snapshot, ctypes.byref(entry)):
                raise OSError('Process32First failed with error code %d' % ctypes.get_last_error())
            while True:
                # Make sure each process has an entry in the dict and add this process to its parent
                processes.setdefault(entry.th32ProcessID, set())
                processes.setdefault(entry.th32ParentProcessID, set()).add(entry.th32ProcessID)
                if not Process32Next(snapshot, ctypes.byref(entry)):
                    break
        finally:
            CloseHandle(snapshot)
        return processes

    def kill_proc_tree(pid):
        running_procs = get_running_procs()
        if pid not in running_procs:
            return

        # Terminate the process
        proc = OpenProcess(_PROCESS_TERMINATE, False, pid)
        if not proc:
            raise OSError("OpenProcess failed with error code %d" % ctypes.get_last_error())
        try:
            if 0 == TerminateProcess(proc, -9):
                raise OSError("TerminateProcess failed with error code %d" % ctypes.get_last_error())
        finally:
            CloseHandle(proc)

        # Then terminate any remaining child processes
        children = running_procs[pid]
        if children:
            # Check which are still running after terminating the parent process
            running_procs = get_running_procs()
            children = [c for c in children if c in running_procs]
            for child_pid in children:
                kill_proc_tree(child_pid)

    try:
        kill_proc_tree(proc.pid)
    except:
        _log.warning("Failed to kill Jupyter process %d" % proc.pid, exc_info=True)


@atexit.register
def _kill_jupyter_processes():
    """Ensure all Jupyter processes are killed."""
    global _all_jupyter_processes
    for proc in _all_jupyter_processes:
        kill_process(proc)
    _all_jupyter_processes = [x for x in _all_jupyter_processes if x.poll() is None]
