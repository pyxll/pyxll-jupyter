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
from zmq.eventloop import ioloop
from pyxll import schedule_call
import importlib.util
import subprocess
import logging
import threading
import queue
import atexit
import sys
import os
import re

_log = logging.getLogger(__name__)
_all_jupyter_processes = []

try:
    import win32api
except ImportError:
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
        self.loop = ioloop.IOLoop.current()

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

    # call the API embed function, which will use the monkey-patched method above
    embed_kernel(local_ns={})

    ipy = IPKernelApp.instance()

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


def _find_jupyter_script():
    """Returns the path to 'jupyter-notebook-script.py' used to start
    the Jupyter notebook server. Returns None if the script can't be found.
    """
    # Look for it using importlib first
    spec = importlib.util.find_spec("jupyter-notebook-script")
    if spec is not None and spec.origin and os.path.exists(spec.origin):
        return os.path.abspath(spec.origin)

    # If that doesn't work look in the Scripts folder
    if sys.executable and os.path.basename(sys.executable).lower() in ("python.exe", "pythonw.exe"):
        path = os.path.join(os.path.dirname(sys.executable), "Scripts", "jupyter-notebook-script.py")
        if os.path.exists(path):
            return os.path.abspath(path)

    return None


def _find_jupyter_cmd():
    """Find the 'jupyter-notebook' executable or bat file.
    Returns None if it can't be found.
    """
    # Look in the python folder and in the scripts folder
    if sys.executable and os.path.basename(sys.executable).lower() in ("python.exe", "pythonw.exe"):
        for ext in (".exe", ".bat"):
            for path in (os.path.dirname(sys.executable), os.path.join(os.path.dirname(sys.executable), "Scripts")):
                jupyter_cmd = os.path.join(path, "jupyter-notebook" + ext)
                if os.path.exists(jupyter_cmd):
                    return os.path.abspath(jupyter_cmd)

    # If it wasn't found look for it on the system path
    for ext in (".exe", ".bat"):
        jupyter_cmd = _which("jupyter-notebook" + ext)
        if jupyter_cmd is not None and os.path.exists(jupyter_cmd):
            return os.path.abspath(jupyter_cmd)

    return None


def launch_jupyter(connection_file, cwd=None, timeout=30):
    """Launch a Jupyter notebook server as a child process.

    :param connection_file: File for kernels to use to connect to an existing kernel.
    :param cwd: Current working directory to start the notebook in.
    :param timeout: Timeout in seconds to wait for the Jupyter process to start.
    :return: (Popen2 instance, URL string)
    """
    cmd = []
    pythonpath = list(sys.path)

    if sys.executable and os.path.basename(sys.executable).lower() in ("python.exe", "pythonw.exe"):
        python = os.path.join(os.path.dirname(sys.executable), "python.exe")
        if os.path.exists(python):
            jupyter_script = _find_jupyter_script()
            if jupyter_script:
                module, _ = os.path.splitext(os.path.basename(jupyter_script))
                pythonpath.insert(0, os.path.dirname(jupyter_script))
                cmd.extend([python, "-m", module])
                _log.debug("Using Jupyter script '%s'" % jupyter_script)

    if not cmd:
        jupyter_cmd = _find_jupyter_cmd()
        if not jupyter_cmd:
            raise RuntimeError("jupyter-notebook command not found")
        cmd.append(jupyter_cmd)
        _log.debug("Using Jupyter command '%s'" % jupyter_cmd)

    # Use the current python path when launching
    env = dict(os.environ)
    env["PYTHONPATH"] = ";".join(pythonpath)

    # Set PYXLL_IPYTHON_CONNECTION_FILE so the manager knows what to connect to
    env["PYXLL_IPYTHON_CONNECTION_FILE"] = connection_file

    # run jupyter in it's own process
    cmd.extend([
        "--NotebookApp.kernel_manager_class=pyxll_jupyter.extipy.ExternalIPythonKernelManager",
        "--no-browser",
        "-y"
    ])

    si = subprocess.STARTUPINFO()
    si.wShowWindow=subprocess.SW_HIDE
    proc = subprocess.Popen(cmd,
                            cwd=cwd,
                            env=env,
                            shell=True,
                            stdout=subprocess.PIPE,
                            stderr=subprocess.STDOUT,
                            startupinfo=si)

    if proc.poll() is not None:
        raise Exception("Command '%s' failed to start" % " ".join(cmd))

    # Add it to the list of processes to be killed when Excel exits
    _all_jupyter_processes.append(proc)

    # Monitor the output of the process in a background thread
    def thread_func(proc, url_queue, killed_event):
        encoding = sys.getfilesystemencoding()
        matched_url = None

        while proc.poll() is None:
            line = proc.stdout.readline().decode(encoding, "replace").strip()
            if line.startswith("DEBUG"):
                _log.debug(line)
                continue
            _log.info(line)
            if matched_url is None:
                match = re.search(r"(https?://([a-z|0-9]+\.?)+(:[0-9]+)?/?\?token=[a-f|0-9]+)", line, re.I | re.A)
                if match:
                    matched_url = match.group(1)
                    _log.info("Found Jupyter notebook server running on '%s'" % matched_url)
                    url_queue.put(matched_url)

        if matched_url is None and not killed_event.is_set():
            _log.error("Jupyter notebook process ended without printing a URL.")
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
            _kill_process(proc)
            _all_jupyter_processes.remove(proc)

        if thread.is_alive():
            _log.debug("Waiting for background thread to complete...")
            thread.join(timeout=1)
            if thread.is_alive():
                _log.warning("Timed out waiting for background thread.")

        raise RuntimeError("Timed-out waiting for the Jupyter notebook URL.")

    # Return the proc and url
    return proc, url


def _kill_process(proc):
    """Kill a process using 'taskkill /F /T'."""
    if proc.poll() is not None:
        return

    si = subprocess.STARTUPINFO(wShowWindow=subprocess.SW_HIDE)
    retcode = subprocess.call(['taskkill', '/F', '/T', '/PID', str(proc.pid)],
                              startupinfo=si,
                              shell=True)
    if proc.poll() is None:
        _log.warning("Failed to kill Jupyter process %d: %s" % (proc.pid, retcode))


@atexit.register
def _kill_jupyter_processes():
    """Ensure all Jupyter processes are killed."""
    global _all_jupyter_processes
    for proc in _all_jupyter_processes:
        _kill_process(proc)
    _all_jupyter_processes = [x for x in _all_jupyter_processes if x.poll() is None]
