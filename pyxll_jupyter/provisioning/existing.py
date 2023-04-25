from jupyter_client import KernelProvisionerBase
import logging
import json
import os

_log = logging.getLogger(__name__)


class ExistingProvisioner(KernelProvisionerBase):
    """
    A Kernel Provisioner that re-uses an existing kernel.
    The kernel connection file is set in the environment variable
    'PYXLL_IPYTHON_CONNECTION_FILE'.
    """

    async def launch_kernel(self, cmd, **kwargs):
        # Connect to kernel started by PyXLL
        connection_file = os.environ["PYXLL_IPYTHON_CONNECTION_FILE"]
        if not os.path.abspath(connection_file):
            connection_dir = os.path.join(os.environ["APPDATA"], "jupyter", "runtime")
            connection_file = os.path.join(connection_dir, connection_file)

        if not os.path.exists(connection_file):
            _log.warning(f"Jupyter connection file '{connection_file}' does not exist.")

        _log.info(f'PyXLL IPython kernel = {connection_file}')
        with open(connection_file) as f:
            file_info = json.load(f)

        file_info["key"] = file_info["key"].encode()
        return file_info

    async def pre_launch(self, **kwargs):
        kwargs = await super().pre_launch(**kwargs)
        kwargs.setdefault('cmd', None)
        return kwargs

    def has_process(self) -> bool:
        return True

    async def poll(self):
        pass

    async def wait(self):
        pass

    async def send_signal(self, signum: int):
        pass

    async def kill(self, restart=False):
        if restart:
            _log.warning("Cannot restart kernel running in Excel.")

    async def terminate(self, restart=False):
        if restart:
            _log.warning("Cannot restart kernel running in Excel.")

    async def cleanup(self, restart):
        pass
