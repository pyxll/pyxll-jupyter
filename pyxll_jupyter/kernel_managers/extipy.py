"""
Kernel manager for connecting to a IPython kernel started outside of Jupyter.
Use this kernel manager if you want to connect a Jupyter notebook to a IPython
kernel started outside of Jupyter.

This is for notebook versions that do not have the KernelProvisionerFactory option
and so need to patch the kernel in the mananger to connect to the existing kernel.

Most Jupyter configurations should use the kernel provisioner factory option
instead of this manager.
"""
import os
import os.path

from jupyter_client.multikernelmanager import MultiKernelManager
from notebook.services.kernels.kernelmanager import MappingKernelManager


import logging
logging.basicConfig(level=logging.DEBUG)
_log = logging.getLogger(__name__)


class ExternalMappingKernelManager(MappingKernelManager):
    """A Kernel manager that connects to a IPython kernel started outside of Jupyter"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.pinned_superclass = ExternaMultiKernelManager
        self.pinned_superclass.__init__(self, *args, **kwargs)

    def _attach_to_pyxll_kernel(self, kernel_id):
        """Attach to the externally started IPython kernel
        """
        self.log.info(f'Attaching {kernel_id} to an existing kernel...')
        kernel = self.get_kernel(kernel_id)
        port_names = ['shell_port', 'stdin_port', 'iopub_port', 'hb_port', 'control_port']
        port_names = kernel._random_port_names if getattr(kernel, '_random_port_names', None) else port_names
        for port_name in port_names:
            setattr(kernel, port_name, 0)

        # Connect to kernel started by PyXLL
        connection_file = os.environ["PYXLL_IPYTHON_CONNECTION_FILE"]
        if not os.path.abspath(connection_file):
            connection_dir = os.path.join(os.environ["APPDATA"], "jupyter", "runtime")
            connection_file = os.path.join(connection_dir, connection_file)

        if not os.path.exists(connection_file):
            _log.warning(f"Jupyter connection file '{connection_file}' does not exist.")

        _log.info(f'PyXLL IPython kernel = {connection_file}')
        kernel.load_connection_file(connection_file)

    async def start_kernel(self, **kwargs):
        """Attach to the kernel started by PyXLL.
        """
        kernel_id = await super(ExternalMappingKernelManager, self).start_kernel(**kwargs)
        self._attach_to_pyxll_kernel(kernel_id)
        return kernel_id



class ExternaMultiKernelManager(MultiKernelManager):
    """Subclass of MultiKernelManager to prevent restarting"""    

    def restart_kernel(self, *args, **kwargs):
        raise NotImplementedError("Restarting a kernel running in Excel is not supported.")
    
    async def _async_restart_kernel(self, *args, **kwargs):
        raise NotImplementedError("Restarting a kernel running in Excel is not supported.")

    def shutdown_kernel(self, *args, **kwargs):
        raise NotImplementedError("Shutting down a kernel running in Excel is not supported.")

    async def _async_shutdown_kernel(self, *args, **kwargs):
        raise NotImplementedError("Shutting down a kernel running in Excel is not supported.")

    def shutdown_all(self, *args, **kwargs):
        raise NotImplementedError("Shutting down a kernel running in Excel is not supported.")

    async def _async_shutdown_all(self, *args, **kwargs):
        raise NotImplementedError("Shutting down a kernel running in Excel is not supported.")
