"""
Kernel manager for connecting to a IPython kernel started outside of Jupyter.
Use this kernel manager if you want to connect a Jupyter notebook to a IPython
kernel started outside of Jupyter.
"""
from jupyter_client.multikernelmanager import MultiKernelManager

try:
    # Notebook < 7.0.0 has it's own copy of MappingKernelManager
    from notebook.services.kernels.kernelmanager import MappingKernelManager
except ImportError:
    # Notebook >= 7.0.0 uses the one from jupyter_server
    from jupyter_server.services.kernels.kernelmanager import MappingKernelManager


class ExternalMappingKernelManager(MappingKernelManager):
    """A Kernel manager that connects to a IPython kernel started outside of Jupyter"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.pinned_superclass = ExternaMultiKernelManager
        self.pinned_superclass.__init__(self, *args, **kwargs)


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
