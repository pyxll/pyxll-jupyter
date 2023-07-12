"""
Kernel manager for connecting to a IPython kernel started outside of Jupyter.
Use this kernel manager if you want to connect a Jupyterlab notebook to a IPython
kernel started outside of Jupyter.
"""
from jupyter_server.services.kernels.kernelmanager import MappingKernelManager, MultiKernelManager


class ExternalMappingKernelManager(MappingKernelManager):
    """A Kernel manager that connects to a IPython kernel started outside of Jupyter"""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.pinned_superclass = ExternalMultiKernelManager
        self.pinned_superclass.__init__(self, *args, **kwargs)


class ExternalMultiKernelManager(MultiKernelManager):
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