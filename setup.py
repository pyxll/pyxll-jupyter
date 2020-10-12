from setuptools import setup, find_packages

setup(
    name="pyxll_jupyter",
    description="Adds Jupyter notebooks to Microsoft Excel using PyXLL.",
    version="0.1.3",
    packages=find_packages(),
    include_package_data=True,
    package_data={
        "pyxll_jupyter": [
            "pyxll_jupyter/resources/ribbon.xml",
            "pyxll_jupyter/resources/jupyter.png",
        ]
    },
    entry_points={
        "pyxll": [
            "modules = pyxll_jupyter.pyxll:modules",
            "ribbon = pyxll_jupyter.pyxll:ribbon"
        ]
    },
    install_requires=[
        #"pyxll >= 5.0.0",
        "jupyter >= 1.0.0",
        "PySide2",
        "pywin32"
    ]
)
