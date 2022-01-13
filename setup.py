"""
PyXLL-Jupyter

This package integrated Jupyter notebooks into Microsoft Excel.

To install it, first install PyXLL (see https://www.pyxll.com).

Briefly, to install PyXLL do the following::

    pip install pyxll
    pyxll install

Once PyXLL is installed then installing this package will add a
button to the PyXLL ribbon toolbar that will start a Jupyter
notebook browser as a custom task pane in Excel.

To install this package use::

    pip install pyxll_jupyter
"""
from setuptools import setup, find_packages
from os import path


this_directory = path.abspath(path.dirname(__file__))
with open(path.join(this_directory, 'README.md'), encoding='utf-8') as f:
    long_description = f.read()


setup(
    name="pyxll_jupyter",
    description="Adds Jupyter notebooks to Microsoft Excel using PyXLL.",
    long_description=long_description,
    long_description_content_type='text/markdown',
    version="0.3.0",
    packages=find_packages(),
    include_package_data=True,
    package_data={
        "pyxll_jupyter": [
            "pyxll_jupyter/resources/ribbon.xml",
            "pyxll_jupyter/resources/jupyter.png",
        ]
    },
    project_urls={
        "Source": "https://github.com/pyxll/pyxll-jupyter",
        "Tracker": "https://github.com/pyxll/pyxll-jupyter/issues",
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows"
    ],
    entry_points={
        "pyxll": [
            "modules = pyxll_jupyter.pyxll:modules",
            "ribbon = pyxll_jupyter.pyxll:ribbon"
        ],
        "jupyter_client.kernel_provisioners": [
            "pyxll-provisioner = pyxll_jupyter.provisioning:ExistingProvisioner"
        ]
    },
    install_requires=[
        "pyxll >= 5.1.0",
        "jupyter >= 1.0.0",
        "notebook >= 6.0.0",
        "PySide2;python_version<'3.10'",
        "PySide6;python_version>='3.10'"
    ]
)
