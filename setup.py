from setuptools import setup, find_packages
from importlib.util import  module_from_spec, spec_from_file_location

spec = spec_from_file_location("constants", "./pyxlsx/_constants.py")
constants = module_from_spec(spec)
spec.loader.exec_module(constants)

with open('README.md', 'r') as fp:
    long_description = fp.read()

__author__ = constants.__author__
__url__ = constants.__url__
__version__ = constants.__version__
__license__ = constants.__license__

setup(
    name='pyxlsx',
    packages=find_packages(
        # where='pyxlsx',
        exclude=['*.tests', '*.test_.*', 'tests', 'develop']
    ),
    package_dir={},
    # metadata
    author=__author__,
    long_description=long_description,
    long_description_content_type='text/markdown',
    url=__url__,
    version=__version__,
    license=__license__,
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.5',
    install_requires=[
        "openpyxl>=3.0.3",
        "pycel"
    ]
)