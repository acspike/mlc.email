##############################################################################
# This is the distutils script for creating a Python-based com (exe or dll)
# server using win32com.  This script should be run like this:
#
#  % python setup.py py2exe
#
# After you run this (from this directory) you will find two directories here:
# "build" and "dist".  The .dll or .exe in dist is what you are looking for.
#
# http://www.py2exe.org/index.cgi/Py2exeAndWin32com
#
##############################################################################

# http://www.py2exe.org/index.cgi/win32com.shell
# ModuleFinder can't handle runtime changes to __path__, but win32com uses them
try:
    # py2exe 0.6.4 introduced a replacement modulefinder.
    # This means we have to add package paths there, not to the built-in
    # one.  If this new modulefinder gets integrated into Python, then
    # we might be able to revert this some day.
    # if this doesn't work, try import modulefinder
    try:
        import py2exe.mf as modulefinder
    except ImportError:
        import modulefinder
    import win32com, sys
    for p in win32com.__path__[1:]:
        modulefinder.AddPackagePath("win32com", p)
    for extra in ["win32com.shell"]: #,"win32com.mapi"
        __import__(extra)
        m = sys.modules[extra]
        for p in m.__path__[1:]:
            modulefinder.AddPackagePath(extra, p)
except ImportError:
    # no build path setup, no worries.
    pass


from distutils.core import setup
import py2exe
import sys

class Target:
    def __init__(self, **kw):
        self.__dict__.update(kw)
        # for the version info resources (Properties -- Version)
        self.version = "0.0.1"
        self.company_name = "Martin Luther College"
        self.copyright = "(c) 2012, Aaron C Spike"
        self.name = "MLC.Email"

my_com_server_target = Target(
    description = "a simple object providing email functionality to COM aware applications",
    # use module name for win32com exe/dll server
    modules = ["mlc_email"],
    # specify which type of com server you want (exe and/or dll)
    create_exe = True,
    create_dll = False
    )

setup(
    name="MLC.Email",
    # the following two parameters embed support files within exe/dll file
    options={"py2exe": {
        "bundle_files": 1, 
        }},
    zipfile=None,
    version="0.0.1",
    description="a simple object providing email functionality to COM aware applications",
    # author, maintainer, contact go here:
    author="Aaron Spike",
    author_email="aaron@ekips.org",
    py_modules=["mlc_email"],
    com_server=[my_com_server_target]
    )
