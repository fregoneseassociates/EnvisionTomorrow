'''
Locate ArcPy and add it to the path
Original author: Jamesramm (https://github.com/JamesRamm/archook)
Created on 13 Feb 2015
Modified version author: Thomas York (sustainablegrowth at gmail dot com)
Date: May 26 2016
Modifications summary:
    Locate ArcGIS python installation and add its site-packages directory to path
    Programatically protects against common problem that results in failure to import arcpy
    http://mattmakesmaps.com/blog/2013/07/10/fixing-arcgis-10-dot-1-python-console-numpy-import-error/
License:
    Licensed under the Apache License, Version 2.0 (the "License"); you may not use
    this project except in compliance with the License. You may obtain a copy of the
    License at

       http://www.apache.org/licenses/LICENSE-2.0

    Unless required by applicable law or agreed to in writing, software distributed
    under the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR
    CONDITIONS OF ANY KIND, either express or implied. See the License for the
    specific language governing permissions and limitations under the License.
'''
import _winreg
import sys
from os import path

def arcgis_version():
    try:
        # get version
        key = _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE, 'SOFTWARE\\Wow6432Node\\ESRI\\ArcGIS', 0)
        version = _winreg.QueryValueEx(key, "RealVersion")[0][:4]

        return version
    except WindowsError:
        raise ImportError("Could not locate the ArcGIS version on this machine")


def arcgis_program_dir(version):
    '''
    Find the path to the ArcGIS Desktop installation.

    Keys to check:

    HLKM/SOFTWARE/ESRI/ArcGIS 'RealVersion' - will give the version, then we can use
    that to go to
    HKLM/SOFTWARE/ESRI/DesktopXX.X 'InstallDir'. Where XX.X is the version

    We may need to check HKLM/SOFTWARE/Wow6432Node/ESRI instead
    '''
    try:
        # get location of InstallDir
        key_string = "SOFTWARE\\Wow6432Node\\ESRI\\Desktop{0}".format(version)
        desktop_key = _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE, key_string, 0)
        install_dir = _winreg.QueryValueEx(desktop_key, "InstallDir")[0]

        return install_dir
    except WindowsError:
        raise ImportError("Could not locate the ArcGIS instllation directory on this machine")

def arcgis_python_dir(version):
    '''
    Find the path to the ArcGIS Python installation.

    Keys to check:

    HLKM/SOFTWARE/ESRI/ArcGIS 'RealVersion' - will give the version, then we can use
    that to go to
    HKLM/SOFTWARE/ESRI/PythonX.X 'PythonDir'. Where X.X is the version

    We may need to check HKLM/SOFTWARE/Wow6432Node/ESRI instead
    '''
    try:
        # get location of PythonDir
        key_string = "SOFTWARE\\Wow6432Node\\ESRI\\Python{0}".format(version)
        desktop_key = _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE, key_string, 0)
        python_dir = _winreg.QueryValueEx(desktop_key, "PythonDir")[0]

        return python_dir
    except WindowsError:
        raise ImportError("Could not locate the ArcGIS Python installation directory on this machine")


def get_arcpy_sp():
    '''
    Allows arcpy to imported on 'unmanaged' python installations (i.e. python installations
    arcgis is not aware of).
    Gets the location of arcpy and related libs and adds it to sys.path
    '''
    arcgis_ver = arcgis_version()
    install_dir = arcgis_program_dir(arcgis_ver)
    arcpy = path.join(install_dir, "arcpy")
    # Check we have the arcpy directory.
    if not path.exists(arcpy):
        raise ImportError("Could not find arcpy directory in {0}".format(install_dir))

    # First check if we have a bin64 directory - this exists when arcgis is 64bit
    bin_dir = path.join(install_dir, "bin64")
    if not path.exists(bin_dir):
        # Fall back to regular 'bin' dir otherwise.
        bin_dir = path.join(install_dir, "bin")

    scripts = path.join(install_dir, "ArcToolbox", "Scripts")


    # Next, get and set sys.path for ArcGIS python site-packages directory
    python_dir = arcgis_python_dir(arcgis_ver)
    verdir = "ArcGIS" + str(arcgis_ver)
    sp = path.join(python_dir, verdir,"Lib","site-packages")
    # Check we have the site-packages directory.
    if not path.exists(sp):
        raise ImportError("Could not find expected python site-packages directory at {0}".format(sp))

    sys.path.extend([arcpy, bin_dir, scripts, sp])

