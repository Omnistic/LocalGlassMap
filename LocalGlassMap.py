import clr, os, winreg
import matplotlib.pyplot as plt
from matplotlib.ticker import FormatStrFormatter
import numpy as np
import chardet

# This boilerplate requires the 'pythonnet' module.
# The following instructions are for installing the 'pythonnet' module via pip:
#    1. Ensure you are running a Python version compatible with PythonNET. Check the article "ZOS-API using Python.NET" or
#    "Getting started with Python" in our knowledge base for more details.
#    2. Install 'pythonnet' from pip via a command prompt (type 'cmd' from the start menu or press Windows + R and type 'cmd' then enter)
#
#        python -m pip install pythonnet

# determine the Zemax working directory
aKey = winreg.OpenKey(winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER), r"Software\Zemax", 0, winreg.KEY_READ)
zemaxData = winreg.QueryValueEx(aKey, 'ZemaxRoot')
NetHelper = os.path.join(os.sep, zemaxData[0], r'ZOS-API\Libraries\ZOSAPI_NetHelper.dll')
winreg.CloseKey(aKey)

# add the NetHelper DLL for locating the OpticStudio install folder
clr.AddReference(NetHelper)
import ZOSAPI_NetHelper

pathToInstall = ''
# uncomment the following line to use a specific instance of the ZOS-API assemblies
#pathToInstall = r'C:\C:\Program Files\Zemax OpticStudio'

# connect to OpticStudio
success = ZOSAPI_NetHelper.ZOSAPI_Initializer.Initialize(pathToInstall);

zemaxDir = ''
if success:
    zemaxDir = ZOSAPI_NetHelper.ZOSAPI_Initializer.GetZemaxDirectory();
    print('Found OpticStudio at:   %s' + zemaxDir);
else:
    raise Exception('Cannot find OpticStudio')

# load the ZOS-API assemblies
clr.AddReference(os.path.join(os.sep, zemaxDir, r'ZOSAPI.dll'))
clr.AddReference(os.path.join(os.sep, zemaxDir, r'ZOSAPI_Interfaces.dll'))
import ZOSAPI

TheConnection = ZOSAPI.ZOSAPI_Connection()
if TheConnection is None:
    raise Exception("Unable to intialize NET connection to ZOSAPI")

TheApplication = TheConnection.ConnectAsExtension(0)
if TheApplication is None:
    raise Exception("Unable to acquire ZOSAPI application")

if TheApplication.IsValidLicenseForAPI == False:
    raise Exception("License is not valid for ZOSAPI use.  Make sure you have enabled 'Programming > Interactive Extension' from the OpticStudio GUI.")

TheSystem = TheApplication.PrimarySystem
if TheSystem is None:
    raise Exception("Unable to acquire Primary system")

print('Connected to OpticStudio')

# The connection should now be ready to use.  For example:
print('Serial #: ', TheApplication.SerialCode)

# Insert Code Here

# Glass catalog folder
cat_folder = TheApplication.GlassDir

# Number of surfaces
n_sur = TheSystem.LDE.NumberOfSurfaces

# List of materials
mats = []

# List of indices of refraction
refr = []

# List of Abbe numbers
abbe = []

# List of relative costs
cost = []

# Relative cost scale factor
scal = 100

# Material found flag
mat_found = False

# Loop over the surfaces
for ii in range(n_sur):
    # Current surface data
    sur = TheSystem.LDE.GetSurfaceAt(ii)
    
    # Surface material
    mat = sur.Material
    
    # If the material isn't already in the list and isn't empty and isn't
    # a MIRROR
    if mat not in mats and mat!= '' and mat.casefold() != 'mirror': 
        # Append material to the list
        mats.append(mat)

        # Surface material catalog
        cat = sur.MaterialCatalog

        # Create a path to the catalog file
        cat_path = os.path.join(cat_folder, cat)

        # Check if catalog file exists
        if os.path.exists(cat_path):
            # Attempting to determine character encoding in the text file
            # using chardet. This was taken from:
            # https://stackoverflow.com/questions/3323770/character-detection-in-a-text-file-in-python-using-the-universal-encoding-detect
            rawdata = open(cat_path, 'rb').read()
            result = chardet.detect(rawdata)
            charenc = result['encoding']
            
            # Try to parse the catalog file for the specified material
            with open(cat_path, 'r', encoding=charenc) as glass_cat:
                for line in glass_cat:
                    # Try to find the material line
                    if line.startswith('NM ' + mat):
                        mat_found = True
                        mat_line = line
                        continue
                    
                    # Try to find the other data line (for relative cost)
                    if mat_found:
                        if line.startswith('OD '):
                            od_line = line
                            break
            
            # If the material data was found
            if mat_found:
                # Reset material found flag
                mat_found = False
                
                # Split material name line
                mat_line = mat_line.split()
                
                # Split other data line
                od_line = od_line.split()

                # Index of refraction
                refr.append(float(mat_line[4]))
                
                # Abbe number
                abbe.append(float(mat_line[5]))
                
                # Relative cost times scale factor
                cost.append(scal * float(od_line[1]))
                if cost[-1] <= 0:
                    cost[-1] = scal
        
# Plot results
plt.figure()
plt.scatter(abbe, refr, c=np.random.rand(len(abbe),3), s=cost, alpha=.7)
plt.grid()

plt.title('Current lens glass map')
plt.xlabel('Abbe number')
plt.ylabel('Index of refraction')

ax = plt.gca()
ax.yaxis.set_major_formatter(FormatStrFormatter('%.5f'))
ax.invert_xaxis()

for ii in range(len(abbe)):
    ax.annotate(mats[ii], (abbe[ii], refr[ii]))

plt.show()