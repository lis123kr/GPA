"""
 Copyright [2017] [Il Seop Lee / Chungbuk National University]
 
 Licensed under the Apache License, Version 2.0 (the "License");
 you may not use this file except in compliance with the License.
 You may obtain a copy of the License at
 
 http://www.apache.org/licenses/LICENSE-2.0
 
 Unless required by applicable law or agreed to in writing, software
 distributed under the License is distributed on an "AS IS" BASIS,
 WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 See the License for the specific language governing permissions and
 limitations under the License.
 """
import cx_Freeze, os, sys, matplotlib
import PyQt5
os.environ['TCL_LIBRARY'] = r'C:\Users\lis12\Anaconda3\envs\python\tcl\tcl8.6'
os.environ['TK_LIBRARY'] = r'C:\Users\lis12\Anaconda3\envs\python\tcl\tk8.6'

# exe = [ cx_Freeze.Executable("GPA_UI.py", base="Console") ]
PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))
build_exe_options = {
					"packages": [
						"xlrd",
						"pandas", 
						"numpy",
						"openpyxl", 
						"matplotlib"
						], 
					"includes":   ["PyQt5.QtCore", "PyQt5.QtGui", "PyQt5.QtWidgets"],
                            # "tkinter.filedialog","numpy"],
                     "include_files":[(matplotlib.get_data_path(), "mpl-data"), "libEGL.dll", "qwindows.dll"],
                     "excludes":[],
                    }
base = None

if sys.platform == "win32":
    base = "Win32GUI"
    DLLS_FOLDER = os.path.join(PYTHON_INSTALL_DIR, 'Library', 'bin')

    dependencies = ['libiomp5md.dll', 'mkl_core.dll', 'mkl_def.dll', 'mkl_intel_thread.dll']

    for dependency in dependencies:
        build_exe_options['include_files'].append(os.path.join(DLLS_FOLDER, dependency))

executables = [cx_Freeze.Executable("GPA_UI.py", base=base)]

cx_Freeze.setup( 
	name = "test", 
	version = "1.0", 
	options = {
		"build_exe": build_exe_options
	}, 
	executables = executables 
) 