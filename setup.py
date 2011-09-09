from distutils.core import setup
import py2exe
import sys
 
if len(sys.argv) == 1:
    sys.argv.append("py2exe")

# creates a standalone .exe file, no zip files
setup(options = {"py2exe": {"compressed": 1, "optimize": 2, "ascii": 1, "bundle_files": 1}},
    zipfile = None,
    console = [{"script": 'imgstitcher.py'}] )
