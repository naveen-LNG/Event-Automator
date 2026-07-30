import sys
import os

with open(r"c:\SkipBo\branches\Misc\develop\client\Skipbo\Assets\Editor\EventAutomation\PythonBackend\args_debug.txt", "w") as f:
    f.write("ARGV: " + repr(sys.argv) + "\n")
    f.write("CWD: " + os.getcwd() + "\n")
    f.write("EXEC: " + sys.executable + "\n")
