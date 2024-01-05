import sys
import os
import importlib.util

script_module = None
try:
    import fusionscript as script_module
except ImportError:
    # Look for installer based environment variables:
    lib_path = os.getenv("RESOLVE_SCRIPT_LIB")
    if lib_path:
        try:
            spec = importlib.util.spec_from_file_location("fusionscript", lib_path)
            script_module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(script_module)
        except ImportError:
            pass
    if not script_module:
        # Look for default install locations:
        ext = ".so"
        if sys.platform.startswith("darwin"):
            path = "/Applications/DaVinci Resolve/DaVinci Resolve.app/Contents/Libraries/Fusion/"
        elif sys.platform.startswith("win") or sys.platform.startswith("cygwin"):
            ext = ".dll"
            path = "C:\\Program Files\\Blackmagic Design\\DaVinci Resolve\\"
        elif sys.platform.startswith("linux"):
            path = "/opt/resolve/libs/Fusion/"

        try:
            spec = importlib.util.spec_from_file_location("fusionscript", path + "fusionscript" + ext)
            script_module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(script_module)
        except ImportError:
            pass

if script_module:
    sys.modules[__name__] = script_module
else:
    raise ImportError("Could not locate module dependencies")
