import sys
from cx_Freeze import setup, Executable

build_exe_options = {
    "packages": ["os", "sys", "PyQt5", "pptx"],
    "include_files": [("assets/icon.ico", "icon.ico")],
}

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="Text to PowerPoint",
    version="1.0",
    description="Generate PowerPoint presentations from text",
    options={"build_exe": build_exe_options},
    executables=[Executable("src/app.py", base=base, icon="assets/icon.ico")]
)