from cx_Freeze import setup, Executable

base = None
executables = [Executable("filetoserial.py", base=base)]

packages = ["serial","openpyxl","os","re","datetime"]
options = {
    'build_exe': {

        'packages':packages,
    },

}

setup(
    name = "File2serial",
    options = options,
    version = "1.0",
    description = 'File a seriale',
    executables = executables
)
