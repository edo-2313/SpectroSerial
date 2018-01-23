from cx_Freeze import setup, Executable

base = None
executables = [Executable("main.py", base=base)]

packages = ["serial","openpyxl","os","re","datetime"]
options = {
    'build_exe': {

        'packages':packages,
    },

}

setup(
    name = "SpectroSerial",
    options = options,
    version = "0.1",
    description = 'Spettrofotometro seriale',
    executables = executables
)
