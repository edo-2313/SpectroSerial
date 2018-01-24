from cx_Freeze import setup, Executable

base = None
executables = [Executable("main.py", base=base,icon="icon.ico")]

packages = ["serial","openpyxl","os","re","datetime"]
options = {
    'build_exe': {

        'packages':packages,
    },

}

setup(
    name = "Spettrofotometro Excel",
    options = options,
    version = "1.0",
    description = 'Spettrofotometro seriale',
    executables = executables
)
