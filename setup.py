from cx_Freeze import setup, Executable

base = None    

executables = [Executable("main.py", base=base)]

packages = ["idna"]
options = {
    'build_exe': {    
        'packages':packages,
    },    
}

setup(
    name = "WBSearchParser",
    options = options,
    version = "1",
    description = '<This python script can do some parsing in wildberries.ru\n For correct work you need latest version of chrome browser installed on pc and latest chromedriver.exe version in folser with .exe file>',
    executables = executables
)