from cx_Freeze import setup, Executable

# Include your script file and any additional files or packages it depends on
executables = [Executable(script='main.py')]

setup(
    name='PandasAutomator',              # Replace with your app's name
    version='1.0',
    description='EXCEL Automation',
    executables=executables
)
