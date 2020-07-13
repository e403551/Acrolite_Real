import cx_Freeze

exe = [cx_Freeze.Executable(script="mk2_main_go_at_GUI.py",base="Win32GUI",icon="gui.ico")]

cx_Freeze.setup(
    name = "Acrolite",
    version = "1.1",
    description = "Customer list generator",
    author = "Nic Benedetto",
    options = {"build_exe": {"packages": ["docx2txt","docx","csv","copy","time","os"], "include_files": ["mk2_main_go_at_GUI.py","blacklistlite.csv","lite_gloss.csv","Customer_List.xlsx","logo2.png","topleft_logo.png"]}},
    executables = exe
)