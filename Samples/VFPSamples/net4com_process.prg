CLEAR
LOCAL x as NET4COM.IProcess
x=CREATEOBJECT("NET4COM.Process")
x.StartApp("Notepad.exe")
INKEY(1)
x.StopApp("notepad")
RETURN
