CLEAR
LOCAL x as NET4COM.IClipboard
x=CREATEOBJECT("NET4COM.Clipboard")
x.SetText("Clipboard text from VFP")
? x.ContainsText()
? x.GetText()
RETURN
