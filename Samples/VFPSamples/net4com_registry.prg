CLEAR
LOCAL x as NET4COM.IRegistry
x=CREATEOBJECT("NET4COM.Registry")
? "Current Setting for Color Schemes in the Registry: "+ ;
		x.GetValue("HKEY_CURRENT_USER\Control Panel\Current","Color Schemes")
x.SetValue("HKEY_CURRENT_USER\Software\MyApp","Name","Test Value")
? "Current Setting App Name value: " + x.GetValue("HKEY_CURRENT_USER\Software\MyApp","Name")
RETURN
