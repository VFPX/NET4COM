CLEAR
LOCAL cNET4COMClassLib
cNET4COMClassLib=HOME()+"\ffc\_NET4COM.vcx"

* Use NET4COM.Mouse COM object directly, IntelliSense via NET4COM.Mouse COM interface.
LOCAL oNET4COM_Mouse as NET4COM.IMouse
oNET4COM_Mouse=CREATEOBJECT("NET4COM.Mouse")
? oNET4COM_Mouse.WheelExists()

* Use FFC wrapper class to access NET4COM.Mouse, IntelliSense via FFC Custom based class.
LOCAL oMouse AS _NET4COM_Mouse OF (HOME()+"\ffc\_NET4COM.vcx")
oMouse=NEWOBJECT("_NET4COM_Mouse",cNET4COMClassLib)
? oMouse.WheelExists()

* Use FFC wrapper class to access NET4COM.Mouse, IntelliSense via NET4COM.Mouse COM interface.
LOCAL oMouse2 AS NET4COM.IMouse
oMouse2=NEWOBJECT("_NET4COM_Mouse",cNET4COMClassLib)
? oMouse2.WheelExists()

* Use dynamic FFC class via this_access to access methods in NET4COM classes, no IntelliSense.
LOCAL oNET4COM
oNET4COM=NEWOBJECT("_NET4COM",cNET4COMClassLib)
? oNET4COM.Mouse.WheelExists()

LOCAL i,n,s,s1
n=10000
? "Loops: "+STR(n)

s1=SECONDS()
FOR i = 1 TO n
	=oNET4COM_Mouse.WheelExists()
ENDFOR
s=SECONDS()-s1
? "NET4COM.Mouse via COM: "+STR(s,9,4)

s1=SECONDS()
FOR i = 1 TO n
	=oMouse.WheelExists()
ENDFOR
s=SECONDS()-s1
? "_Mouse in NET4COM.vcx: "+STR(s,9,4)

s1=SECONDS()
FOR i = 1 TO n
	=oNET4COM.Mouse.WheelExists()
ENDFOR
s=SECONDS()-s1
? "_NET4COM in NET4COM.vcx: "+STR(s,9,4)

RETURN
