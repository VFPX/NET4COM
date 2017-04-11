CLEAR
LOCAL x as NET4COM.IMouse
x=CREATEOBJECT("NET4COM.Mouse")
? IIF(x.ButtonsSwapped(),"Buttons are swapped","Buttons are NOT swapped")
? IIF(x.WheelExists(),"Wheel exists","Wheel does NOT exist")
? "One notch of mouse will scroll "+ALLTRIM(STR(x.WheelScrollLines()))+" lines"
RETURN
