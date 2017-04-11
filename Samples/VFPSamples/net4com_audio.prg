CLEAR
LOCAL x as NET4COM.IAudio
x=CREATEOBJECT("NET4COM.Audio")
? "Playing system sound..."
x.PlaySystemSound(1)
?? "Complete."
RETURN
