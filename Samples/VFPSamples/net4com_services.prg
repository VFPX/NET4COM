CLEAR
LOCAL x as NET4COM.IServices
x=CREATEOBJECT("NET4COM.Services")
? "Stopping MSMQ service..."
x.StopService("MSMQ")
INKEY(1)
? "Starting MSMQ service..."
x.StartService("MSMQ")
?? "Complete."
RETURN
