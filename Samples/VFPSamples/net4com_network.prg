CLEAR
LOCAL x AS NET4COM.INetwork
x=CREATEOBJECT("NET4COM.Network")
? IIF(x.IsAvailable(),"Network is available","Network is NOT available")
? IIF(x.Ping("localhost"),"Ping successful","Ping UNseccessful")
? "Downloading file..."
x.DownloadFile("http://foxcentral.net","_output.txt")
?? "Complete."
MODIFY FILE _output.txt
ERASE _output.txt
RETURN
