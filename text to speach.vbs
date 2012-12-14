Dim message, sapi
Set sapi=CreateObject("sapi.spvoice")
message=InputBox("What do you want me to say?","Speak to Me")
sapi.Speak message