Attribute VB_Name = "Module2"
Sub CATMain()

 

CATIA.DisplayFileAlerts = True

path = CATIA.ActiveDocument.path

Name = CATIA.ActiveDocument.Name

initial = path & "\" & Name

Set Folder = CATIA.FileSystem.CreateFolder("oTemp")

oFolder = path & "\oTemp"

 

Set Send = CATIA.CreateSendTo()

Call Send.SetInitialFile(initial)

Send.SetDirectoryFile (oFolder)

Send.Run

 

End Sub
 

