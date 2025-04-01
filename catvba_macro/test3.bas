'Const catCaptureFormatPNG As Integer =6
Const catCaptureFormatJPEG As Integer =1
Const catCaptureFormatBMP As Integer =0Const catCaptureFormatTIFF As Integer=2
Set annotationSet = mCATIA-ActiveDocument .part .AnnotationSets.Item(1)Set captures = annotationSet.captures
'Loop Through CapturesForb=1To captures.CountSet captureItem = captures.Item(b)If captureItem.Name=Then
'Activate the capture
captureItem.DisplayCapture
Set viewer =mCATIA.Activeindow.ActiveviewermCATIA.RefreshDisplay =True
Set targetSheet = newWorkbook.Sheets(6)
imagePath = ThisWorkbook.path &"¡±&¡°¡±&timeStampCat &¡±.bmp¡±
viewer.CaptureToFile catCaptureFormatBMP,imagePath
DiminsertedPicture As Shape
SetinsertedPicture = targetSheet.Shapes.AddPicture(Filename:=imagePath,LinkToFile:=msoFalse, SaveWithDocument:=msoTrue,Left:=10, Top:=10,Width:=1, Height:=1)
With insertedPicture.LockAspectRatio = msoTrue.Width =300
End With
Height = 200
targetSheet.Range("A1").Select Deselect image to avoid errorstargetSheet .Activate
End If
Next b