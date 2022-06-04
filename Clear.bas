Attribute VB_Name = "Module2"
'Main - clear "Raw" worksheet
Sub Clear_Raw()
    
    Application.ScreenUpdating = False
    
    Worksheets("Raw").Range("20:" & Rows.Count).ClearContents
    
    Application.ScreenUpdating = True
    
End Sub
