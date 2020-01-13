Attribute VB_Name = "Module9"
Function AddDate(X As Range)
If X.Value <> "" Then
    AddDate = Format(Now, "dd-mmm-yy hh:mm:ss")
    Else
        AddDate = ""
End If
End Function
