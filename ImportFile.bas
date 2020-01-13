Attribute VB_Name = "Module1"
Sub ImportFile_WMS()
    Dim xRg As Range
    Dim PilihItem As Variant
    Dim PilihFile As FileDialog
    Dim NamaFile, NamaSheet, RangeData As String
    Dim FileSumber, CekSheetHasil As Workbook
    Dim ImportKe As Worksheet

    On Error Resume Next
    With Application
        .DisplayAlerts = False
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    NamaSheet = "Loadinglist"
    RangeData = "A2:AC1000"
    Set PilihFile = Application.FileDialog(msoFileDialogFolderPicker)
    With PilihFile
        If .Show = -1 Then
            PilihItem = .SelectedItems.Item(1)
            Set CekSheetHasil = ThisWorkbook
            Set ImportKe = CekSheetHasil.Sheets("Hasil Import WMS")
            If ImportKe Is Nothing Then
                CekSheetHasil.Sheets.Add(after:=CekSheetHasil.Worksheets(CekSheetHasil.Worksheets.Count)).Name = "Hasil Import"
                Set ImportKe = CekSheetHasil.Sheets("Hasil Import WMS")
            End If
            NamaFile = Dir(PilihItem & "\*.xls", vbNormal)
            If NamaFile = "" Then Exit Sub
            Do Until NamaFile = ""
               Set FileSumber = Workbooks.Open(PilihItem & "\" & NamaFile)
                Set xRg = FileSumber.Worksheets(NamaSheet).Range(RangeData)
                xRg.Copy ImportKe.Range("A65536").End(xlUp).Offset(1, 0)
                NamaFile = Dir()
                FileSumber.Close
            Loop
        End If
    End With
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
MsgBox "Data Berhasil di Import", vbInformation, "Informasi"
End Sub

Sub HapusData_WMS()
Dim Yakin As Integer
Dim Ws As Worksheet
Set Ws = Worksheets("Hasil Import WMS")
Yakin = MsgBox("Apakah Anda akan menghapus semua hasil import?", vbOKCancel, "Verifikasi")
If Yakin = 1 Then
Ws.Range("A5:AC65536").Clear
End If
End Sub

