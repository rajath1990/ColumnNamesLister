Attribute VB_Name = "UserForm1_Code"
Option Explicit
Dim fso
Dim ts
Dim txtFileName As String
Dim file As Office.FileDialog
Dim fldr As Object
Dim strLine
Dim selection As Range
Dim NRow As Long, TargetCell As Range
Dim total As Long
Dim WB As Workbook
Dim Ws As Worksheet
Private Sub cmdCancel2_Click()
    Unload Me
End Sub

'---CODE FOR OPEN BROWSE PATH FOR SHEET1---
Private Sub cmdSheet1_Click()
 Call OpenFilePath
End Sub
'---CODE FOR OPEN BROWSE PATH FOR SHEET2---
Private Sub cmdSheet2_Click()
 Call OpenFolderPath
End Sub
Sub OpenFilePath()
    Set file = Application.FileDialog(msoFileDialogFilePicker)
   With file
      .AllowMultiSelect = False
      ' Set the title of the dialog box.
      .Title = "Please select the file."
      ' Clear out the current filters, and add our own.
      .Filters.Clear
      '.Filters.Add "Excel 2013", "*.xls"
      .Filters.Add "All Files", "*.*"
      
    If .Show = True Then
        txtFileName = .SelectedItems(1) 'replace txtFileName with your textbox
        If txtsheet1.Value <> "" Then
        txtsheet2.Value = txtFileName
        Else
        txtsheet1.Value = txtFileName
        End If
      End If
    End With
    End Sub
Sub OpenFolderPath()
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
   With fldr
      .AllowMultiSelect = False
      ' Set the title of the dialog box.
      .Title = "Please select working folder."
      ' Clear out the current filters
      .Filters.Clear
    If .Show = True Then
        txtFileName = .SelectedItems(1) 'replace txtFileName with your textbox
        If txtsheet1.Value <> "" Then
        txtsheet2.Value = txtFileName
        Else
        txtsheet1.Value = txtFileName
        End If
      End If
    End With
    
    End Sub
Sub cmdExecute_Click()

Application.ScreenUpdating = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(txtsheet1.Value)
    Do While Not ts.AtEndOfStream
         strLine = ts.ReadLine()
        'Open the file
         On Error Resume Next
         Excel.Workbooks.Open (txtsheet2.Value & "\" & strLine)
         ActiveWorkbook.CheckCompatibility = False
         Workbooks(strLine).Activate
         Worksheets(1).Activate
         Range("A1").Select
         
       If Range("B1").Value <> "" Then
          If rbvertical.Value = True Then
            Set selection = Range("A1", Range("A1").End(xlToRight))
            total = Range("A1", Range("A1").End(xlToRight)).Count
            selection.Copy
            ThisWorkbook.Sheets("Vertical List").Activate
            If Range("A2") <> "" Then
            Range("A1").Select
            ActiveCell.End(xlDown).Offset(1, 0).Select
            ActiveCell.Value = strLine
            Else
            Range("A2").Select
            ActiveCell.Value = strLine
            End If
             
            ActiveCell.Offset(0, 1).PasteSpecial , , , Transpose:=True
            NRow = Range("A" & Rows.Count).End(xlUp).Row + 1
            Set TargetCell = Range("A" & NRow - 1)
            TargetCell.Resize(total, 1).Value = strLine
            Workbooks(strLine).Close True
        
        ElseIf rbhorizontal.Value = True Then
        Set selection = Range("A1", Range("A1").End(xlToRight))
        selection.Copy
        ThisWorkbook.Sheets("Horizontal List").Activate
        If Range("A2") <> "" Then
        Range("A1").Select
        ActiveCell.End(xlDown).Offset(1, 0).Select
        ActiveCell.Value = strLine
        ActiveCell.Offset(0, 1).PasteSpecial
        Else
        Range("A2").Select
        ActiveCell.Value = strLine
        ActiveCell.Offset(0, 1).PasteSpecial
        End If
        Workbooks(strLine).Close True
    End If
End If
    Loop
    MsgBox "Complete!"
    Unload Me
    Application.ScreenUpdating = True
    Range("A1").Select
End Sub

