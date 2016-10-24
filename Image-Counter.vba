Dim PathToReport As String
Dim Row As Integer
Dim CurDir As String
Sub SpecifyPath()
CurDir = ActiveDocument.Path
'Specify file path and name below
PathToReport = CurDir & "\ImageCountReport.xlsx"
    If Dir(PathToReport) <> "" Then
    Kill (PathToReport)
    End If
Dim ExcelApp As Object
Set ExcelApp = CreateObject("Excel.Application")
ExcelApp.WorkBooks.Add
ExcelApp.Cells(1, 1) = "File name"
ExcelApp.Cells(1, 2) = "Images"
ExcelApp.ActiveWorkBook.SaveAs FileName:=PathToReport
ExcelApp.ActiveWorkBook.Close

Row = 2
Dim intResult As Integer
Dim strPath As String
Dim arrFiles() As String
Dim i As Integer

intResult = Application.FileDialog(msoFileDialogFolderPicker).Show

If intResult <> 0 Then
    
    strPath = Application.FileDialog( _
        msoFileDialogFolderPicker).SelectedItems(1)
    arrFiles() = GetAllFilePaths(strPath)
    For i = LBound(arrFiles) To UBound(arrFiles)
        Call ModifyFile(arrFiles(i))
    Next i
End If

    ExcelApp.WorkBooks.Open (PathToReport)
    ExcelApp.Application.Visible = False
    ExcelApp.Columns("A").ColumnWidth = 45
    ExcelApp.Columns("B").ColumnWidth = 10
    ExcelApp.Range("A1:B1").HorizontalAlignment = xlCenter
    ExcelApp.Columns("B").HorizontalAlignment = xlCenter
    ExcelApp.Range("A1:B1").Font.Bold = True
    ExcelApp.Cells(Row, 1) = "TOTAL:"
    ExcelApp.Cells(Row, 1).HorizontalAlignment = xlRight
    ExcelApp.Cells(Row, 1).Font.Bold = True
    Dim forCalculation As Integer
    forCalculation = Row - 1
    ExcelApp.Cells(Row, 2).Value = "=Sum(B2:B" & forCalculation & ")"
    ExcelApp.ActiveWorkBook.Save
    ExcelApp.ActiveWorkBook.Close
    Application.Quit SaveChanges:=wdDoNotSaveChanges
End Sub
Private Sub ModifyFile(ByVal strPath As String)
Dim objDocument As Document
Set objDocument = Documents.Open(strPath)
objDocument.Activate

Dim counter As Integer
counter = 0
Dim TextFile As Integer
TextFile = FreeFile
Dim DocName As String
DocName = objDocument.Name

For Each iShape In objDocument.InlineShapes
    'Filter can be adjusted. In MS Word width 112 = ~4 cm, height 20 = ~70 mm.
    If iShape.Width >= 112 And iShape.Height >= 20 Then
    counter = counter + 1
    End If
Next iShape
For Each Shape In objDocument.Shapes
    'Filter can be adjusted. In MS Word width 112 = ~4 cm, height 20 = ~70 mm.
    If Shape.Width >= 112 And Shape.Height >= 20 Then
    counter = counter + 1
    End If
Next Shape
counter = counter - 1

If counter > 0 Then
    Dim ExcelApp As Object
    Set ExcelApp = CreateObject("Excel.Application")
    ExcelApp.WorkBooks.Open (PathToReport)
    ExcelApp.Application.Visible = False
    ExcelApp.Cells(Row, 1) = DocName
    ExcelApp.Cells(Row, 2) = counter
    ExcelApp.ActiveWorkBook.Save
    ExcelApp.ActiveWorkBook.Close
    Row = Row + 1
End If

objDocument.Close (True)
End Sub
Private Function GetAllFilePaths(ByVal strPath As String) _
As String()
Dim objFSO As Object
Dim objFolder As Object
Dim objFile As Object
Dim i As Integer
Dim arrOutput() As String
ReDim arrOutput(1 To 1)
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(strPath)
i = 1
For Each objFile In objFolder.Files
    ReDim Preserve arrOutput(1 To i)
    arrOutput(i) = objFile.Path
    i = i + 1
Next objFile
GetAllFilePaths = arrOutput
End Function
