Sub Example1()
Dim intResult As Integer
Dim strPath As String
Dim arrFiles() As String
Dim i As Integer
'the dialog is displayed to the user
intResult = Application.FileDialog(msoFileDialogFolderPicker).Show
'checks if user has cancled the dialog
If intResult <> 0 Then
    'dispaly message box
    strPath = Application.FileDialog( _
        msoFileDialogFolderPicker).SelectedItems(1)
    arrFiles() = GetAllFilePaths(strPath)
    For i = LBound(arrFiles) To UBound(arrFiles)
        Call ModifyFile(arrFiles(i))
    Next i
End If
End Sub

Private Sub ModifyFile(ByVal strPath As String)
Dim objDocument As Document
Set objDocument = Documents.Open(strPath)
objDocument.Activate
Selection.WholeStory
Selection.Font.Name = "GOST Type A"
objDocument.Save
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
'Create an instance of the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Get the folder object
Set objFolder = objFSO.GetFolder(strPath)
i = 1
'loops through each file in the directory and
'prints their names and path
For Each objFile In objFolder.Files
    ReDim Preserve arrOutput(1 To i)
    'print file path
    arrOutput(i) = objFile.Path
    i = i + 1
Next objFile
GetAllFilePaths = arrOutput
End Function
