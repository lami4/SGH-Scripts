Sub Example1()
lol = ActiveDocument.Sections(1).Footers(1).Range.Tables(1).Cell(5, 8).Range.Text
lolz = ActiveDocument.Tables(1).Cell(2, 1).Range.Text

MsgBox (lol)
End Sub
