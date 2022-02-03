Attribute VB_Name = "TestFieldChanges"
'@Folder("FieldChanges")
Option Explicit

Public Function GetTestFieldChanges() As FieldChanges
    With New FieldChanges
        .Items.Add FieldChange.Create("JKL", "FieldA", CStr("sit"), CStr("abc"))
        .Items.Add FieldChange.Create("JKL", "FieldB", CDbl("44652"), CDate("2022/01/31"))
        .Items.Add FieldChange.Create("ABC", "FieldC", CDbl("1"), CDbl("52"))
        Set GetTestFieldChanges = .Self
    End With
End Function
