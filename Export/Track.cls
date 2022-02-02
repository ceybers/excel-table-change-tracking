VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Track"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Objects"
Option Explicit

Public ID As Long
Public ValidFrom As Date
Public ValidUntil As Date

Public Function Self() As Track
    Set Self = Me
End Function

Public Function ToString() As String
    ToString = ID & ",#" & Format(ValidFrom, "yyyy/mm/dd") & "#,#" & Format(ValidUntil, "yyyy/mm/dd") & "#"
End Function
