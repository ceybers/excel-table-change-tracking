VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GroupedDatabaseChanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("GroupedDatabaseChanges")
Option Explicit

Private Type TGroupedDatabaseChanges
    Items As Collection
End Type

Private This As TGroupedDatabaseChanges

' Properties
'@DefaultMember
Public Property Get Items() As Collection
Attribute Items.VB_UserMemId = 0
    Set Items = This.Items
End Property

Public Property Set Items(ByVal RHS As Collection)
    Set This.Items = RHS
End Property

' Ctor
Private Sub Class_Initialize()
    Set Items = New Collection
End Sub

' Public Methods
Public Function Create(ByVal dbChanges As DatabaseChanges) As GroupedDatabaseChanges
    Set Create = New GroupedDatabaseChanges
    Create.Group dbChanges
End Function

Public Sub Group(ByVal dbChanges As DatabaseChanges)
    Dim pairVar As Variant
    For Each pairVar In dbChanges.DistinctPairs
        With New GroupedDatabaseChange
            .Name = pairVar
            Items.Add .Self, Key:=pairVar
            .Load dbChanges
        End With
    Next pairVar
End Sub

Public Sub UpdateCommitKeys(ByVal CommitKeys As Scripting.Dictionary)
    Dim grpChange As GroupedDatabaseChange
    For Each grpChange In Items
        grpChange.CommitFK = CommitKeys(grpChange.KeyName)
    Next grpChange
End Sub
