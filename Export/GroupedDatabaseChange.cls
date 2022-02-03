VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GroupedDatabaseChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "GroupedDatabaseChanges"
Option Explicit

Private Type TGroupedDatabaseChange
    Name As String
    KeyName As String
    KeyFK As Long
    TableName As String
    CommitFK As Long
    trackFK As Long
    Items As Collection
End Type

Private this As TGroupedDatabaseChange

' Properties
Public Property Get Name() As String
    Name = this.Name
End Property

Public Property Let Name(ByVal RHS As String)
    this.Name = RHS
End Property

Public Property Get KeyName() As String
    KeyName = this.KeyName
End Property

Public Property Get KeyFK() As Long
    KeyFK = this.KeyFK
End Property

Public Property Get TableName() As String
    TableName = this.TableName
End Property

Public Property Get CommitFK() As Long
    CommitFK = this.CommitFK
End Property

Public Property Let CommitFK(ByVal RHS As Long)
    this.CommitFK = RHS
End Property

Public Property Get trackFK() As Long
    trackFK = this.trackFK
End Property

Public Property Let trackFK(ByVal RHS As Long)
    this.trackFK = RHS
End Property

Public Property Get Items() As Collection
    Set Items = this.Items
End Property

Private Sub Class_Initialize()
    Set this.Items = New Collection
End Sub

' Public Methods
Public Function Self() As GroupedDatabaseChange
    Set Self = Me
End Function

Public Sub Load(ByVal dbChanges As DatabaseChanges)
    Dim dbChange As DatabaseChange
    For Each dbChange In dbChanges.Items
        If dbChange.Pair = Name Then
            Items.Add dbChange ' Key?
        End If
    Next dbChange
    
    SetGroupFromFirstItem
End Sub

' Private Methods
Private Sub SetGroupFromFirstItem()
    If Items.Count > 0 Then
        Dim firstItem As DatabaseChange
        Set firstItem = Items(1)
        With firstItem
            this.KeyName = .KeyName
            this.KeyFK = .KeyFK
            this.TableName = .TableName
        End With
    End If
End Sub
