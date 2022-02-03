Attribute VB_Name = "Access"
'@Folder "Access"
Option Explicit

Public Const SCHEMA_PATH As String = "C:\Users\User\Documents\excel-table-change-tracking\Schema.csv"
Private Const DATABASE_PATH As String = "C:\Users\User\Documents\excel-table-change-tracking\Database1.accdb"

Public Const COMMITS_TABLE_NAME As String = "metaCommits"
Public Const KEY_TABLE_NAME As String = "metaKeys"
Public Const TRACKS_TABLE_NAME As String = "metaTracks"

Public Const TRACK_FIELD_NAME As String = "TrackFK"
Public Const COMMIT_FIELD_NAME  As String = "CommitFK"
Public Const KEY_FIELD_NAME As String = "KeyFK"
Public Const VALID_UNTIL_FIELD_NAME As String = "ValidUntil"
        
Public Function GetConnectionString() As String
    GetConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DATABASE_PATH & ";"
End Function

Public Function GetConnection(Optional ByVal DoClose As Boolean = False) As ADODB.Connection
    Static Connection As ADODB.Connection
    
    If DoClose Then
        If Not Connection Is Nothing Then
            Connection.Close
        End If
        Set Connection = Nothing
        Exit Function
    End If
    
    If Connection Is Nothing Then
        Set Connection = New ADODB.Connection
    End If
    
    If Connection.State <> adStateOpen Then
        Connection.Open ConnectionString:=Access.GetConnectionString
    End If
    
    Set GetConnection = Connection
End Function
