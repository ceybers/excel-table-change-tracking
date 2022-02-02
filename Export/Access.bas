Attribute VB_Name = "Access"
'@Folder("VBAProject")
Option Explicit

Public Const COMMITS_TABLE_NAME As String = "metaCommits"
Public Const KEY_TABLE_NAME As String = "metaKeys"

Public Const SCHEMA_PATH As String = "C:\Users\User\Documents\excel-table-change-tracking\Schema.csv"
Private Const DATABASE_PATH As String = "C:\Users\User\Documents\excel-table-change-tracking\Database1.accdb"

Public Function GetConnectionString()
    GetConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DATABASE_PATH & ";"
End Function
