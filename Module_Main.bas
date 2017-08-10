Attribute VB_Name = "Module_Main"
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
Public Function GetShortPath(strFileName As String) As String
    'KPD-Team 1999
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    Dim lngRes As Long, strPath As String
    'Create a buffer
    strPath = String$(254, 0)
    'retrieve the short pathname
    lngRes = GetShortPathName(strFileName, strPath, 253)
    'remove all unnecessary chr$(0)'s
    GetShortPath = Left$(strPath, lngRes)
End Function

Public Sub Main()
    WriteStdOut GetShortPath(Command)
End Sub

