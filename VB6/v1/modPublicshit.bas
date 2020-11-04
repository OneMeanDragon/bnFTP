Attribute VB_Name = "modPublicshit"
Public Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)


Public Type FILETIME
    dwLowDateTime   As Long
    dwHighDateTime  As Long
End Type

Function MakeFileTime(ft As String) As FILETIME
    If Len(ft) < 8 Then
        MakeFileTime.dwHighDateTime = 0
        MakeFileTime.dwLowDateTime = MakeFileTime.dwHighDateTime
        Exit Function
    End If
    Dim H As String
    Dim l As String
    l = Mid(ft, 1, 4)
    H = Mid(ft, 5, 4)
    
    RtlMoveMemory MakeFileTime.dwHighDateTime, ByVal H, 4
    RtlMoveMemory MakeFileTime.dwLowDateTime, ByVal l, 4
End Function
Function MakeLong(x As String) As Long
    If Len(x) < 4 Then
        Exit Function
    End If
    RtlMoveMemory MakeLong, ByVal x, 4
End Function
Function MakeShort(x As String) As Integer
    If Len(x) < 2 Then
        Exit Function
    End If
    RtlMoveMemory MakeShort, ByVal x, 2
End Function

Public Function MakeDWORD(value As Long) As String
Dim strReturn As String * 4
    RtlMoveMemory ByVal strReturn, value, 4
    MakeDWORD = strReturn
End Function
Public Function MakeWORD(value As Integer) As String
Dim strReturn As String * 2
    RtlMoveMemory ByVal strReturn, value, 2
    MakeWORD = strReturn
End Function
