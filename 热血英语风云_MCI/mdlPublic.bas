Attribute VB_Name = "mdlPublic"
Public Type StartLen
    tStart As Long
    tLen As Long
End Type

Public StartLenA() As StartLen
Public StartLenB() As StartLen
Public gstrA As String
Public gstrB As String

Public strCurrentMediaFileName As String
Public IsQuickCompare As Boolean

Public Function Max(A As Long, B As Long) As Long
    If A > B Then
        Max = A
    Else
        Max = B
    End If
End Function

Public Function Min(A As Long, B As Long) As Long
    If A < B Then
        Min = A
    Else
        Min = B
    End If
End Function

