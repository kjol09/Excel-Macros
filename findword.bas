Attribute VB_Name = "Module1"
Function FindWord(Source As String, Position As Integer)
'Update 20131202
Dim arr() As String
arr = VBA.Split(Source, " ")
xCount = UBound(arr)
If xCount < 1 Or (Position - 1) > xCount Or Position < 0 Then
    FindWord = ""
Else
    FindWord = arr(Position - 1)
End If
End Function

