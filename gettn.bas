Attribute VB_Name = "Module2"
Option Explicit
Function GetTN(ByRef TNs As String, ByVal Position As Integer) As String
' Get 1st 2nd 3rd... word or number from a string of text and numbers separated by a space or spaces.
' Author:  Stanley D. Grom, Jr., stanleydgromjr at ExcelForum.com, hiker95 at MrExcel.com, Stanley D. Grom at Ozgrid.com
' Updated: September 03, 2010
'
' A1 is equal to (without the " marks): "Aa 11 Bb 22 Cc 33Dd Ee44"
' B1: =GetTN(A1, 5)
' The result is "Cc"
'
' =GetTN(A1,Column()-1)
'
'Modification from:
'Function GetClaim(ByVal Claims As Range, ByVal Position As Integer) As String
'Function GetClaim(ByVal Claims As String, ByVal Position As Integer) As String
'Author:  Leith Ross
'http://www.excelforum.com/excel-programming/744090-string-manipulation-split-into-multiple-variables.html
'
Dim Cnt As Integer
Dim Matches As Object
Dim RegExp As Object
Dim S As String, Text As String
Application.Volatile
Set RegExp = CreateObject("VBScript.RegExp")
RegExp.Global = True
RegExp.IgnoreCase = True
RegExp.Pattern = "\s*(\S+)\s+(.*)"
Text = TNs & " "
Do While RegExp.Test(Text)
  S = RegExp.Replace(Text, "$1")
  Text = RegExp.Replace(Text, "$2")
  Cnt = Cnt + 1
  If Cnt = Position Then GetTN = S
Loop
End Function
