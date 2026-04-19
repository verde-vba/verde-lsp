Attribute VB_Name = "LikeOperator"
Option Explicit

Sub TestLike()
    Dim text As String
    text = "foo bar"
    
    ' Wildcard patterns
    If text Like "*bar*" Then Debug.Print "contains bar"
    If text Like "*bar" Then Debug.Print "ends in bar"
    If text Like "???" Then Debug.Print "three characters long"
    If text Like "*#*" Then Debug.Print "contains digit"
    If text Like "[A-F]" Then Debug.Print "single letter A-F"
    If text Like "[!A-Z]*" Then Debug.Print "starts with non-uppercase"
End Sub
