Attribute VB_Name = "modSQL"
Option Base 0
Option Explicit

Public Function SQL_NumArray(ByVal strName As String, _
    ByVal strConnector As String, ByVal intHighNumber As Integer, Optional ByVal intLowNumber As Integer = 0, _
    Optional ByVal strCondition As String, Optional ByVal strValue As String) As String
Dim x As Integer

'strConnector = "OR", "AND", "," ...
'strValue = "abc def" for string OR 123 for numbers

If Not strConnector = "," Then strConnector = " " & strConnector & " "

If intHighNumber <= intLowNumber Then
    SQL_NumArray = "[" & strName & " " & intHighNumber & "]"
    Exit Function
Else
    SQL_NumArray = "[" & strName & " " & intLowNumber & "]" & strCondition & strValue
End If

For x = intLowNumber + 1 To intHighNumber
    SQL_NumArray = SQL_NumArray & strConnector & "[" & strName & " " & x & "]" _
        & strCondition & strValue
Next x

End Function
