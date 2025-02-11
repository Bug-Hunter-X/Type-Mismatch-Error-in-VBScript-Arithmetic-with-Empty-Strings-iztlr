Function f(a, b)
  If IsEmpty(a) Or IsEmpty(b) Then
    f = 0 ' Or handle the error in a more appropriate manner
    Exit Function
  End If

  If IsNumeric(a) And IsNumeric(b) Then 
    c = CDbl(a) + CDbl(b) 'Explicit type conversion
    f = c
  Else
    f = 0  ' Handle non-numeric input
  End If
End Function

MsgBox f(1, "") 'This will now correctly handle the empty string.