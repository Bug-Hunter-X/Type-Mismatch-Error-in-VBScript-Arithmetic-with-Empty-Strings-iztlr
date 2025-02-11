Function f(a, b)
  If IsEmpty(a) Then Exit Function
  If IsEmpty(b) Then Exit Function
  c = a + b
  f = c
End Function

MsgBox f(1, "") 'This will show a type mismatch error because an empty string is treated as a null value, leading to a type mismatch. 