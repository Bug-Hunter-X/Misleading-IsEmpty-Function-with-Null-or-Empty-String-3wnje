Function MyFunc(param)
  If IsEmpty(param) Or Len(Trim(param)) = 0 Then
    Err.Raise 13, , "Parameter cannot be empty or null"
  End If
  ' ... rest of the function
End Function