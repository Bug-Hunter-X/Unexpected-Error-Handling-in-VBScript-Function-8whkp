Function MyFunction(param1)
  On Error GoTo ErrorHandler
  If IsEmpty(param1) Then
    Err.Raise vbError, , "Parameter cannot be empty"
    Exit Function
  End If
  ' ... rest of the function
  Exit Function
ErrorHandler:
  MsgBox "An error occurred: " & Err.Description, vbCritical
  Err.Clear
End Function