Function MyFunction(param)
  If param Is Nothing Or param = "" Then
    ' Handle both NULL and empty string cases separately
    If param Is Nothing Then
      ' Handle NULL
      param = Null ' Or assign a default value based on your needs
    Else
      ' Handle Empty String
      param = "" ' Or assign a default value based on your needs
    End If
  End If
  ' ... rest of your function ...
End Function