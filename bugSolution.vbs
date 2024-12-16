Improved Function in VBScript:

```vbscript
Function IsWithinRange(number, min, max)
  ' Check if number is numeric
  If IsNumeric(number) Then
    ' Handle null values
    If IsNull(number) Or number = "" Then
      IsWithinRange = False
    ElseIf number >= min And number <= max Then
      IsWithinRange = True
    Else
      IsWithinRange = False
    End If
  Else
    ' Handle non-numeric input
    IsWithinRange = False  
  End If
End Function
```

This improved version explicitly checks if the input `number` is numeric using `IsNumeric()`.  It also handles null or empty string values and returns `False` in these cases. If the input is not numeric it also returns `False`. This provides a more robust and reliable range check function.