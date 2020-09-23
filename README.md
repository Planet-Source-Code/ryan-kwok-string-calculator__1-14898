<div align="center">

## String Calculator


</div>

### Description

This is a recursive function to calculate a string formula. For example: You have a formula like "3*((2+6)*2-2)/(2+2)". You just call the function to get the answer. Hope this is useful to somebody:)
 
### More Info
 
sFormula - string

return the result of the formula


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ryan Kwok](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ryan-kwok.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ryan-kwok-string-calculator__1-14898/archive/master.zip)





### Source Code

```
Public Function calc(sFormula As String) As Double
'This is a recursive function to calculate a valid
'math formula.
 Dim sHead As String, sTail As String
 Dim sTemp As String, lPos As Long
 Dim cnt As Long, dblTemp As Double
 Dim I As Long
 cnt = 0
 If InStr(sFormula, "(") > 0 Then
  'calculate the string within bracket first
  lPos = InStr(sFormula, "(")
  For I = lPos + 1 To Len(sFormula)
   If Mid(sFormula, I, 1) = "(" Then cnt = cnt + 1
   If Mid(sFormula, I, 1) = ")" Then
    If cnt = 0 Then Exit For
    cnt = cnt - 1
   End If
  Next
  sTemp = Mid(sFormula, lPos + 1, I - lPos - 1)
  dblTemp = calc(sTemp)
  sTemp = Replace(sFormula, "(" & sTemp & ")", CStr(dblTemp))
  calc = calc(sTemp)
 ElseIf InStr(sFormula, "+") > 0 Then
  'Add
  lPos = InStr(sFormula, "+")
  sHead = Left(sFormula, lPos - 1)
  sTail = Right(sFormula, Len(sFormula) - lPos)
  calc = calc(sHead) + calc(sTail)
 ElseIf InStr(sFormula, "-") > 0 Then
  'Subtract
  lPos = InStr(sFormula, "-")
  sHead = Left(sFormula, lPos - 1)
  sTail = Right(sFormula, Len(sFormula) - lPos)
  calc = calc(sHead) - calc(sTail)
 ElseIf InStr(sFormula, "*") > 0 Then
  'Multiply
  lPos = InStr(sFormula, "*")
  sHead = Left(sFormula, lPos - 1)
  sTail = Right(sFormula, Len(sFormula) - lPos)
  calc = calc(sHead) * calc(sTail)
 ElseIf InStr(sFormula, "/") > 0 Then
  'Divide
  lPos = InStr(sFormula, "/")
  sHead = Left(sFormula, lPos - 1)
  sTail = Right(sFormula, Len(sFormula) - lPos)
  calc = calc(sHead) / calc(sTail)
 Else
  calc = CDbl(sFormula)
 End If
End Function
```

