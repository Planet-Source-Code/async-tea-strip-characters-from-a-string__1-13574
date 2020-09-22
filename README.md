<div align="center">

## Strip Characters from a string


</div>

### Description

This function is to strip all instances of a character out of a string. Its fairly compact and simple. Hope its helpful to someone. :)
 
### More Info
 
The original string without the character in str2Strip


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[async tea](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/async-tea.md)
**Level**          |Beginner
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/async-tea-strip-characters-from-a-string__1-13574/archive/master.zip)





### Source Code

```
Function stripChar(str2BStriped As String, str2Strip As String) As String
  Dim sPos As Long
  Dim newStr As String
  sPos = 1
  Do
    sPos = InStr(str2BStriped, str2Strip)
    If sPos > 0 Then
      newStr = newStr & Left(str2BStriped, sPos - 1)
    Else
      newStr = newStr & str2BStriped
    End If
    str2BStriped = Right(str2BStriped, Len(str2BStriped) - sPos)
  Loop Until sPos = 0
  stripChar = newStr
End Function
```

