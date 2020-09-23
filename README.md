<div align="center">

## \_ String Functions \_


</div>

### Description

Includes many common useful string functions. Reverse string, Remove extra spaces, Delimit string, Alternating caps, Proper case, and Count number of occurances of a string in a string. Vote if you like it!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[KRYO\_11](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kryo-11.md)
**Level**          |Intermediate
**User Rating**    |4.3 (30 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kryo-11-string-functions__1-50381/archive/master.zip)





### Source Code

```
Public Function ReverseString(TheString As String) As String
  ReverseString = ""
  For i = 0 To Len(TheString) - 1
    ReverseString = ReverseString & Mid(TheString, Len(TheString) - i, 1)
  Next i
End Function
Public Function RemoveExtraSpaces(TheString As String) As String
  Dim LastChar As String
  Dim NextChar As String
  LastChar = Left(TheString, 1)
  RemoveExtraSpaces = LastChar
  For i = 2 To Len(TheString)
    NextChar = Mid(TheString, i, 1)
    If NextChar = " " And LastChar = " " Then
    Else
      RemoveExtraSpaces = RemoveExtraSpaces & NextChar
    End If
    LastChar = NextChar
  Next i
End Function
Public Function DelimitString(TheString As String, Delimiter As String) As String
  DelimitString = ""
  For i = 1 To Len(TheString)
    If i <> Len(TheString) Then
      DelimitString = DelimitString & Mid(TheString, i, 1) & Delimiter
    Else
      DelimitString = DelimitString & Mid(TheString, i, 1)
    End If
  Next i
End Function
Public Function AltCaps(TheString As String, Optional StartWithFirstCharacter As Boolean = True) As String
  Dim LastCap As Boolean
  AltCaps = ""
  If StartWithFirstCharacter = False Then LastCap = True
  For i = 1 To Len(TheString)
    If LastCap = False Then
      AltCaps = AltCaps & UCase(Mid(TheString, i, 1))
      LastCap = True
    Else
      AltCaps = AltCaps & LCase(Mid(TheString, i, 1))
      LastCap = False
    End If
  Next i
End Function
Public Function Propercase(TheString As String) As String
  Propercase = UCase(Left(TheString, 1))
  For i = 2 To Len(TheString)
    If Mid(TheString, i - 1, 1) = " " Then
      Propercase = Propercase & UCase(Mid(TheString, i, 1))
    Else
      Propercase = Propercase & LCase(Mid(TheString, i, 1))
    End If
  Next i
End Function
Public Function CountCharacters(TheString As String, CharactersToCheckFor As String) As Integer
   Dim Char As String
   Dim ReturnAgain As Boolean
   CountCharacters = 0
   For i = 1 To Len(TheString)
    If i < (Len(TheString) + 1 - Len(CharactersToCheckFor)) Then
      Char = Mid(TheString, i, Len(CharactersToCheckFor))
      ReturnAgain = True
    Else
      Char = Mid(TheString, i)
      ReturnAgain = False
    End If
    If Char = CharactersToCheckFor Then CountCharacters = CountCharacters + 1
    If ReturnAgain = False Then GoTo NextPos
  Next i
NextPos:
End Function
```

