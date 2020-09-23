<div align="center">

## CheckKeypress


</div>

### Description


 
### More Info
 
KeyAscii from controls keypress-event

cAllowed as a string ("N" for numbers only...)

Put the call in a controls keypress-event

Ex: KeyAscii=CheckKeypress(KeyAscii, "N")

Will only allow the user to enter digits in a control

KeyAscii for pressed key if ok or nothing and a beep for not ok


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Martin Idman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/martin-idman.md)
**Level**          |Unknown
**User Rating**    |3.5 (14 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/martin-idman-checkkeypress__1-1576/archive/master.zip)





### Source Code

```
Public Function CheckKeyPress(iKeyIn As Integer, cAllowed As String) As Integer
  Dim cValidKeys As String
  Select Case cAllowed
   Case "N" ' Just numbers
     cValidKeys = "1234567890" & vbCr & vbTab & vbBack
   Case "N1" ' Decimal numbers
     cValidKeys = "1234567890," & vbCr & vbTab & vbBack
   Case "N2" ' Simple math
     cValidKeys = "1234567890+-*/=," & vbCr & vbTab & vbBack
   Case "C" ' Simple characterset(I'm Swedish, hence some strange ones)
     cValidKeys = "ABCDEFGHIJKLMNOPQRSTUVWXYZÅÄÖAAÉEÜI- " & vbCr & vbTab & vbBack
   Case "C1" ' Enhanced characterset
     cValidKeys = "ABCDEFGHIJKLMNOPQRSTUVWXYZÅÄÖAAÉEÜI&#,.-/\+-*%$<>:;@!?=() " & vbCr & vbTab & vbBack
   Case "C2" ' Enhanced + digits
     cValidKeys = "ABCDEFGHIJKLMNOPQRSTUVWXYZÅÄÖAAÉEÜI1234567890½&#,.-/\+-*%$<>:;@!?=() " & vbCr & vbTab & vbBack
   Case "M" ' Mail and WWW
     cValidKeys = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890-_/\~:@." & vbCr & vbTab & vbBack
   Case "D" ' Date or telephonenumbers
     cValidKeys = "0123456789-" & vbCr & vbTab & vbBack
  End Select
  If InStr(cValidKeys, UCase(Chr(iKeyIn))) Then
     CheckKeyPress = iKeyIn
  Else
   Beep
   CheckKeyPress = 0
  End If
End Function
```

