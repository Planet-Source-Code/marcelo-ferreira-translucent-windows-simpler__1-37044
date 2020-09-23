<div align="center">

## Translucent Windows \- Simpler


</div>

### Description

Creates a translucent window. No DLL/OCX, No flick, No Static, No headache !!

You create a new form, copy and paste into General Declarations section, and [F5]... " Já foi pra conta!" It's done...

<<Credits: http://support.microsoft.com/default.aspx?scid=kb;EN-US;q249341>>
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Marcelo Ferreira](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/marcelo-ferreira.md)
**Level**          |Beginner
**User Rating**    |4.9 (34 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0, VBA MS Access
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/marcelo-ferreira-translucent-windows-simpler__1-37044/archive/master.zip)





### Source Code

```
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias _
"GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias _
"SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
 ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
 (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, _
 ByVal dwFlags As Long) As Long
Private Const Estilo = (-20)
Private Const Camada = &H80000
Private Const CorAlpha = &H2&
'-------------------
Private Sub Form_Load()
  Dim AntigoEstilo As Long
  Dim Nivel As Byte ' Transparency (0 - 255)
  Nivel = 180
  AntigoEstilo = GetWindowLong(Me.hwnd, Estilo)
  SetWindowLong Me.hwnd, Estilo, AntigoEstilo Or Camada
  SetLayeredWindowAttributes Me.hwnd, 0, Nivel, CorAlpha
End Sub
```

