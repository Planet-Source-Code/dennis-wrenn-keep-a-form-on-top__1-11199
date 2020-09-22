<div align="center">

## Keep a form on top\!


</div>

### Description

This code keeps a form on top of all other windows.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dennis Wrenn](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dennis-wrenn.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dennis-wrenn-keep-a-form-on-top__1-11199/archive/master.zip)

### API Declarations

```
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
```


### Source Code

```
'code:
Private Sub FormOnTop(frm As Form, blnOnTop As Boolean)
  Dim lPos As Long
  Select Case blnOnTop
    Case True
      lPos = HWND_TOPMOST
    Case False
      lPos = HWND_NOTOPMOST
  End Select
  Call SetWindowPos(frm.hwnd, lPos, 0, 0, 0, 0, SWP_FLAGS)
End Sub
'usage:
Private Sub Form_Load()
'makes a form on top
  Call FormOnTop(Me, True)
End Sub
Private Sub Command1_Click()
'makes a form not always on top anymore..
  Call FormOnTop(Me, False)
End Sub
```

