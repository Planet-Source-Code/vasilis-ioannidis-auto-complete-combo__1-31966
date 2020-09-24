<div align="center">

## Auto Complete Combo


</div>

### Description

It's sole purpose is to autocomplete-enable a combo box with as few coding as possible. Uses windows messaging sub-system to send directly a message to to the window's windows procedure. It's not mine but I thought it was great and wanted to share it with you. So no voting people. Cheers :)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Vasilis Ioannidis](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vasilis-ioannidis.md)
**Level**          |Intermediate
**User Rating**    |4.8 (24 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vasilis-ioannidis-auto-complete-combo__1-31966/archive/master.zip)

### API Declarations

```
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Const CB_FINDSTRING = &H14C
```


### Source Code

```
Option Explicit
Private Sub Form_Load()
With Combo1
  .AddItem "ABCD"
  .AddItem "ACDE"
  .AddItem "ADEF"
  .AddItem "AEFG"
  .AddItem "ACFG"
  .AddItem "AFGH"
  .AddItem "AGHI"
 End With
End Sub
Private Sub Combo1_KeyPress(KeyAscii As Integer)
Dim i As Long
Dim iNewStart As Integer
Dim strTemp As String
'Figure out the string prefix to search gor
  If Combo1.SelStart = 0 Then
    strTemp = Combo1.Text & Chr(KeyAscii)
  Else
    strTemp = Left(Combo1.Text, Combo1.SelStart) & Chr(KeyAscii)
  End If
'Pass -1 as lParam to search entire list
  i = SendMessage(Combo1.hwnd, CB_FINDSTRING, -1, strTemp)
'-1 return code indicates failure to find the string
  If i <> -1 Then
    'SendMessage returns the index of the first occurrence
    'of strTemp in the combo's list.
    Combo1.Text = Combo1.List(i)
    'Set the text selection appropriately for the
    'suggested match
    Combo1.SelStart = Len(strTemp)
    Combo1.SelLength = Len(Combo1.List(i)) - Len(strTemp)
    KeyAscii = 0
  End If
End Sub
```

