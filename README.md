<div align="center">

## Fast Search CoboBox and ListBox using Windows API


</div>

### Description

Using windows API SendMessage Call, this class Searches for a matching string in ListBox (In Association with a textbox) or ComboBox. And believe me its really Fast, Super Fast... ;-)
 
### More Info
 
ctlSource: The Source control (ComboBox or TextBox in case to search in ListBox)

str: The string to search (i.e. <ctlSource>.Text)

intKey : Keycode of Key pressed (i.e. KeyAscii Parameter in KeyPress Event)

Optional ctlTarget : If to search in ListBox The ListBox Control

Usage:

' 1 - In the module declaration declare

Dim cBS As New clsBoxSearch

' 2 - Write on TextBox or ComboBox Keypress event

Private Sub cmbSearch_KeyPress(KeyAscii As Integer)

cBS.FindIndexStr cmbSearch, cmbSearch.Text, KeyAscii

End Sub

Private Sub txtSearchItem_KeyPress(KeyAscii As Integer)

cBS.FindIndexStr txtSearchItem, txtSearchItem.Text, KeyAscii, lstSearchName

End Sub

None: Just sets the ListIndex to the Found String


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Pankaj Nagar](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/pankaj-nagar.md)
**Level**          |Advanced
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/pankaj-nagar-fast-search-cobobox-and-listbox-using-windows-api__1-24966/archive/master.zip)

### API Declarations

```
'**********************************************************************
'Declaration for Search Routines in ListBox (LB) and ComboBox (CB)
Public Const LB_FINDSTRING As Long = &H18F
Public Const LB_FINDSTRINGEXACT As Long = &H1A2
Public Const CB_ERR As Long = (-1)
Public Const LB_ERR As Long = (-1)
Public Const WM_USER As Long = &H400
Public Const CB_FINDSTRING As Long = &H14C
Public Const CB_SHOWDROPDOWN As Long = &H14F
Public Declare Function SendMessageStr Lib _
  "user32" Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   ByVal lParam As String) As Long
'***********************************************************************
```


### Source Code

```
Option Explicit
Public Sub FindIndexStr(ctlSource As Control, _
  ByVal str As String, intKey As Integer, _
  Optional ctlTarget As Variant)
Dim lngIdx As Long
Dim FindString As String
If (intKey < 32 Or intKey > 127) And _
  (Not (intKey = 13 Or intKey = 8)) Then Exit Sub
If Not intKey = 13 Or intKey = 8 Then
  If Len(ctlSource.Text) = 0 Then
    FindString = str & Chr$(intKey)
  Else
    FindString = Left$(str, ctlSource.SelStart) & Chr$(intKey)
  End If
End If
If intKey = 8 Then
  If Len(ctlSource.Text) = 0 Then Exit Sub
  Dim numChars As Integer
  numChars = ctlSource.SelStart - 1
  'FindString = Left(str, numChars)
  If numChars > 0 Then FindString = Left(str, numChars)
End If
If IsMissing(ctlTarget) And TypeName(ctlSource) = "ComboBox" Then
  Set ctlTarget = ctlSource
    If intKey = 13 Then
     Call SendMessageStr(ctlTarget.hWnd, _
       CB_SHOWDROPDOWN, True, 0&)
     Exit Sub
    End If
  lngIdx = SendMessageStr(ctlTarget.hWnd, _
    CB_FINDSTRING, -1, FindString)
ElseIf TypeName(ctlTarget) = "ListBox" Then
  If intKey = 13 Then Exit Sub '???
  lngIdx = SendMessageStr(ctlTarget.hWnd, _
    LB_FINDSTRING, -1, FindString)
Else
  Exit Sub
End If
If lngIdx <> -1 Then
    ctlTarget.ListIndex = lngIdx
    If TypeName(ctlSource) = "TextBox" Then ctlSource.Text = ctlTarget.List(lngIdx)
    ctlSource.SelStart = Len(FindString)
    ctlSource.SelLength = Len(ctlSource.Text) - ctlSource.SelStart
End If
intKey = 0
End Sub
```

