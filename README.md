<div align="center">

## FindWindowWild


</div>

### Description

Find window using full or part of it's caption. Allow wild characters (*,?,[]). For example, using this string :"*Mi??OSoFt In[s-u]ernet*" you can find Microsoft Internet Explorer window.
 
### More Info
 
Full or part of window's caption. Wild characters accepted.

Handle of window if find, zero otherwise.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ark](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ark.md)
**Level**          |Intermediate
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ark-findwindowwild__1-10259/archive/master.zip)





### Source Code

```
'---Bas module code------
Private Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function IsWindowVisible& Lib "user32" (ByVal hwnd As Long)
Private Declare Function GetParent& Lib "user32" (ByVal hwnd As Long)
Dim sPattern As String, hFind As Long
Function EnumWinProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
 Dim k As Long, sName As String
 If IsWindowVisible(hwnd) And GetParent(hwnd) = 0 Then
   sName = Space$(128)
   k = GetWindowText(hwnd, sName, 128)
   If k > 0 Then
    sName = Left$(sName, k)
    If lParam = 0 Then sName = UCase(sName)
    If sName Like sPattern Then
      hFind = hwnd
      EnumWinProc = 0
      Exit Function
    End If
   End If
 End If
 EnumWinProc = 1
End Function
Public Function FindWindowWild(sWild As String, Optional bMatchCase As Boolean = True) As Long
 sPattern = sWild
 If Not bMatchCase Then sPattern = UCase(sPattern)
 EnumWindows AddressOf EnumWinProc, bMatchCase
 FindWindowWild = hFind
End Function
'----Using (Form code)----
Private Sub Command1_Click()
 Debug.Print FindWindowWild("*Mi??OSoFt In[s-u]ernet*", False)
End Sub
```

