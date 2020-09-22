<div align="center">

## Character shaped forms\!\!


</div>

### Description

Have you ever wanted to make your form's shape odd? Ok, there are several samples and programs around that can make your forms like a shape(circle, rounded box or something a little bit more complicated). But here is the example to make your form's shape to be ANY TEXT, in ANY FONT, in ANY SIZE and also any two colour's gradient. It's a really good example. Imagine you can shape the form not to be just plain text, but the shape of special fonts(such as Windings and Webdings). Just change the GetTextRgn function's variables(Font, Size, Text) and the variable Color1 and Color2. Easy. And the result is outstanding! You can also use the Chr$ function to add a text(this is useful for spec. chars).
 
### More Info
 
Change the GetTextRgn function's variables(Font, Size, Text) and the variable Color1 and Color2. Easy. And the result is outstanding! You can also use the Chr$ function to add a text(this is useful for spec. chars).

Copy ALL the code to a blank form.(remove Form_Load() first) Then after setting the parameters mentioned (or leave them for first check) run the project.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[bbence](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bbence.md)
**Level**          |Intermediate
**User Rating**    |5.0 (55 globes from 11 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bbence-character-shaped-forms__1-5448/archive/master.zip)

### API Declarations

It has Api calls, but I think its much easier to copy everything at once.


### Source Code

```
Option Explicit
Private Type RECT
 Left As Long
 Top As Long
 Right As Long
 Bottom As Long
End Type
'API calls required for doing this cool stuff
Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function PathToRegion Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const RGN_AND = 1
Dim Color1 As Long
Dim Color2 As Long
Private Function GetTextRgn(Font As String, Size As Integer, Text As String) As Long
Me.Font = Font
Me.FontSize = Size
 Dim hRgn1 As Long, hRgn2 As Long
 Dim rct As RECT
 BeginPath hdc
 TextOut hdc, 10, 10, Text, Len(Text)
 EndPath hdc
 hRgn1 = PathToRegion(hdc)
 GetRgnBox hRgn1, rct
 hRgn2 = CreateRectRgnIndirect(rct)
 CombineRgn hRgn2, hRgn2, hRgn1, RGN_AND
 DeleteObject hRgn1
 GetTextRgn = hRgn2
End Function
Private Sub GradateColors(Colors() As Long, ByVal Color1 As Long, ByVal Color2 As Long)
 On Error Resume Next
 Dim i As Integer
 Dim dblR As Double, dblG As Double, dblB As Double
 Dim addR As Double, addG As Double, addB As Double
 Dim bckR As Double, bckG As Double, bckB As Double
 dblR = CDbl(Color1 And &HFF)
 dblG = CDbl(Color1 And &HFF00&) / 255
 dblB = CDbl(Color1 And &HFF0000) / &HFF00&
 bckR = CDbl(Color2 And &HFF&)
 bckG = CDbl(Color2 And &HFF00&) / 255
 bckB = CDbl(Color2 And &HFF0000) / &HFF00&
 addR = (bckR - dblR) / UBound(Colors)
 addG = (bckG - dblG) / UBound(Colors)
 addB = (bckB - dblB) / UBound(Colors)
 For i = 0 To UBound(Colors)
  dblR = dblR + addR
  dblG = dblG + addG
  dblB = dblB + addB
  If dblR > 255 Then dblR = 255
  If dblG > 255 Then dblG = 255
  If dblB > 255 Then dblB = 255
  If dblR < 0 Then dblR = 0
  If dblG < 0 Then dblG = 0
  If dblG < 0 Then dblB = 0
  Colors(i) = RGB(dblR, dblG, dblB)
 Next
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'these are for moving the form without its titlebar
 ReleaseCapture
 SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
End Sub
Private Sub Form_Paint()
 Dim Colors() As Long
 Dim Iter As Long
 Const Banding = 8
 ReDim Colors(ScaleHeight \ Banding) As Long
 GradateColors Colors(), Color1, Color2
 For Iter = 0 To ScaleHeight Step Banding
  Line (0, Iter)-(ScaleWidth, Iter + Banding), Colors(Iter \ Banding), BF
 Next
End Sub
Private Sub Form_Load()
 Dim hRgn As Long
 hRgn = GetTextRgn("Wingdings", 100, "J" & "<") 'change the values: Font, Size (font), Text
 SetWindowRgn hWnd, hRgn, 1
 Color1 = vbBlack 'set this colours for gradient effect (use vb colour constants for easy use)
 Color2 = vbBlue
 Me.Refresh
End Sub
```

