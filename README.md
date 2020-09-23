<div align="center">

## Write 3D Text on Form/Picture


</div>

### Description

A small sub for 3D Text on a Form or Picture box. You can define the depth of the Text, the color, the font and the fontsize.
 
### More Info
 
To write on a picturebox replace "Form1" with "Picture1"


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Michael Frey](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/michael-frey.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/michael-frey-write-3d-text-on-form-picture__1-7487/archive/master.zip)





### Source Code

```
Private Sub Form_Load()
Text3D "Hallo", "Times New Roman", 26, 1500, 200, 100, 146, 16, 46
End Sub
Public Sub Text3D(Strng As String, Fnt As String, Font_size As Integer, XVal As Integer, YVal As Integer, Depth As Integer, Redcol As Integer, Greencol As Integer, Bluecol As Integer)
Form1.AutoRedraw = True
Form1.FontSize = Font_size
Form1.Font = Fnt
Form1.ForeColor = RGB(Redcol, Greencol, Bluecol)
ShadowY = YVal
ShadowX = XVal
For i = 0 To Depth
Form1.CurrentX = ShadowX - i
Form1.CurrentY = ShadowY + i
If i = Depth Then Form1.ForeColor = RGB(Redcol + 80, Greencol + 80, Bluecol + 80)
Form1.Print Strng
Next i
Form1.AutoRedraw = False
End Sub
```

