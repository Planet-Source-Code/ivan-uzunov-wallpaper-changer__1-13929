<div align="center">

## Wallpaper Changer


</div>

### Description

If you are wondering how you can change your desktop picture this code is the awnser. It's very easy to understend and it's very short.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ivan Uzunov](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ivan-uzunov.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ivan-uzunov-wallpaper-changer__1-13929/archive/master.zip)





### Source Code

```
Option Explicit
'This code is developed by Ivan Uzunov
'e-mail: kicheto@goatrance.com
'Just add this code on a form add a Command1 and press F5
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_SETDESKWALLPAPER = 20
Private Sub Command1_Click()
Dim WallPaper As Long
  'Just change "C:\REDCAP.bmp" with a existing bitmap on your computer
  WallPaper = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, "C:\REDCAP.bmp", 0)
End Sub
```

