<div align="center">

## Progress Bar Non\-OCX


</div>

### Description

Fill the progressbar up to the percent given.
 
### More Info
 
You need to add the PictureBox to a form and that's about it.

Progressbar value.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Slikk](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/slikk.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/slikk-progress-bar-non-ocx__1-36112/archive/master.zip)





### Source Code

```
'Visit: http://slikk.emunx.net/ or http://slikk.emunx.net/forum and support our site, we just got it up and running and we need visitors, thanks.
Public Sub UpdateProgress(pBar As PictureBox, Percent As Single, Optional ByVal ShowPercent = False)
Dim sData As String
If Not pBar.AutoRedraw Then '//no memory?
pBar.AutoRedraw = -1 'nope, we make some
End If
pBar.Cls '//clear memory
pBar.ScaleWidth = 100 'set scale's width
pBar.DrawMode = 10 '//set draw modus
If ShowPercent = True Then
sData$ = Format(Percent, "###0") + "%" '//format percent
pBar.CurrentX = 50 - pBar.TextWidth(sData$) 'set the currentX for the progress
pBar.CurrentY = pBar.ScaleHeight - pBar.TextHeight(sData$) 'set the currentY for the progress
Debug.Print sData$ 'show percent in debug window
End If
pBar.Line (0, 0)-(Percent, pBar.ScaleHeight), , BF
pBar.Refresh
End Sub
'Call UpdateProgress (Picture1, 50, False)
```

