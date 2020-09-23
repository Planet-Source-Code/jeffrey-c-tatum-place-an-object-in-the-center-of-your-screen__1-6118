<div align="center">

## Place an Object in the Center of your Screen


</div>

### Description

This will take an object on a form, and place it in the center of the screen. Good for password protection programs that take up the whole screen, this will ensure that an object that you want in the middle of the screen, will remain there for any resolution.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jeffrey C\. Tatum](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jeffrey-c-tatum.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jeffrey-c-tatum-place-an-object-in-the-center-of-your-screen__1-6118/archive/master.zip)





### Source Code

```
'Place this in a module for use in the future:
Public Sub Object_Center(Frm As Form, Cntrl As Object)
'To call it, you would simply do:
'Call Object_Center(Form1, Text1)
'and that will center Text1 in the middle of your
'screen, no matter how big, or where Form1 is
Cntrl.Top = (Screen.Height * 1#) / 2 - Cntrl.Height / 2
Cntrl.Left = (Screen.Width * 1#) / 2 - Cntrl.Width / 2
End Sub
```

