<div align="center">

## Almost all the Different MsgBox Constants and howto use them


</div>

### Description

This is really helpful if you're making a program that asks the user a yes/no question with a message box. That's only one of many great examples. Please Rate.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jason Ryczek](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jason-ryczek.md)
**Level**          |Beginner
**User Rating**    |3.9 (27 globes from 7 users)
**Compatibility**  |VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jason-ryczek-almost-all-the-different-msgbox-constants-and-howto-use-them__1-22382/archive/master.zip)





### Source Code

```
' This works best if you make an array and use the select case statement, but I left it up to the user to do what they want with it.
Dim x as integer ' x = answer value
' For the OKOnly
x = MsgBox("This is Default, Okay Only", vbOKOnly)
' For Ok, cancel
x = MsgBox("Give choice of OK or Cancel", vbOKCancel)
' This is Abort, Retry, Ignore
x = MsgBox("Blow up your computer?", vbAboutRetryIgnore)
' Most Common - :)
x = MsgBox("Would you like to save your dirty pictures before you exit?", vbYesNoCancel)
' Yes or No
x = MsgBox("Was that you or the dog?", vbYesNo)
' Retry or Cancel
x = MsgBox("This is what you might want to do with your diet...", vbRetryCancel)
' Critical Message
x = MsgBox("Your Reports due today and you haven't started on it yet!", vbCritical)
' Question
x = MsgBox("<- Question", vbQuestion)
' now, do decypher the answers
Select Case x
 Case 1:MsgBox "The Okay button"
 Case 2:MsgBox "The Cancel button"
 Case 3:MsgBox "The Abort button"
 Case 4:MsgBox "The Retry button"
 Case 5:MsgBox "The Ignore button"
 Case 6:MsgBox "The Yes button"
 Case 7:MsgBox "The No button"
End Select
' Well, hope that helps all you programmers out there
```

