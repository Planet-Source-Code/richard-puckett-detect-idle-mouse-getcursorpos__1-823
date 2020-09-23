<div align="center">

## Detect Idle Mouse \(GetCursorPos\)


</div>

### Description

This code can be used to determine the X,Y coordinates of the mouse cursor

and use them to check for idle mouse activity. This code is useful in that it does

not require your current form to be in focus (active windows status). The

GetCursorPos can be used in conjunction with or be replaced by another API

call GetCaretPos, which determines the X,Y coordinates of the text cursor.

Hopefully this will be useful to anyone looking to check for an idle desktop.

(Richard Puckett, puckettr@mindspring.com)
 
### More Info
 
1. Take out all of the ME.PRINT statements, since they are only to illustrate how

the function works (I used this code in a login program to monitor mapped

network drive connections in lab environments.)  2. All of the declarations are

in a module. 3. This example is only using the X coords to determine activity, I

am sure a more complex method can be devised.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Richard Puckett](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/richard-puckett.md)
**Level**          |Unknown
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/richard-puckett-detect-idle-mouse-getcursorpos__1-823/archive/master.zip)

### API Declarations

```
'API Call establishes mouse coords
Public Declare Function GetCursorPos Lib "user32" _
(lpPoint As PointAPI) As Long
Public Pnt As PointAPI
'These values MUST be public
Public OldX As Long
Public OldY As Long
Public NewX As Long
Public NewY As Long
Public Type PointAPI
    X As Long
    Y As Long
End Type
'This Const determines the total timeout value in minutes
Global Const MINUTES = 15
Public TimeExpired
Public ExpiredMinutes
```


### Source Code

```

Public Sub Form_Load()
  Timer1.Interval = 1000
  OldX = 0
  OldY = 0
End Sub
Public Sub Timer1_Timer()
  GetCursorPos Pnt
    Me.Cls
    Me.Print "The current mouse coordinates are "; _
    Pnt.X; ","; Pnt.Y
  NewX = Pnt.X
  NewY = Pnt.Y
    Me.Print "OldX coords", OldX
    Me.Print "OldY coords", OldY
    Me.Print "NewX coords", NewX
    Me.Print "NewY coords", NewY
    If OldX - NewX = 0 Then
      Me.Print "No Movement Detected"
      TimeExpired = TimeExpired + Timer1.Interval
      Me.Print "Total Time Expired", TimeExpired
    Else
      Me.Print "Mouse is Moving"
      TimeExpired = 0
    End If
  OldX = NewX
  OldY = NewY
    ExpiredMinutes = (TimeExpired / 1000) / 60
    If ExpiredMinutes >= MINUTES Then
    TimeExpired = 0
    Me.Print "Times Up!!!"
    End If
End Sub
```

