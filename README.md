<div align="center">

## CPause \(Class for pausing\)


</div>

### Description

This is a class that you can use to make your app pause for a few hours/minutes/seconds or milliseconds. This will work even if midnight occurrs while pausing! (Thus, this is midnight-compliant, just like my CStopwatch class I submitted earlier.)
 
### More Info
 
Function pPause(ByVal Number as Single, Optional Byval Unit as iConstants)

Number: the number of units, specified in Unit, that you want to pause

Unit: Contains one of the Following:

iMillisec ( = 3)

iSeconds ( = 0)

iMinutes ( = 1)

iHours  ( = 2)

NOTE: If Unit is omitted, Unit will equal iSeconds

Unfortunately, if you want to see if this can pause through midnight, you'll have to set your system time. Nothing serious. :)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jan Botha](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jan-botha.md)
**Level**          |Unknown
**User Rating**    |5.0 (30 globes from 6 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jan-botha-cpause-class-for-pausing__1-2497/archive/master.zip)





### Source Code

```
'*************************************
'This goes into a class module
'Important: NAME THE MODULE "CPause"
'*************************************
Const iSecsInDay As Long = 86400
Enum iConstants
  iSeconds = 0
  iMinutes = 1
  iHours = 2
  iMilliSec = 3
End Enum
Public Function pPause(ByVal Number As Single, _
     Optional ByVal Unit As iConstants)
  Dim iStopTime, fakeTimer, sAfterMidnight, sBeforeMidnight
  If Unit = iSeconds Then
    Number = Number
   ElseIf Unit = iMinutes Then
    Number = Number * 60
   ElseIf Unit = iHours Then
    Number = Number * 3600
   ElseIf Unit = iMilliSec Then
    Number = Number / 1000
  End If
  fakeTimer = Timer
  iStopTime = fakeTimer + Number
  If iStopTime > iSecsInDay Then
    sAfterMidnight = iStopTime - iSecsInDay
    sBeforeMidnight = Number - sAfterMidnight
    fakeTimer = Timer
    While Timer < fakeTimer + sBeforeMidnight And Timer <> 0
      DoEvents
    Wend
    fakeTimer = Timer
    While Timer < fakeTimer + sAfterMidnight
      DoEvents
    Wend
   Else 'if pausing won't continue through midnight
    While Timer < iStopTime
      DoEvents
    Wend
  End If
End Function
'************************************
'Put the following in the Declaration
'section of a form
'************************************
Dim mytimer as CPause
'***************************************************
'Put the following into any Sub (eg. Command1_Click)
'***************************************************
Set mytimer = New CPause
'to pause for 10 seconds, use the following call
i = mytimer.pPause(10, iSeconds)
'**************************************************
'End of Code
'I welcome any comments bug reports or enhancements that can be made!
'<c03jabot@prg.wcape.school.za>
```

