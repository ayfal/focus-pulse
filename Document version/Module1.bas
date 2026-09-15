Attribute VB_Name = "Module1"
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function SetTimer Lib "user32" ( _
        ByVal hwnd As LongPtr, _
        ByVal nIDEvent As LongPtr, _
        ByVal uElapse As Long, _
        ByVal lpTimerFunc As LongPtr) As LongPtr

    Private Declare PtrSafe Function KillTimer Lib "user32" ( _
        ByVal hwnd As LongPtr, _
        ByVal nIDEvent As LongPtr) As Long
#Else
    Private Declare Function SetTimer Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal nIDEvent As Long, _
        ByVal uElapse As Long, _
        ByVal lpTimerFunc As Long) As Long

    Private Declare Function KillTimer Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal nIDEvent As Long) As Long
#End If

Public TimerID As LongPtr
Public UsedCounter As Date
Public WastedCounter As Date
Public IsWorking As Boolean
Public StartedAt As Date

Public Sub TimerCallback(ByVal hwnd As LongPtr, ByVal uMsg As Long, _
                         ByVal idEvent As LongPtr, ByVal dwTime As Long)
    If IsWorking Then
        UsedCounter = UsedCounter + TimeSerial(0, InitializerForm.Minutes.Text, 0)
        CreateObject("WScript.Shell").Popup genMsg("Well done, time's up! Go back to your tasklist to reschedule this task"), 0, "To-Do Reminder", 4096
    Else
        WastedCounter = WastedCounter + TimeSerial(0, InitializerForm.Minutes.Text, 0)
        InitializerForm.BackToDoc
        Selection.InsertBefore "1111-11-11 11:11 U " & InputBox(genMsg("Add whatever you're doing now to the tasklist. Don't worry, you'll get back to it soon"), "Input Required") & vbCr
        InitializerForm.RescheduleButton_Click
    End If
   IsWorking = False
   InitializerForm.UntilLabel = "Get to work!"
End Sub
'ToDo change the reschedule to happen on timer callback. (or on start???)
'cases:
'1. while isworking (finished early): give me option to update then reschedule and restart. this is wrong, as there isn't a timer callback when i start the doc
'2. while not isworking (stopped wroking after the message): give me option to update then reschedule and restart
'3. while not is working (just started the doc): start. this is wrong, as there isn't a timer callback when i start the doc
'4. while not is working (slacking off): give option to make a new task then reschedule and restart

Function genMsg(m As String) As String
    genMsg = m & vbCrLf & _
        "So far:" & vbCrLf & _
        "Time Used: " & format(UsedCounter, "hh:mm:ss") & vbCrLf & _
        "Time Wasted: " & format(WastedCounter, "hh:mm:ss")
End Function

Sub StartMyTimer(Minutes As Double)
    TimerID = SetTimer(0, 0, CLng(1000# * 60# * Minutes), AddressOf TimerCallback)
End Sub

Sub StopMyTimer()
    If TimerID <> 0 Then
        KillTimer 0, TimerID
        TimerID = 0
    End If
End Sub

