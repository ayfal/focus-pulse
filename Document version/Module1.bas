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
Public IsTimerRunning As Boolean

Public Sub TimerCallback(ByVal hwnd As LongPtr, ByVal uMsg As Long, _
                         ByVal idEvent As LongPtr, ByVal dwTime As Long)
    If IsWorking Then
        UsedCounter = UsedCounter + Now - StartedAt
        CreateObject("WScript.Shell").Popup "Well done, time's up! Go back to your tasklist to reschedule this task", 0, "To-Do Reminder", 4096
    Else
        WastedCounter = WastedCounter + Now - StartedAt
        InitializerForm.BackToDoc
        Selection.InsertBefore "1111-11-11 11:11 U " & InputBox("Add whatever you're doing now to the tasklist. Don't worry, you'll get back to it soon", "Input Required") & vbCr
        InitializerForm.RescheduleButton_Click
    End If
    IsWorking = False   
End Sub

Sub StartMyTimer(Minutes As Double)
    TimerID = SetTimer(0, 0, CLng(1000# * 60# * Minutes), AddressOf TimerCallback)
    StartedAt = Now
    IsTimerRunning = True
    UpdateLabels
End Sub

Sub StopMyTimer()
    If TimerID <> 0 Then
        KillTimer 0, TimerID
        TimerID = 0
    End If
    IsTimerRunning = False
End Sub

Sub UpdateLabels()
    If IsTimerRunning Then
        If IsWorking Then
            InitializerForm.UsedLabel = format(UsedCounter + Now - StartedAt, "hh:mm:ss")
        Else
            InitializerForm.WastedLabel = format(WastedCounter + Now - StartedAt, "hh:mm:ss")
        End If
        InitializerForm.TimerLabel = format(Now - StartedAt, "hh:mm:ss")

        Application.OnTime Now + TimeValue("00:00:01"), "UpdateLabels"
    End If
End Sub
        

