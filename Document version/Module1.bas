Attribute VB_Name = "Module1"
#If VBA7 Then
    Private Declare PtrSafe Function SetTimer Lib "user32" ( _
        ByVal hWnd As LongPtr, _
        ByVal nIDEvent As LongPtr, _
        ByVal uElapse As Long, _
        ByVal lpTimerFunc As LongPtr) As LongPtr

    Private Declare PtrSafe Function KillTimer Lib "user32" ( _
        ByVal hWnd As LongPtr, _
        ByVal nIDEvent As LongPtr) As Long
    
#Else
    Private Declare Function SetTimer Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal nIDEvent As Long, _
        ByVal uElapse As Long, _
        ByVal lpTimerFunc As Long) As Long

    Private Declare Function KillTimer Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal nIDEvent As Long) As Long
        
#End If

Public TimerID As LongPtr
Public UsedCounter As Integer
Public WastedCounter As Integer
Public IsWorking As Boolean

Public Sub TimerCallback(ByVal hWnd As LongPtr, ByVal uMsg As Long, _
                         ByVal idEvent As LongPtr, ByVal dwTime As Long)
    If IsWorking Then
        UsedCounter = UsedCounter + 1
        msg = "Well done"
    Else
        WastedCounter = WastedCounter + 1
        msg = "Too bad"
    End If
    CreateObject("WScript.Shell").Popup msg & ", time's up! Go Back to Your tasks list to reschedule this task" & vbCrLf & _
         "So far:" & vbCrLf & _
         "• Time Used: " & UsedCounter & vbCrLf & _
         "• Time Wasted: " & WastedCounter, _
         0, "To-Do Reminder", 4096
   IsWorking = False
   InitializerForm.UntilLabel = "Get to work!"
End Sub

Sub StartMyTimer(Minutes As Double)
    TimerID = SetTimer(0, 0, CLng(1000# * 60# * Minutes), AddressOf TimerCallback)
End Sub

Sub StopMyTimer()
    If TimerID <> 0 Then
        KillTimer 0, TimerID
        TimerID = 0
    End If
End Sub



