VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InitializerForm 
   Caption         =   "Welcome"
   ClientHeight    =   5475
   ClientLeft      =   -15
   ClientTop       =   30
   ClientWidth     =   10470
   OleObjectBlob   =   "InitializerForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InitializerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' this is an app to organize tasks in a ms-word document
' tasks are prioritized by their due date and time
' every task is a paragraph in the document, and it starts with a due date and time
' the app alerts the user to switch tasks every few minutes
' and the user reschedules his current task

Sub start_click()
    ' Input validation
    If Minutes.Text = "" Or Not IsNumeric(Minutes.Text) Then
        MsgBox "Please enter a numeric value.", vbExclamation
        Exit Sub
    End If
    UntilLabel = "Work until " & Now + TimeSerial(0, Minutes.Text, 0)
    If IsWorking Then
        UsedCounter = UsedCounter + 1
        CreateObject("WScript.Shell").Popup "Well done!" & vbCrLf & _
         "So far:" & vbCrLf & _
         "• Time Used: " & UsedCounter & vbCrLf & _
         "• Time Wasted: " & WastedCounter, _
         0, "To-Do Reminder", 4096
    End If
    StopMyTimer
    IsWorking = True
    StartMyTimer (Minutes.Text)
End Sub

Private Sub Reschedule(Mnts As Integer)
    ' if the first task isn't scheduled then give it dummy schedule
    Selection.start = 0
    Selection.End = 16
    If Not IsDate(Selection.Text) Then Selection.InsertBefore format(Now, "yyyy-mm-dd hh:mm") + " "
    
    ' reschedule
    Selection.End = 16
    Selection = format(DateAdd("n", Int((Now - CDate(Selection)) * 60 * 24 / Mnts + 1) * Mnts, Selection), "yyyy-mm-dd hh:mm")
    
    ' sort all the paragraphs
    Selection.WholeStory
    Selection.Sort
    Selection.End = 0
    
    ' save the document
    ActiveDocument.Save
End Sub

Private Sub TaskButton_Click()
    Reschedule Minutes.Text
End Sub

Private Sub DayButton_Click()
    Reschedule 60 * 24
End Sub

Private Sub WeekButton_Click()
    Reschedule 60 * 24 * 7
End Sub

Private Sub UserForm_Initialize()
    Me.Width = 293
    Me.Height = 152
    Me.StartUpPosition = 1 ' CenterOwner or 0 for manual
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    StopMyTimer
End Sub

