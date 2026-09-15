VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InitializerForm 
   Caption         =   "Welcome"
   ClientHeight    =   3030
   ClientLeft      =   -15
   ClientTop       =   30
   ClientWidth     =   5865
   OleObjectBlob   =   "InitializerForm.frx":0000
   ShowModal       =   0   'False
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
#If VBA7 Then
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
#Else
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
#End If
Public SourceDoc As Word.Document

Sub start_click()
    ' Input validation
    If Minutes.Text = "" Or Not IsNumeric(Minutes.Text) Then
        MsgBox "Please enter a numeric value.", vbExclamation
        Exit Sub
    End If
    UntilLabel = "Work until " & Now + TimeSerial(0, Minutes.Text, 0)
    If IsWorking Then
        UsedCounter = UsedCounter + Now - StartedAt
        CreateObject("WScript.Shell").Popup "Well done!" & vbCrLf & _
         "So far:" & vbCrLf & _
         "Time Used: " & format(UsedCounter, "hh:mm:ss") & vbCrLf & _
         "Time Wasted: " & format(WastedCounter, "hh:mm:ss"), _
         0, "To-Do Reminder", 4096
    End If
    StopMyTimer
    IsWorking = True
    StartedAt = Now
    StartMyTimer (Minutes.Text)
    BackToDoc
End Sub

Public Sub RescheduleButton_Click()
   With SourceDoc
        Dim r As Range
    
        ' First 16 characters
        Set r = .Range(0, 16)
    
        ' If the first task isn't scheduled then give it dummy schedule
        If Not IsDate(r.Text) Then
            r.InsertBefore format(Now, "yyyy-mm-dd hh:mm") & " "
        End If
        
        ' Determine interval according to the frequency flag (the 18th character)
        Select Case .Range.Words(9)
            Case "D ": Mnts = 60 * 24 ' Daily
            Case "W ": Mnts = 60 * 24 * 7 ' Weekly
            Case "U " ' Urgent
                .Range(0, 16).Text = "1111-11-11 11:11"
                Selection.HomeKey Unit:=wdStory
                Do While Selection.Next(Unit:=wdParagraph, Count:=1).Words(9) = "U "
                    Selection.Range.Relocate wdRelocateDown
                Loop
                Do
                    Selection.Range.Relocate wdRelocateDown
                Loop While Selection.Next(Unit:=wdParagraph, Count:=1).Words(9) = "U "
                GoTo Save
            Case Else: Mnts = CLng(Minutes.Text) ' One time task
        End Select
        
        ' Reschedule
        Set r = .Range(0, 16)
        r.Text = format(DateAdd("n", Int((Now - CDate(r.Text)) * 60 * 24 / Mnts + 1) * Mnts, r.Text), "yyyy-mm-dd hh:mm")
    
        ' Sort all paragraphs
        .Range.WholeStory
        .Range.Sort SortFieldType:=wdSortFieldAlphanumeric
    
Save:
        .Save
        BackToDoc
    End With
End Sub

Private Sub UserForm_Initialize()
    Me.Width = 305
    Me.Height = 180
    Me.StartUpPosition = 1 ' CenterOwner or 0 for manual
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    StopMyTimer
End Sub

Public Sub BackToDoc()
    SourceDoc.Activate
    ' Return to the document
     SetForegroundWindow Word.Application.ActiveWindow.hwnd
    ' Equivalent to pressing Ctrl + Home on your keyboard
    Selection.HomeKey Unit:=wdStory
End Sub

