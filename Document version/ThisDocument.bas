Private Sub Document_Open()
    InitializerForm.Show False
    StartMyTimer 10
End Sub

Private Sub Document_Close()
    StopMyTimer
End Sub