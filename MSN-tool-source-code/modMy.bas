Attribute VB_Name = "modMy"
Public Sub Wait(WaitTime)
    Dim StartTime As Double
    StartTime = Timer
    Do While Timer < StartTime + WaitTime
        If Timer > 86395 Or Timer = 0 Then Exit Do
        DoEvents
    Loop
End Sub
