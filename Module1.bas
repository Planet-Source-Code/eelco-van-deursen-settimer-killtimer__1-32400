Attribute VB_Name = "Module1"
' TIMER FUNCTION
Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal nIDEvent As Long, ByVal dwTime As Long)
    ' EACH TIMER TRIGGERS ON DIFFERENT INTERVALS
    ' nIDEvent CONTAINS THE TIMER'S EVENTID THAT IS TRIGGERING
    ' CAN ALSO USE THIS CONSTRUCTION
    ' Select Case nIDEvent
    ' Case 1
    '   TODO : TIMER1 CODE HERE
    ' Case 2
    '   TODO : TIMER2 CODE HERE
    ' ..
    ' ..
    ' Case 19
    '   TODO : TIMER19 CODE HERE
    ' Case 20
    '   TODO : TIMER20 CODE HERE
    ' End Select
    
    ' FILL BACKCOLOR OF TRIGGERED TIMER WITH RANDOM COLOR
    Form1.Label1(nIDEvent - 1).BackColor = RGB(Int(255 * Rnd()), Int(255 * Rnd()), Int(255 * Rnd()))
End Sub
