Attribute VB_Name = "Module1"

Public Function Pause(seconds As Integer)
On Error Resume Next: Dim dTimer: Dim dTimer2 As Double
dTimer2 = Timer
Do While Timer < dTimer2 + seconds / 1000
DoEvents
Loop
End Function
Public Function RandomString(cb As Integer) As String

    Randomize
    Dim random3 As Long
    Dim rgch As String
    rgch = "abcdefghijklmnopqrstuvwxyz"
    random3 = Int((Rnd * 89) + 10)
    Dim i As Long
    For i = 1 To cb
        RandomString = RandomString & Mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
    Next
RandomString = RandomString & Str(random3)
RandomString = Replace(RandomString, " ", "")
End Function
