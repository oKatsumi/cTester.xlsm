VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This is for Testing Class.

Option Explicit
Const LINE = "*******************"
Private xNG As Boolean
Private xTime As Double

Public Sub DoTest()
    Call StartSeq
    Call TestSuite
    Call EndSeq

End Sub
Private Sub TestSuite()
    
    Debug.Print "TestSample ..."
    Dim TestSample As New cSample_Test
    xNG = TestSample.TestPlas(xNG)
    
End Sub
Private Sub StartSeq()
    Dim i As Integer
    For i = 0 To 30
        Debug.Print "|" & vbLf
    Next
    Debug.Print LINE
    Debug.Print "Test Start !"
    Debug.Print LINE
    xTime = Timer

End Sub
Private Sub EndSeq()
    
    xTime = Timer - xTime
    Debug.Print ""
    
    If xNG = True Then
        Debug.Print LINE
        Debug.Print "(*3*)9m NG !!"
        Debug.Print LINE
    Else
        Debug.Print LINE
        Debug.Print "ALL Green"
        Debug.Print LINE
    End If
    Debug.Print "CalcTime : " & Format(xTime, "0.00") & " [s]"

End Sub
