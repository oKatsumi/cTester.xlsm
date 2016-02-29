VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSample_Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function TestPlas(ByRef NG As Boolean)
    Dim mySample As New cSample
    If 3 = mySample.myPlas(1, 2) Then
    Else
        Debug.Print "NG : Sample.myPlas_1"
        NG = True
    End If
    If 7 = mySample.myPlas(5, 2) Then
    Else
        Debug.Print "NG : Sample.myPlas_2"
        NG = True
    End If
    TestPlas = NG
End Function
