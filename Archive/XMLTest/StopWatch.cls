VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mStartTime As Long
Private mFinishTime As Long
Private mElapsedTime As Long

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Sub Start()
   mStartTime = timeGetTime
End Sub

Public Sub Finish()
   mFinishTime = timeGetTime
   mElapsedTime = mFinishTime - mStartTime
End Sub

Public Property Get ElaspsedTime()
   ElaspsedTime = mElapsedTime / 1000
End Property

Public Property Get CurrentTime()
   CurrentTime = (timeGetTime - mStartTime) / 1000
End Property


