VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVT100Cache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Parallel Arrays (same dimensions)
Private masData()   As String   'String cache storage
Private mabValid()  As Boolean  'Flag that indicates if entry is valid (meaning has not been invalidated)

Private Sub Class_Initialize()
  'nothing
End Sub

Private Sub Class_Terminate()
  'nothing
End Sub

Public Property Get Size() As Integer
  On Error Resume Next
  Size = UBound(masData) - LBound(masData) + 1
  If Err.Number <> 0 Then Size = 0
End Property

Public Sub Clear()
  On Error Resume Next
  Erase masData
  Erase mabValid
End Sub

'We don't check boundaries here because we want max speed
Public Function GetLine(ByVal piLine As Integer) As String
  GetLine = masData(piLine)
End Function

Public Sub SetLine(ByVal piLine As Integer, ByVal psData As String)
  If (piLine < 1) Then Exit Sub
  If piLine > Me.Size Then
    ReDim Preserve masData(1 To piLine) As String
    ReDim Preserve mabValid(1 To piLine) As Boolean
  End If
  masData(piLine) = psData
  mabValid(piLine) = True
End Sub

Public Function IsCached(ByVal piLine As Integer) As Boolean
  If (piLine > 0) And (piLine <= Me.Size) Then
    IsCached = mabValid(piLine)
  End If
End Function

Public Sub InvalidateLine(ByVal piLine As Integer)
  If (piLine > 0) And (piLine <= Me.Size) Then
    mabValid(piLine) = False
  End If
End Sub
