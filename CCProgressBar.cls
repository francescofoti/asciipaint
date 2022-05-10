VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private miLine      As Integer

Public Value        As Long
Public Caption      As String
Public Text         As String
Public Max          As Long
Public CharWidth    As Integer
Public EnsureVisible  As Boolean

Private Sub Class_Initialize()
  Me.CharWidth = 40
  Me.Max = 100
End Sub

Private Sub Class_Terminate()
  '...
End Sub

Public Sub Render(ByRef pconRender As CConsoul)
  If miLine = 0 Then
    'output the line, keep the id
    miLine = pconRender.OutputLn("")
  End If
  If Me.EnsureVisible Then
    If (miLine < pconRender.TopLine) Or (pconRender.LineCount = 0) Then
      miLine = pconRender.OutputLn("")  'render on a new line
    End If
  End If
  
  Dim iValueCharCt      As Integer
  Dim sPct              As String
  
  On Error Resume Next 'this computation can trigger an overflow error
  iValueCharCt = (CLng(Me.Value) * CLng(Me.CharWidth)) / Me.Max
  
  sPct = VT_FCOLOR(QBColor(QBCOLOR_GREEN))
  sPct = sPct & String$(iValueCharCt, ChrW$(&H2588))
  sPct = sPct & VT_RESET()
  sPct = sPct & String$(Me.CharWidth - iValueCharCt, ChrW$(&H2591))
  If Len(Me.Caption) > 0 Then
    sPct = Me.Caption & " " & sPct
  End If
  If Len(Me.Text) Then
    sPct = sPct & " " & Me.Text
  End If
  pconRender.SetLine miLine, sPct
End Sub

