VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IProgressIndicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Sub BeginProgress(ByVal psMessage As String)

End Sub

Public Sub EndProgress()

End Sub

Public Sub SetValue(ByVal plValue As Long)

End Sub

Public Sub SetMax(ByVal plMax As Long)

End Sub

Public Sub SetText(ByVal psText As String)

End Sub

Public Sub SetCaption(ByVal plCaption As String)

End Sub

Public Sub ShowProgressIndicator(ByVal pfShow As Boolean)

End Sub

Public Property Get Console() As CConsoul

End Property
