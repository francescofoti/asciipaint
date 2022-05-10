VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IConsoul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Sub ConOut(ByVal psInfo As String)

End Sub

Public Sub ConOutLn(ByVal psInfo As String, Optional ByVal piQBColorText As Integer = -1)

End Sub

Public Property Get AutoRedraw() As Boolean

End Property

Public Property Let AutoRedraw(ByVal pfFlag As Boolean)

End Property

Public Sub RefreshWindow()

End Sub

Public Function GetConsoul() As CConsoul

End Function
