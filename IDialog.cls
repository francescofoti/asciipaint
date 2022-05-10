VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Property Get DialogID() As String

End Property

Public Function ShowDialog(ByVal pfShowModal As Boolean) As Boolean

End Function

Public Property Get IsModal() As Boolean

End Property

Public Property Get Cancelled() As Boolean

End Property

Public Property Let Cancelled(ByVal pfCancelled As Boolean)

End Property
