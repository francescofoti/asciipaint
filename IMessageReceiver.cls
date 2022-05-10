VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMessageReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Property Get ClientID() As String

End Property

Public Function OnMessageReceived(ByVal psSenderID As String, ByVal psTopic As String, ByRef pvData As Variant) As Long

End Function
