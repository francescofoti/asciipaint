VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICsMouseEventSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Return 1 from implementation to stop broadcast
Public Function OnMouseButton( _
  ByVal phWnd As Long, _
  ByVal piEvtCode As Integer, _
  ByVal pwParam As Integer, _
  ByVal piZoneID As Integer, _
  ByVal piRow As Integer, _
  ByVal piCol As Integer, _
  ByVal piPosX As Integer, _
  ByVal piPosY As Integer) As Integer

End Function

