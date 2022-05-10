VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICsWmPaintEventSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Return 1 from implementation to stop broadcast
Public Function OnConsolePaint( _
  ByVal phWnd As Long, _
  ByVal pwCbkMode As Integer, _
  ByVal phDC As LongPtr, _
  ByVal lprcLinePos As LongPtr, _
  ByVal lprcLineRect As LongPtr, _
  ByVal lprcPaint As LongPtr) As Integer

End Function



