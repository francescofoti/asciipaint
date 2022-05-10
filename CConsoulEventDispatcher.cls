VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CConsoulEventDispatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' We receive mouse events via a sub in a module.
' These class registers ICsMouseEventSink target interface for hWnds
' and sends them to the registered ICsMouseEventSink target

'The event's sinks list hold interface pointers for each type
'of event that we'll dispatch for an hwnd
Private mlstEventSinks As CList

'Those are bit flags
Public Enum eConsoulEventType
  eCsUnsupportedEventType = 0
  eCsMouseEvent       'ICsMouseEventSink
  eCsVirtualLineEvent 'ICsVirtualLineEventSink
  eCsWmPaint          'ICsWmPaintEventSink
End Enum

Private Sub Class_Initialize()
  Set mlstEventSinks = New CList
  'We'll store the hwnd as a double, as in 64bits this is a LongPrt (or LongLong) and CList/CRow classes
  'do not know of these "new" data types (will fix that one day)
  mlstEventSinks.ArrayDefine Array("hwnd", "ICsMouseEventSink", "ICsVirtualLineEventSink", "ICsWmPaintEventSink"), _
                             Array(vbDouble, vbObject, vbObject, vbObject)
End Sub

Private Sub Class_Terminate()
  Set mlstEventSinks = Nothing
End Sub

'We can register one interface at a time for an hwnd, ie
'we can call this method multiple times for an hwnd, one for
'interface we want to register for it.
'We must receive here the type of event from the caller, because we cannot
'determine and choose the interface type from pIIEventSink ourselves.
'UnregisterEventSink unregisters them all.
Public Sub RegisterEventSink(ByVal phWnd As Long, ByRef pIIEventSink As Object, ByVal peEventType As eConsoulEventType)
  Dim iHwnd       As Long
  
  If peEventType <> eCsUnsupportedEventType Then
    iHwnd = mlstEventSinks.Find("hwnd", phWnd)
    If iHwnd = 0 Then
      iHwnd = mlstEventSinks.AddValues(phWnd, Nothing, Nothing, Nothing)
    End If
    If peEventType = eCsMouseEvent Then
      Set mlstEventSinks("ICsMouseEventSink", iHwnd) = pIIEventSink
    ElseIf peEventType = eCsVirtualLineEvent Then
      Set mlstEventSinks("ICsVirtualLineEventSink", iHwnd) = pIIEventSink
    ElseIf peEventType = eCsWmPaint Then
      Set mlstEventSinks("ICsWmPaintEventSink", iHwnd) = pIIEventSink
    Else
      Err.Raise 13, "RegisterEventSink", "Interface type is not supported"
    End If
  Else
    Err.Raise 13, "RegisterEventSink", "Interface type is not supported"
  End If
End Sub

Public Sub UnregisterEventSink(phWnd)
  Dim iHwnd       As Long
  iHwnd = mlstEventSinks.Find("hwnd", phWnd)
  If iHwnd > 0 Then
    mlstEventSinks.Remove iHwnd
  End If
End Sub

Public Sub BroadcastMouseEvent(ByVal phWnd As Long, ByVal piEvtCode As Integer, ByVal pwParam As Integer, ByVal piZoneID As Integer, ByVal piRow As Integer, ByVal piCol As Integer, ByVal piPosX As Integer, ByVal piPosY As Integer)
  Dim iiEventSink As ICsMouseEventSink
  Dim iResponse   As Integer
  Dim i           As Long
  On Error Resume Next
  For i = 1 To mlstEventSinks.Count
    If mlstEventSinks("hwnd", i) = phWnd Then
      Set iiEventSink = mlstEventSinks("ICsMouseEventSink", i)
      If Not iiEventSink Is Nothing Then
        iResponse = iiEventSink.OnMouseButton(phWnd, piEvtCode, pwParam, piZoneID, piRow, piCol, piPosX, piPosY)
        If CBool(iResponse) = True Then Exit For  'Called interface implementor handled the message and stopped propagation
        Set iiEventSink = Nothing
      End If
    End If
  Next i
End Sub

Public Sub BroadcastVirtualLineEvent(ByVal phWnd As Long, ByVal piLine As Long)
  Dim iiEventSink As ICsVirtualLineEventSink
  Dim iResponse   As Integer
  Dim i           As Long
  On Error Resume Next
  For i = 1 To mlstEventSinks.Count
    If mlstEventSinks("hwnd", i) = phWnd Then
      Set iiEventSink = mlstEventSinks("ICsVirtualLineEventSink", i)
      If Not iiEventSink Is Nothing Then
        iResponse = iiEventSink.OnVirtualLineGetData(phWnd, piLine)
        If CBool(iResponse) = True Then Exit For  'Called interface implementor handled the message and stopped propagation
        Set iiEventSink = Nothing
      End If
    End If
  Next i
End Sub

Public Sub BroadcastConsolePaint(ByVal phWnd As Long, ByVal pwCbkMode As Integer, ByVal phDC As Long, ByVal lprcLinePos As Long, ByVal lprcLineRect As Long, ByVal lprcPaint As Long)
  Dim iiEventSink As ICsWmPaintEventSink
  Dim iResponse   As Integer
  Dim i           As Long
  On Error Resume Next
  For i = 1 To mlstEventSinks.Count
    If mlstEventSinks("hwnd", i) = phWnd Then
      Set iiEventSink = mlstEventSinks("ICsWmPaintEventSink", i)
      If Not iiEventSink Is Nothing Then
        iResponse = iiEventSink.OnConsolePaint(phWnd, pwCbkMode, phDC, lprcLinePos, lprcLineRect, lprcPaint)
        If CBool(iResponse) = True Then Exit For  'Called interface implementor handled the message and stopped propagation
        Set iiEventSink = Nothing
      End If
    End If
  Next i
End Sub

