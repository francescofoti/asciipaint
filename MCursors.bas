Attribute VB_Name = "MCursors"
Option Compare Database
Option Explicit

' Standard Cursor IDs
Public Const IDC_ARROW = 32512&
Public Const IDC_IBEAM = 32513&
Public Const IDC_WAIT = 32514&
Public Const IDC_CROSS = 32515&
Public Const IDC_UPARROW = 32516&
Public Const IDC_SIZE = 32640&
Public Const IDC_ICON = 32641&
Public Const IDC_SIZENWSE = 32642&
Public Const IDC_SIZENESW = 32643&
Public Const IDC_SIZEWE = 32644&
Public Const IDC_SIZENS = 32645&
Public Const IDC_SIZEALL = 32646&
Public Const IDC_NO = 32648&
Public Const IDC_HAND = 32649&
Public Const IDC_APPSTARTING = 32650&

#If Win64 Then
Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As LongPtr, ByVal lpCursorName As LongPtr) As LongPtr
Declare PtrSafe Function CreateCursor Lib "user32" (ByVal hInstance As LongPtr, ByVal nXhotspot As Long, ByVal nYhotspot As Long, ByVal nWidth As Long, ByVal nHeight As Long, lpANDbitPlane As Any, lpXORbitPlane As Any) As LongPtr
Declare PtrSafe Function DestroyCursor Lib "user32" (ByVal hCursor As LongPtr) As Long
Declare PtrSafe Function CopyCursor Lib "user32" (ByVal hcur As LongPtr) As LongPtr
#Else
Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Declare Function CreateCursor Lib "user32" (ByVal hInstance As Long, ByVal nXhotspot As Long, ByVal nYhotspot As Long, ByVal nWidth As Long, ByVal nHeight As Long, lpANDbitPlane As Any, lpXORbitPlane As Any) As Long
Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Declare Function CopyCursor Lib "user32" (ByVal hcur As Long) As Long
#End If

#If Win64 Then
Public Function WinGetCursorArrow() As LongPtr
#Else
Public Function WinGetCursorArrow() As Long
#End If
  WinGetCursorArrow = LoadCursor(0, IDC_ARROW)
End Function

#If Win64 Then
Public Function WinGetCursorHand() As LongPtr
#Else
Public Function WinGetCursorHand() As Long
#End If
  WinGetCursorHand = LoadCursor(0, IDC_HAND)
End Function

#If Win64 Then
Public Function WinGetCursor(ByVal pIDCCursor As LongPtr) As LongPtr
#Else
Public Function WinGetCursor(ByVal pIDCCursor As Long) As Long
#End If
  WinGetCursor = LoadCursor(0, pIDCCursor)
End Function

#If Win64 Then
Public Function WinDestroyCursor(ByVal phCursor As LongPtr) As Long
#Else
Public Function WinDestroyCursor(ByVal phCursor As Long) As Long
#End If
  WinDestroyCursor = DestroyCursor(phCursor)
End Function


