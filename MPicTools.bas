Attribute VB_Name = "MPicTools"
Option Compare Database
Option Explicit

Private Declare Function CLSIDFromString& Lib "ole32" (ByVal lpsz As Any, pclsid As Any)
Private Declare Function CreateStreamOnHGlobal& Lib "ole32" (ByVal hGlobal&, ByVal fDeleteOnRelease&, ppstm As Any)
Private Declare Function OleLoadPicture& Lib "olepro32" (pStream As Any, ByVal lSize&, ByVal fRunmode&, rIID As Any, ppvObj As Any)
#If Win64 Then
Private Declare PtrSafe Function apiDeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As LongPtr) As Long
#Else
Private Declare PtrSafe Function apiDeleteObject Lib "gdi32" Alias "DeleteObject" (ByVal hObject As Long) As Long
#End If

Private Declare Sub OleCreatePictureIndirect Lib "oleaut32.dll" (ByRef lpPICTDESC As PICTDESC, ByVal rIID As Long, ByVal fOwn As Long, ByVal lplpvObj As Long)
Private Type PICTDESC
   cbSizeOfStruct As Long
   picType As Long
   hGDIObj As Long
   hPalOrXYExt As Long 'this member isn't used for icons, but is for bitmaps
End Type
Private Type IID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(0 To 7) As Byte
End Type
Private Const hNull As Long = 0

#If Win64 Then
Public Function DeleteBitmapObject(ByVal hBitmap As LongPtr) As Long
#Else
Public Function DeleteBitmapObject(ByVal hBitmap As Long) As Long
#End If
  DeleteBitmapObject = apiDeleteObject(hBitmap)
End Function

Public Function BitmapToPicture(ByVal hBitmap As Long) As IPicture
  Dim ipic As IPicture
  Dim picdes As PICTDESC
  Dim iidIPicture As IID
  If hBitmap = hNull Then Exit Function
  'Fill picture description
  With picdes
    .cbSizeOfStruct = Len(picdes)
    .picType = 1 'PICTYPE_BITMAP, from https://docs.microsoft.com/en-us/windows/win32/com/pictype-constants
    .hGDIObj = hBitmap
  End With
  'Fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
  With iidIPicture
    .Data1 = &H7BF80980
    .Data2 = &HBF32
    .Data3 = &H101A
    .Data4(0) = &H8B
    .Data4(1) = &HBB
    .Data4(2) = &H0
    .Data4(3) = &HAA
    .Data4(4) = &H0
    .Data4(5) = &H30
    .Data4(6) = &HC
    .Data4(7) = &HAB
  End With
  'Create picture from icon handle
  OleCreatePictureIndirect picdes, VarPtr(iidIPicture), True, VarPtr(ipic)
  'Result will be valid a valid IPicture reference or Nothing
  Set BitmapToPicture = ipic
End Function
    
Private Function LoadFromByteStream(B() As Byte) As StdPicture
  Static IPicture(0 To 15) As Byte
  If IPicture(0) = 0 Then CLSIDFromString StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IPicture(0)
 
  Dim iStream As stdole.IUnknown
  If CreateStreamOnHGlobal(VarPtr(B(LBound(B))), 0, iStream) Then Exit Function
  OleLoadPicture ByVal ObjPtr(iStream), UBound(B) - LBound(B) + 1, 0, IPicture(0), LoadFromByteStream
End Function
