VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CColorPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mlCount     As Long
Private malColor()  As Long
Private mlArraySize As Long

'V02.00.00
Private mfDirty           As Boolean  'if something changes, exposed via Dirty property

' IClassError implementation
Implements IClassError
Private mlErrNo   As Long
Private msErrCtx  As String
Private msErrDesc As String

Private Sub ClearErr()
  mlErrNo = 0&
  msErrCtx = ""
  msErrDesc = ""
End Sub

Private Sub SetErr(ByVal psErrCtx As String, ByVal plErrNum As Long, ByVal psErrDesc As String)
  mlErrNo = plErrNum
  msErrCtx = psErrCtx
  msErrDesc = psErrDesc
End Sub

Public Property Get LastErr() As Long
  LastErr = mlErrNo
End Property

Public Property Get LastErrDesc() As String
  LastErrDesc = msErrDesc
End Property

Public Property Get IIClassError() As IClassError
  Set IIClassError = Me
End Property

Private Property Get IClassError_LastErr() As Long
  IClassError_LastErr = mlErrNo
End Property

Private Property Get IClassError_LastErrCtx() As String
  IClassError_LastErrCtx = msErrCtx
End Property

Private Property Get IClassError_LastErrDesc() As String
  IClassError_LastErrDesc = msErrDesc
End Property

'
' Properties and member functions
'

Public Property Get Dirty() As Boolean
  Dirty = mfDirty
End Property

Public Property Get Count() As Long
  Count = mlCount
End Property

Public Property Get Color(ByVal plIndex As Long) As Long
  If (plIndex > 0) And (plIndex <= mlCount) Then
    Color = malColor(plIndex)
  End If
End Property

Public Property Let Color(ByVal plIndex As Long, ByVal plNewColor As Long)
  If (plIndex > 0) And (plIndex <= mlCount) Then
    malColor(plIndex) = plNewColor
    mfDirty = True
  End If
End Property

Public Sub Clear()
  mlCount = 0
  mfDirty = False
End Sub

Public Function AddColor(ByVal plColor As Long) As Integer
  If ColorIndex(plColor) > 0 Then
    'Debug.Print "Color " & Hex$(plColor) & " already exists"
    AddColor = 0
    Exit Function
  End If
  
  On Error GoTo AddColor_Err
  ClearErr
  
  If (mlCount + 1) > mlArraySize Then
    If mlArraySize > 0 Then
      mlArraySize = UBound(malColor)
    Else
      mlArraySize = 8
    End If
    ReDim Preserve malColor(1 To mlArraySize * 2)
    mlArraySize = UBound(malColor)
  End If
  mlCount = mlCount + 1
  malColor(mlCount) = plColor

  AddColor = mlCount
  mfDirty = True
  
AddColor_Exit:
  Exit Function

AddColor_Err:
  SetErr "AddColor", Err.Number, Err.Description
  Resume AddColor_Exit
End Function

Public Function ColorIndex(ByVal plColor As Long) As Long
  Dim i     As Long
  For i = 1 To mlCount  'simple seq search
    If malColor(i) = plColor Then
      ColorIndex = i
      Exit Function
    End If
  Next i
End Function

Public Function DeleteColor(ByVal plIndex As Long) As Boolean
  If (plIndex < 1) Or (plIndex > mlCount) Then
    SetErr "DeleteColor", 9&, Error$(9) 'index out of bounds
    Exit Function
  End If
  
  Dim i     As Long
  
  For i = plIndex To mlCount - 1
    malColor(i) = malColor(i + 1)
  Next i
  mlCount = mlCount - 1
  mfDirty = True
  
  DeleteColor = True
End Function

Public Function LoadFromFile(ByVal psFilename As String, ByVal pfMerge As Boolean) As Boolean
  On Error GoTo LoadFromFile_Err
  ClearErr
  
  Dim fh        As Integer
  Dim fIsOpen   As Boolean
  Dim lColor    As Long
  Dim iCount    As Long
  Dim i         As Long
  
  fh = FreeFile
  Open psFilename For Binary Access Read As #fh
  
  If Not pfMerge Then
    Clear
  End If
  
  Get #fh, , iCount
  For i = 1 To iCount
    Get #fh, , lColor
    AddColor lColor
  Next i
  
  Close fh
  fIsOpen = False
  LoadFromFile = True
  If Not pfMerge Then mfDirty = False
  
LoadFromFile_Exit:
  If fIsOpen Then
    Close fh
  End If
  Exit Function

LoadFromFile_Err:
  SetErr "LoadFromFile", Err.Number, Err.Description
  Resume LoadFromFile_Exit
End Function

Public Function SaveToFile(ByVal psFilename As String) As Boolean
  On Error GoTo SaveToFile_Err
  ClearErr
  
  Dim fh        As Integer
  Dim fIsOpen   As Boolean
  Dim i         As Long
  
  fh = FreeFile
  Open psFilename For Binary Access Write Lock Read Write As #fh
  fIsOpen = True
  
  Put #fh, , mlCount
  For i = 1 To mlCount
    Put #fh, , malColor(i)
  Next i
  
  Close fh
  fIsOpen = False
  SaveToFile = True
  mfDirty = False
  
SaveToFile_Exit:
  If fIsOpen Then
    Close fh
  End If
  Exit Function

SaveToFile_Err:
  SetErr "SaveToFile", Err.Number, Err.Description
  Resume SaveToFile_Exit
End Function

Public Function MergePalette(poPalette As CColorPalette) As Boolean
  Dim i       As Long
  
  For i = 1 To poPalette.Count
    Call AddColor(poPalette.Color(i))
  Next i
  
  MergePalette = True
  mfDirty = True
End Function

'V02.00.00 This doesn't change the dirty status
Public Sub LoadQBColors()
  'Load QB colors into pristine palette
  Dim i   As Integer
  Dim fSaveDirty As Boolean
  
  fSaveDirty = mfDirty
  For i = 0 To 15
    Me.AddColor QBColor(i)
  Next i
  mfDirty = fSaveDirty
End Sub

Public Sub SortPalette()
  On Error Resume Next
  If mlCount > 1 Then
    QuickSort malColor(), 1, mlCount
  End If
End Sub

Private Sub QuickSort(ByRef palData() As Long, ByVal iLBound As Long, ByVal iUBound As Long)
  Dim lPivot      As Long
  Dim lTemp       As Long
  Dim iLBoundTemp As Long
  Dim iUBoundTemp As Long
  iLBoundTemp = iLBound
  iUBoundTemp = iUBound
  lPivot = palData((iLBound + iUBound) \ 2)
  While (iLBoundTemp <= iUBoundTemp)
    While (palData(iLBoundTemp) < lPivot) And (iLBoundTemp < iUBound)
      iLBoundTemp = iLBoundTemp + 1
    Wend
    While (lPivot < palData(iUBoundTemp)) And (iUBoundTemp > iLBound)
      iUBoundTemp = iUBoundTemp - 1
    Wend
    If iLBoundTemp < iUBoundTemp Then
      lTemp = palData(iLBoundTemp)
      palData(iLBoundTemp) = palData(iUBoundTemp)
      palData(iUBoundTemp) = lTemp
    End If
    If iLBoundTemp <= iUBoundTemp Then
      iLBoundTemp = iLBoundTemp + 1&
      iUBoundTemp = iUBoundTemp - 1&
    End If
  Wend
  'the function calls itself until everything is in good order
  If (iLBound < iUBoundTemp) Then QuickSort palData, iLBound, iUBoundTemp
  If (iLBoundTemp < iUBound) Then QuickSort palData, iLBoundTemp, iUBound
End Sub

