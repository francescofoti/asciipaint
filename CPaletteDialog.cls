VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPaletteDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Palette     As New CColorPalette

'IDialog implmentation
Implements IDialog
Private msDialogID    As String
Private mfCancelled   As Boolean
Private mfModal       As Boolean

Private msFilename    As String

' IClassError implementation
Implements IClassError
Private mlErrNo   As Long
Private msErrCtx  As String
Private msErrDesc As String

Public SelectedColorIndex  As Long '0 = no selection. Dialog user maintains it
Event ColorSelected(ByVal plColor As Long)

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
' IDialog implementation
'

Private Sub Class_Initialize()
  msDialogID = RegDialogClass(Me)
End Sub

Private Sub Class_Terminate()
  UnregDialogClass msDialogID
End Sub

Private Property Let IDialog_Cancelled(ByVal pfCancelled As Boolean)
  mfCancelled = pfCancelled
End Property

Private Property Get IDialog_Cancelled() As Boolean
  IDialog_Cancelled = mfCancelled
End Property

Private Property Get IDialog_DialogID() As String
  IDialog_DialogID = msDialogID
End Property

Private Property Get IDialog_IsModal() As Boolean
  IDialog_IsModal = mfModal
End Property

Private Function IDialog_ShowDialog(ByVal pfShowModal As Boolean) As Boolean
  ClearErr
  On Error GoTo ShowDialog_Err
  
  mfCancelled = False
  
  'Create the dialog
  Dim sFormName As String
  mfModal = pfShowModal
  sFormName = GetPaletteFormName()
  If pfShowModal Then
    DoCmd.OpenForm sFormName, acNormal, WindowMode:=acDialog, OpenArgs:=msDialogID
  Else
    DoCmd.OpenForm sFormName, acNormal, WindowMode:=acWindowNormal, OpenArgs:=msDialogID
  End If
  
  IDialog_ShowDialog = True
  Exit Function
ShowDialog_Err:
  SetErr "ShowDialog", Err.Number, Err.Description
End Function

'
' Public methods
'

Public Property Get IIDialog() As IDialog
  Set IIDialog = Me
End Property

Public Property Get Filename() As String  'R/O
  Filename = msFilename
End Property

Public Function DialogForm() As Form
  On Error Resume Next
  Set DialogForm = Forms(GetPaletteFormName())
End Function

Public Sub Clear()
  Me.Palette.Clear
  msFilename = ""
End Sub

Public Function LoadFromFile(ByVal psFilename As String, ByVal pfMerge As Boolean) As Boolean
  Dim fOK       As Boolean
  fOK = Me.Palette.LoadFromFile(psFilename, pfMerge)
  If fOK Then
    msFilename = psFilename
    LoadFromFile = True
  Else
    SetErr Me.Palette.LastErrDesc, Me.Palette.LastErr, Me.Palette.LastErrDesc
  End If
End Function

Public Function SaveToFile(ByVal psFilename As String) As Boolean
  Dim fOK       As Boolean
  fOK = Me.Palette.SaveToFile(psFilename)
  If fOK Then
    msFilename = psFilename
    SaveToFile = True
  Else
    SetErr Me.Palette.LastErrDesc, Me.Palette.LastErr, Me.Palette.LastErrDesc
  End If
End Function

'
' To generate events (from the form)
'
Public Sub OnColorSelected(ByVal plColor As Long)
  On Error Resume Next
  RaiseEvent ColorSelected(plColor)
End Sub

Public Sub SortPalette()
  Me.Palette.SortPalette
End Sub
