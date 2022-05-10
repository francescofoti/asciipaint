VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCharMapDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public FontName As String
Public FontSize As Integer
'Settings
Public ShowNonPropGlyphs  As Boolean
Public ShowAllFonts       As Boolean  'FF_MODERN family otherwise (fixed width chars)
'helpers
Public SelectedChar       As String

'IDialog implmentation
Implements IDialog
Private msDialogID    As String
Private mfCancelled   As Boolean
Private mfModal       As Boolean

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
' IDialog implementation
'

Private Sub Class_Initialize()
  Me.FontName = "Lucida Console"
  Me.FontSize = 16
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
  sFormName = GetCharMapFormName()
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

Public Function DialogForm() As Form
  On Error Resume Next
  Set DialogForm = Forms(GetCharMapFormName())
End Function

'
' To generate events (from the form)
'
Public Sub OnCharacterSelected(ByVal plCharCode As Long)
  On Error Resume Next
  SelectedChar = ChrW$(plCharCode)
  MessageManager.Broadcast "CCharMapDialog", MSGTOPIC_CHARSELECTED, plCharCode, GetCanvasFormName()
End Sub
