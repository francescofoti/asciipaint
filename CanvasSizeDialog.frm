VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_CanvasSizeDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'The dialog class associated with this form
Private moDialog      As CCanvasSizeDialog

Private Sub cmdCancel_Click()
  On Error Resume Next
  moDialog.IIDialog.Cancelled = True
  DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub cmdOK_Click()
  Dim iRows     As Integer
  Dim iCols     As Integer
  
  On Error Resume Next
  iRows = Val(Me.txtRows)
  iCols = Val(Me.txtCols)
  
  If (iRows = 0) Or (iCols = 0) Then
    MsgBox "Rows or columns cannot be zero", vbCritical
    Exit Sub
  End If
  
  moDialog.SaveAsDefaults = Abs(Me.chkSetDefault)
  
  moDialog.Rows = iRows
  moDialog.Cols = iCols
  DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub Form_Open(Cancel As Integer)
  On Error Resume Next
  
  If Len(Me.OpenArgs) > 0 Then
    Set moDialog = GetDialogClass(Me.OpenArgs)
  Else
    Set moDialog = New CCanvasSizeDialog
  End If
  
  Me.txtRows = moDialog.Rows
  Me.txtCols = moDialog.Cols
End Sub
