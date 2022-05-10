Attribute VB_Name = "MDialogs"
Option Compare Database
Option Explicit

'The mcolDialogs collection holds references to dialog wrapper classes
'(usually CDlgxxx) that dialog forms can access.
Private mcolDialogs       As New Collection

'A simple cyclic counter for generating dialog ids
Private mlNextDialogID    As Long

Public Function RegDialogClass(ByRef poClassInst As Object) As String
  Dim sKey    As String
  Const MAX_ID As Long = 65535
  On Error GoTo RegDialogClass_Err
  If mlNextDialogID < MAX_ID Then
    mlNextDialogID = mlNextDialogID + 1&
  Else
    mlNextDialogID = 1
  End If
  sKey = CStr(mlNextDialogID)
  mcolDialogs.Add Item:=poClassInst, Key:=sKey
  RegDialogClass = sKey
RegDialogClass_Err:
End Function

Public Sub UnregDialogClass(ByVal psDialogID As String)
  On Error Resume Next
  mcolDialogs.Remove psDialogID
End Sub

Public Function GetDialogClass(ByVal psDialogID As String) As Object
  On Error Resume Next
  Set GetDialogClass = mcolDialogs(psDialogID)
End Function

