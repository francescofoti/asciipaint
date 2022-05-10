Attribute VB_Name = "MForms"
Option Compare Database
Option Explicit

Private Const TAGPARAM_RESTORESTATE   As String = "_restorestate"
Private Const TAGPARAM_NOLOC          As String = "nolock"
Private Const TAGPARAM_FORMEDITSTATE  As String = "_editstate"
  
Public Sub SetFormTagParam(poForm As Form, ByVal psTagParamName As String, ByVal psValue As String)
  poForm.Tag = SetTagParam(poForm.Tag, psTagParamName, psValue)
End Sub

Public Function GetFormTagParam(poForm As Form, ByVal psTagParamName As String) As String
  GetFormTagParam = GetTagParam(poForm.Tag, psTagParamName)
End Function

Public Function FormSetAllowEdits( _
  ByRef pfrmTarget As Form, _
  ByVal pfAllowEdits As Boolean, _
  ByVal psTagParamName As String, _
  ByVal psTagParamValue As String, _
  ByVal pfDepth As Integer) As Boolean
  On Error Resume Next
  
  Dim fAllowEdits As Boolean
  Dim fDoIt       As Boolean
  
  'Debug.Print "FormSetAllowEdits(" & pfrmTarget.Name & "," & pcmdToggle.Name & ",AllowEdits=" & IIf(pfAllowEdits, "Vrai", "Faux") & ",TagName=" & psTagParamName & ",Tag=" & psTagParamValue & ",Depth=" & pfDepth
  
  On Error GoTo FormSetAllowEdits_Err
  
  If pfDepth > 0 Then
    'if we're recursing and land on a subform already in the required state,
    'we ensure the form's tag is set and we exit
    If pfrmTarget.Form.AllowEdits = pfAllowEdits Then
      'Debug.Print "SORTIE: le sous-formulaire est déjà dans l'état demandé (" & IIf(pfAllowEdits, "Déverouillé", "Verouillé") & ")"
      SetFormTagParam pfrmTarget.Form, TAGPARAM_FORMEDITSTATE, CStr(Abs(CInt(pfAllowEdits)))
      FormSetAllowEdits = True
      Exit Function
    End If
    pfrmTarget.Form.AllowEdits = pfAllowEdits
    SetFormTagParam pfrmTarget.Form, TAGPARAM_FORMEDITSTATE, CStr(Abs(CInt(pfAllowEdits)))
  End If
  
  'propagate
  Dim vChild As Variant
  For Each vChild In pfrmTarget.Controls
    'Debug.Print vChild.Name; IIf(Len(vChild.Tag & "") > 0, "(" & vChild.Tag & ")", "") & ": ";
    If TypeOf vChild Is SubForm Then
      fDoIt = CBool( _
                (GetTagParam(vChild.Tag, psTagParamName) = psTagParamValue) And _
                (GetTagParam(vChild.Tag, TAGPARAM_NOLOC) <> TAGPARAM_NOLOC) _
              )
      If fDoIt Then
        On Error Resume Next
        'Debug.Print "Délégation au sous-formulaire formulaire [" & vChild.Form.Name & "]"
        On Error GoTo FormSetAllowEdits_Err
        FormSetAllowEdits vChild.Form, pfAllowEdits, psTagParamName, psTagParamValue, pfDepth + 1
        vChild.Form.AllowEdits = pfAllowEdits
        vChild.Form.AllowAdditions = pfAllowEdits
        vChild.Form.AllowDeletions = pfAllowEdits
      Else
        'Debug.Print "SKIPPED, sous-formulaire [" & vChild.Form.Name & "]"
      End If
      'Debug.Print "Err allow add/deL: "; Err.Description
    Else
      If (TypeOf vChild Is CommandButton) Or (TypeOf vChild Is ToggleButton) Then
        fDoIt = CBool( _
                  (GetTagParam(vChild.Tag, psTagParamName) = psTagParamValue) And _
                  (GetTagParam(vChild.Tag, TAGPARAM_NOLOC) <> TAGPARAM_NOLOC) _
                )
        If fDoIt Then
          If Not pfAllowEdits Then
            'save current state
            vChild.Tag = SetTagParam(vChild.Tag, TAGPARAM_RESTORESTATE, CStr(CInt(vChild.Enabled)))
            vChild.Enabled = False
          Else
            vChild.Enabled = CBool(Val(GetTagParam(vChild.Tag, TAGPARAM_RESTORESTATE)))
          End If
          'Debug.Print "Enabled="; pfAllowEdits
        Else
          'Debug.Print "(untouched)"
        End If
      ElseIf (TypeOf vChild Is Textbox) Or (TypeOf vChild Is ComboBox) Or _
             (TypeOf vChild Is ListBox) Or (TypeOf vChild Is CheckBox) Then
        fDoIt = CBool( _
                  (GetTagParam(vChild.Tag, psTagParamName) = psTagParamValue) And _
                  (GetTagParam(vChild.Tag, TAGPARAM_NOLOC) <> TAGPARAM_NOLOC) _
                )
        If fDoIt Then
          If Not pfAllowEdits Then
            'save current state
            vChild.Tag = SetTagParam(vChild.Tag, TAGPARAM_RESTORESTATE, CStr(CInt(vChild.Locked)))
            vChild.Locked = True
          Else
            vChild.Locked = CBool(Val(GetTagParam(vChild.Tag, TAGPARAM_RESTORESTATE)))
          End If
        End If
        'Debug.Print "Locked="; Not pfAllowEdits
      Else
        'Debug.Print "(" & "control type name"""; TypeName(vChild); """ untouched)"
      End If
    End If
  Next
  If Err.Number Then
    'Debug.Print "Allow edits error: "; Err.Description
  End If
  
  SetFormTagParam pfrmTarget.Form, TAGPARAM_FORMEDITSTATE, CStr(Abs(CInt(pfAllowEdits)))
  FormSetAllowEdits = True

FormSetAllowEdits_Exit:
  Exit Function
FormSetAllowEdits_Err:
  'Stop
  Resume Next
  Resume
End Function

Public Function FormToggleEdits(ByRef pfrmTarget As Form, ByRef pcmdToggle As CommandButton, ByVal psTagParamName As String, ByVal psTagParamValue As String) As Boolean
  On Error Resume Next
  Dim fAllowEdits As Boolean
  
  If pfrmTarget.Dirty Then
    MsgBox "Le formulaire est en cours de modification, ile ne peut pas être vérrouillé", vbCritical
    Exit Function
  End If
  
  If GetTagParam(pcmdToggle.Tag, "locked") <> "1" Then
    fAllowEdits = False
  Else
    fAllowEdits = True
  End If
  FormToggleEdits = FormSetAllowEdits(pfrmTarget, fAllowEdits, psTagParamName, psTagParamValue, 0)
End Function

Public Function GetTagParam(ByVal psTags As String, ByVal psParamName As String) As String
  If Len(psTags) = 0 Then Exit Function
  
  Dim iPartsCt    As Integer
  Dim iPart       As Integer
  Dim sPart       As String
  Dim sParamName  As String
  Dim sParamValue As String
  Dim iEqual      As Integer
  
  iPartsCt = CountStringParts(psTags, ";")
  For iPart = 1 To iPartsCt
    sPart = GetStringPart(iPart, ";", psTags)
    iEqual = InStr(1, sPart, "=")
    If iEqual > 0 Then
      sParamName = left$(sPart, iEqual - 1)
      sParamValue = Right$(sPart, Len(sPart) - iEqual)
    Else
      sParamName = sPart
      sParamValue = sPart
    End If
    If StrComp(sParamName, psParamName, vbTextCompare) = 0 Then
      GetTagParam = sParamValue
      Exit Function
    End If
  Next iPart
End Function

Public Function SetTagParam(ByVal psTags As String, ByVal psParamName As String, ByVal psParamValue As String) As String
  If Len(psTags) = 0 Then
    SetTagParam = psParamName & "=" & psParamValue
    Exit Function
  End If
  
  Dim asPart()    As String
  Dim iPartCt     As Integer
  Dim iPart       As Integer
  Dim sPart       As String
  Dim sParamName  As String
  Dim sParamValue As String
  Dim iEqual      As Integer
  Dim iFound      As Integer
  Dim sRes        As String
  
  iPartCt = SplitString(asPart(), psTags, ";")
  
  For iPart = 1 To iPartCt
    sPart = asPart(iPart)
    iEqual = InStr(1, sPart, "=")
    If iEqual > 0 Then
      sParamName = left$(sPart, iEqual - 1)
      sParamValue = Right$(sPart, Len(sPart) - iEqual)
    Else
      sParamName = sPart
      sParamValue = sPart
    End If
    If sParamName = psParamName Then
      asPart(iPart) = psParamName & "=" & psParamValue
      iFound = iPart
      Exit For
    End If
  Next iPart
  If iFound = 0 Then
    sRes = psTags & ";" & psParamName & "=" & psParamValue
  Else
    For iPart = 1 To iPartCt
      If iPart > 1 Then
        sRes = sRes & ";"
      End If
      sRes = sRes & asPart(iPart)
    Next iPart
  End If
  
  SetTagParam = sRes
End Function

Public Function IsFormOpen(ByVal pvFormName As Variant) As Boolean
  Dim i         As Integer
  Dim k         As Integer
  Dim sFormName As String
  On Error Resume Next
  
  If Not IsArray(pvFormName) Then
    sFormName = pvFormName & ""
    For i = 0 To Application.Forms.Count - 1
      If Application.Forms(i).Name = sFormName Then
        IsFormOpen = True
        Exit Function
      End If
    Next i
  Else
    For k = LBound(pvFormName) To UBound(pvFormName)
      sFormName = pvFormName(k) & ""
      For i = 0 To Application.Forms.Count - 1
        If Application.Forms(i).Name = sFormName Then
          IsFormOpen = True
          Exit Function
        End If
      Next i
    Next k
  End If
End Function

Public Function LockForm(pFrm As Form) As Boolean
  LockForm = MForms.FormSetAllowEdits(pFrm, False, "", "", 0)
End Function

Public Sub UnlockForm(pFrm As Form)
  Call MForms.FormSetAllowEdits(pFrm, True, "", "", 0)
End Sub

Private Function CanLockUI() As Boolean
  On Error Resume Next
  Dim oForm As Form
  Set oForm = Forms(GetCanvasFormName())
  CanLockUI = CBool(Err.Number = 0)
  Set oForm = Nothing
End Function

Public Function IsUILocked() As Boolean
  On Error Resume Next
  IsUILocked = CBool(GetFormTagParam(Forms(GetCanvasFormName()), TAGPARAM_FORMEDITSTATE) = "0")
End Function

Public Function LockUI() As Boolean
  Dim fOK As Boolean
  
  If CanLockUI() Then
    fOK = IsUILocked()
    If Not fOK Then
      Call MessageManager.Broadcast("LockUI()", MSGTOPIC_LOCKUI, Nothing)
      LockUI = IsUILocked()
    End If
  Else
    LockUI = True
  End If
End Function

Public Sub UnlockUI()
  Dim fOK As Boolean
  
  fOK = IsUILocked()
  If fOK And CanLockUI() Then
    Call MessageManager.Broadcast("UnLockUI()", MSGTOPIC_UNLOCKUI, Nothing)
  End If
End Sub

