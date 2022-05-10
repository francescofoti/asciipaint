VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_GenerateCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const SUBDIR_CODEGENPROFILES As String = "codegen_profiles"

'The dialog class associated with this form
Private moDialog      As CGenerateCodeDialog

Private Sub cmdClose_Click()
  On Error Resume Next
  moDialog.IIDialog.Cancelled = True
  DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

'From the dialog class to the form
Private Sub DDXFromDialog()
  On Error Resume Next
  With moDialog
    Me.cboLanguage = .TargetLanguage
    Me.txtWrapperName = .WrapperMethodName
    Me.txtInvokeName = .CallMethodName
    Me.txtVariableName = .VariableName
    Me.chkRtrim = Not .PreserveRightSpaces
    Me.fraTarget = .RowGenerationMethod
    Me.chkHexEscape = .HexEscape
    Me.txtHexRangeStart = .HexExclRangeStart
    Me.txtHexRangeEnd = .HexExclRangeEnd
  End With
End Sub

'From the form to the dialog class
Private Sub DDXToDialog()
  On Error Resume Next
  With moDialog
    .TargetLanguage = Me.cboLanguage.Column(0)
    .WrapperMethodName = Replace(Trim$(Me.txtWrapperName), " ", "_")
    .CallMethodName = Replace(Trim$(Me.txtInvokeName), " ", "_")
    .VariableName = Replace(Trim$(Me.txtVariableName), " ", "_")
    .PreserveRightSpaces = Not Me.chkRtrim
    .RowGenerationMethod = Me.fraTarget
    .HexEscape = Me.chkHexEscape
    .HexExclRangeStart = Me.txtHexRangeStart
    .HexExclRangeEnd = Me.txtHexRangeEnd
  End With
End Sub

Private Function ValidateDialog() As Boolean
  'Get control values in dialog class members
  DDXToDialog
  With moDialog
    If Len(Trim$(.WrapperMethodName)) = 0 Then
      MsgBox "Please specifiy a wrapper method name", vbCritical
      Me.txtWrapperName.SetFocus
      Exit Function
    End If
    
    If .RowGenerationMethod = eRowGenerationMethod.eFunctionCall Then
      If .TargetLanguage <> eVisualBasic Then
        If Len(Trim$(.CallMethodName)) = 0 Then
          MsgBox "Please specifiy a name for the row function/method call", vbCritical
          Me.txtInvokeName.SetFocus
          Exit Function
        End If
      End If
    Else
      If Len(Trim$(.VariableName)) = 0 Then
        MsgBox "Please specifiy a name for the row variable", vbCritical
        Me.txtVariableName.SetFocus
        Exit Function
      End If
    End If
    
    If .HexEscape Then
      If (.HexExclRangeStart < 0) Or (.HexExclRangeEnd < 0) Or _
         (.HexExclRangeEnd < .HexExclRangeStart) Then
        MsgBox "Invalid range start/end values", vbCritical
        Me.txtHexRangeStart.SetFocus
        Exit Function
      End If
    End If
  End With
  ValidateDialog = True
End Function

Private Sub cmdGenerate_Click()
  Dim fOK     As Boolean
  Dim sCode   As String
  
  On Error Resume Next
  
  'Validate
  If Not ValidateDialog() Then
    Exit Sub
  End If
  
  DoCmd.Hourglass True
  sCode = moDialog.GenerateCode()
  DoCmd.Hourglass False
  If moDialog.LastErr = 0 Then
    Me.txtCode = sCode
    On Error Resume Next
    Me.txtCode.SelStart = 0
    Me.txtCode.SelLength = Len(Me.txtCode)
    Me.txtCode.SetFocus
  Else
    ShowUFError "Code generation failed.", moDialog.LastErrDesc
  End If
End Sub

Private Function SaveToFile() As Boolean
  Static iDumbCounter As Integer
  
  On Error Resume Next
  If iDumbCounter = 0 Then iDumbCounter = 1
  
  Dim iChoice   As Integer
  Dim fSaved    As Boolean
  Dim sFilename As String
  Dim sInitialDir As String
  
  AppIniFile.GetOption INIOPT_CGLASTSAVEPATH, sInitialDir, ""
  sInitialDir = Trim$(sInitialDir)
  If Len(sInitialDir) = 0 Then
    sInitialDir = CombinePath _
        ( _
          CombinePath _
          ( _
            GetSpecialFolder(Application.hWndAccessApp, CSIDL_PERSONAL), APP_NAME _
          ), _
          SUBDIR_CODEGENPROFILES _
        )
  End If
  If Not ExistDir(sInitialDir) Then
    If Not CreatePath(sInitialDir) Then
      sInitialDir = GetSpecialFolder(Application.hWndAccessApp, CSIDL_PERSONAL)
    End If
  End If

  With Application.FileDialog(msoFileDialogSaveAs)
    .Title = "Save Code Generation Profile"
    .InitialFileName = CombinePath(sInitialDir, "AsciiPaint_CodeGenProfile" & iDumbCounter & ".apcgp")
    .Filters.Clear
    .Filters.Add "Code Generation Profiles", "*.apcgp"
    .FilterIndex = 1
    iChoice = .Show()
    If iChoice <> 0 Then
      sFilename = .SelectedItems(1)
      DoCmd.Hourglass True
      fSaved = moDialog.SaveProfile(sFilename)
      DoCmd.Hourglass False
      If fSaved Then
        AppIniFile.SetOption INIOPT_CGLASTSAVEPATH, (StripFileName(sFilename))
        MsgBox "Code Generation Profile saved in [" & sFilename & "]", vbInformation
      Else
        ShowUFError "Failed to save code generation profile to file [" & sFilename & "]", moDialog.LastErrDesc
      End If
    End If
  End With
  
  If fSaved Then
    iDumbCounter = iDumbCounter + 1
    SaveToFile = True
  End If
End Function

Private Sub cmdLoadProfile_Click()
  Call LoadFromFile
End Sub

Private Sub cmdSaveProfile_Click()
  'save dialog to file
  If ValidateDialog() Then
    Call SaveToFile
  End If
End Sub

Private Sub LoadFromFile()
  Dim iChoice     As Integer
  Dim fLoaded     As Boolean
  Dim sFilename   As String
  Dim sMsg        As String
  Dim sInitialDir As String
  
  On Error Resume Next
  
  AppIniFile.GetOption INIOPT_CGLASTLOADPATH, sInitialDir, ""
  sInitialDir = Trim$(sInitialDir)
  If Len(sInitialDir) = 0 Then
    sInitialDir = CombinePath _
        ( _
          CombinePath _
          ( _
            GetSpecialFolder(Application.hWndAccessApp, CSIDL_PERSONAL), APP_NAME _
          ), _
          SUBDIR_CODEGENPROFILES _
        )
  End If
  If Not ExistDir(sInitialDir) Then
    If Not CreatePath(sInitialDir) Then
      sInitialDir = GetSpecialFolder(Application.hWndAccessApp, CSIDL_PERSONAL)
    End If
  End If

  With Application.FileDialog(msoFileDialogFilePicker)
    .Title = "Load code generation profile from file"
    .InitialFileName = NormalizePath(sInitialDir)
    .Filters.Clear
    .Filters.Add "Code Generation Profiles", "*.apcgp"
    .FilterIndex = 1
    iChoice = .Show()
    If iChoice <> 0 Then
      sFilename = .SelectedItems(1)
      DoCmd.Hourglass True
      fLoaded = moDialog.LoadProfile(sFilename)
      If fLoaded Then
        AppIniFile.SetOption INIOPT_CGLASTLOADPATH, (StripFileName(sFilename))
        DDXFromDialog
        DoCmd.Hourglass False
      Else
        DoCmd.Hourglass False
        ShowUFError "Failed to load code generation profile [" & sFilename & "]", moDialog.LastErrDesc
      End If
    End If
  End With
End Sub

Private Sub Form_Open(Cancel As Integer)
  On Error Resume Next
  
  If Len(Me.OpenArgs) > 0 Then
    Set moDialog = GetDialogClass(Me.OpenArgs)
  Else
    Set moDialog = New CGenerateCodeDialog
  End If
  
  DDXFromDialog
End Sub
