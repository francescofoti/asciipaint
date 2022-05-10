Attribute VB_Name = "MDev"
Option Compare Database
Option Explicit

#If ALLOW_CODE_EXPORT Then

Private Sub LocalShowError(ByVal psErrCtx As String, ByVal plErrNo As Long, ByVal psErrText As String)
  Debug.Print "[ERROR] [" & psErrCtx & "] #" & plErrNo & ": " & psErrText
End Sub

Private Function CreatePath(ByVal psPathToMake As String) As Boolean
  Dim sCurPathSegment As String
  Dim iOffset         As Integer
  Dim iAnchor         As Integer
  Dim sOldPath        As String

  On Error Resume Next

  'Add trailing backslash
  If Right$(psPathToMake, 1) <> PATH_SEP Then psPathToMake = psPathToMake & PATH_SEP
  sOldPath = CurDir$
  iAnchor = 0

  'Loop and make each subdir of the path separately.
  iOffset = InStr(iAnchor + 1, psPathToMake, PATH_SEP)
  iAnchor = iOffset 'Start with at least one backslash, i.e. "C:\FirstDir"
  Do
    iOffset = InStr(iAnchor + 1, psPathToMake, PATH_SEP)
    iAnchor = iOffset
    If iAnchor > 0 Then
      sCurPathSegment = left$(psPathToMake, iOffset - 1)
      ' Determine if this directory already exists
      On Error Resume Next
      ChDir sCurPathSegment
      If Err.Number <> 0 Then
        ' We must create this directory
        On Error GoTo CreatePath_Err
        MkDir sCurPathSegment
      End If
    End If
  Loop Until iAnchor = 0

  CreatePath = True
CreatePath_Exit:
  ChDir sOldPath
  Exit Function

CreatePath_Err:
  LocalShowError "CreatePath", Err.Number, Err.Description
  Resume CreatePath_Exit
End Function

'max 9 milions (9 999 999)
Private Function LocalLongChooseBox(ByVal sText As String, ByVal sTitle As String, ByVal sDefault As String, ByVal lMax As Long, Optional ByVal lMin As Long = 1&) As Long
  Dim sInput    As String
  Dim fValid    As Boolean
  Dim lRet      As Long
  
  If lMin < 1& Then lMin = 1&
  If lMax < lMin Then
    lMax = lMin 'sounds dummy...
  End If
  
  Do
    sInput = InputBox$(sText, sTitle, sDefault)
    If Len(sInput) Then
      If IsNumeric(sInput) Then
        If Len(sInput) <= 7 Then
          lRet = CLng(Val(sInput))
          If (lRet >= lMin) And (lRet <= lMax) Then
            fValid = True
          Else
            MsgBox "Please enter a number between " & lMin & " and " & lMax, vbCritical
          End If
        Else
          MsgBox "The text you typed is too long", vbCritical
        End If
      Else
        MsgBox "The text you typed is not a valid number", vbCritical
      End If
    End If
  Loop Until (sInput = "") Or fValid
  
  If fValid Then
    LocalLongChooseBox = lRet
  End If
End Function

Private Function ExportAllModulesAsText(poProj As VBProject, ByVal piProject As Integer, ByVal psExportPath As String, ByVal pfBasAndClsOnly As Boolean, ByVal pfOnlyListFiles As Boolean) As Boolean
  Dim i           As Long
  Dim oComp       As VBComponent
  Dim sFilename   As String
  Dim sFileExt    As String
  
  On Error GoTo EAMAT_Err
  
  i = 1
  For Each oComp In poProj.VBComponents
    If (oComp.Type = vbext_ct_ClassModule) _
       Or ((oComp.Type = vbext_ct_MSForm) And Not pfBasAndClsOnly) _
       Or (oComp.Type = vbext_ct_StdModule) _
       Or ((oComp.Type = vbext_ct_Document) And Not pfBasAndClsOnly) Then
      sFilename = oComp.Name
      
      Select Case oComp.Type
      Case vbext_ct_ClassModule
        sFileExt = ".cls"
      Case vbext_ct_MSForm
        sFileExt = ".frm"
      Case vbext_ct_StdModule
        sFileExt = ".bas"
      Case vbext_ct_Document
        If StrComp(left$(oComp.Name, 5), "Form_", vbTextCompare) = 0 Then
          sFileExt = ".frm"
          sFilename = Right$(sFilename, Len(sFilename) - 5)
        Else
          sFileExt = ".rpt"
          sFilename = Right$(sFilename, Len(sFilename) - 7)
        End If
      End Select
      
      sFilename = CombinePath(psExportPath, sFilename & sFileExt)
      
      Debug.Print i; ") ";
      If Not pfOnlyListFiles Then
        Debug.Print "Exporting [" & oComp.Name & "] to [" & sFilename & "]"
        oComp.Export sFilename
      Else
        Debug.Print "Listing [" & oComp.Name & "] to [" & sFilename & "]"
      End If
    Else
      Debug.Print oComp.Name & " : not exported"
    End If
    i = i + 1
  Next

  ExportAllModulesAsText = True
  
EAMAT_Exit:
  Debug.Print "ExportAllModulesAsText done."
  Exit Function
EAMAT_Err:
  LocalShowError "ExportAllModulesAsText", Err.Number, "[" & sFilename & "] " & Err.Description
  Resume EAMAT_Exit
End Function

Private Function ExportAllObjects(poProj As VBProject, ByVal piProject As Integer, ByVal psExportPath As String, ByVal pfOnlyListFiles As Boolean) As Boolean
  Dim sFilename     As String
  Dim i             As Integer
  Dim sObjectName   As String
  Dim sExportDbName As String
  
  On Error GoTo ExportAllObjects_Err
  
  Debug.Print "Exporting forms and reports objects"
  
  sExportDbName = CombinePath(psExportPath, poProj.Name & "_exported_objects.accdb")
  Debug.Print "Export database: " & sExportDbName
  
  If Not pfOnlyListFiles Then
    If Not ExistFile(sExportDbName) Then
      Access.DBEngine.CreateDatabase sExportDbName, DB_LANG_GENERAL
    End If
  End If
  
  'forms
  Debug.Print CurrentProject.AllForms.Count & " Form(s):"
  For i = 0 To CurrentProject.AllForms.Count - 1
    sObjectName = CurrentProject.AllForms(i).FullName
    If Not pfOnlyListFiles Then
      Debug.Print (i + 1) & ") " & sObjectName & "... ";
      DoCmd.TransferDatabase acExport, "Microsoft Access", sExportDbName, acForm, sObjectName, sObjectName
      Debug.Print "done."
    Else
      Debug.Print (i + 1) & ") " & sObjectName
    End If
  Next i
  
  'Reports
  Debug.Print CurrentProject.AllReports.Count & " Report(s):"
  For i = 0 To CurrentProject.AllReports.Count - 1
    sObjectName = CurrentProject.AllReports(i).FullName
    If Not pfOnlyListFiles Then
      Debug.Print (i + 1) & ") " & sObjectName & "... ";
      DoCmd.TransferDatabase acExport, "Microsoft Access", sExportDbName, acReport, sObjectName, sObjectName
      Debug.Print "done."
    Else
      Debug.Print (i + 1) & ") " & sObjectName
    End If
  Next i
  
  'Macros
  Debug.Print CurrentProject.AllMacros.Count & " Macro(s):"
  For i = 0 To CurrentProject.AllMacros.Count - 1
    sObjectName = CurrentProject.AllMacros(i).FullName
    If Not pfOnlyListFiles Then
      Debug.Print (i + 1) & ") " & sObjectName & "... ";
      DoCmd.TransferDatabase acExport, "Microsoft Access", sExportDbName, acMacro, sObjectName, sObjectName
      Debug.Print "done."
    Else
      Debug.Print (i + 1) & ") " & sObjectName
    End If
  Next i
  
  ExportAllObjects = True
  
ExportAllObjects_Exit:
  Exit Function
ExportAllObjects_Err:
  LocalShowError "ExportAllObjects", Err.Number, "[" & sFilename & "] " & Err.Description
  Resume ExportAllObjects_Exit
End Function

Private Function ExportProjectSettings(poProj As VBProject, ByVal piProject As Integer, ByVal psExportPath As String, ByVal pfOnlyListFiles As Boolean) As Boolean
  Dim sFilename     As String
  Dim i             As Integer
  Dim sTextFilename As String
  Dim hFile         As Integer
  Dim fIsOpen       As Boolean
  
  On Error GoTo ExportProjectSettings_Err
  
  Debug.Print "Exporting project settings"
  
  sTextFilename = CombinePath(psExportPath, poProj.Name & "_project_settings.txt")
  Debug.Print "Export settings to: " & sTextFilename
  
  If Not pfOnlyListFiles Then
    hFile = FreeFile
    Open sTextFilename For Output Access Write Lock Read Write As #hFile
    fIsOpen = True
  End If
  
  Debug.Print "Exporting general settings..."
  If Not pfOnlyListFiles Then Print #hFile, "[General]"
  If Not pfOnlyListFiles Then Print #hFile, "Name=" & poProj.Name
  If Not pfOnlyListFiles Then Print #hFile, "FileName=" & poProj.Filename
  If Not pfOnlyListFiles Then Print #hFile, "BuildFileName=" & poProj.BuildFileName
  If Not pfOnlyListFiles Then Print #hFile, "HelpFile=" & poProj.HelpFile
  If Not pfOnlyListFiles Then Print #hFile, "HelpContextID=" & poProj.HelpContextId
  If Not pfOnlyListFiles Then Print #hFile, "Description=" & poProj.Description
  If Not pfOnlyListFiles Then Print #hFile, "Mode=" & poProj.Mode
  If Not pfOnlyListFiles Then Print #hFile, "Protection=" & poProj.Protection
  If Not pfOnlyListFiles Then Print #hFile, "Saved=" & poProj.Saved
  If Not pfOnlyListFiles Then Print #hFile, "Type=" & poProj.Type
  If Not pfOnlyListFiles Then Print #hFile, "ConditionalCompilationArgs=" & Application.GetOption("Conditional Compilation Arguments")
  If Not pfOnlyListFiles Then Print #hFile, ""
  If Not pfOnlyListFiles Then Print #hFile, "[References]"
  
  Dim iRefCount   As Integer
  Dim oRef        As Object
  
  iRefCount = poProj.References.Count
  If Not pfOnlyListFiles Then Print #hFile, "Count=" & iRefCount
  For i = 1 To iRefCount
    Set oRef = poProj.References(i)
    If Not pfOnlyListFiles Then Print #hFile, "Reference_" & i & "=" & oRef.Guid
    Set oRef = Nothing
  Next i
  If Not pfOnlyListFiles Then Print #hFile, ""
  
  For i = 1 To iRefCount
    Set oRef = poProj.References(i)
    If Not pfOnlyListFiles Then Print #hFile, ""
    If Not pfOnlyListFiles Then Print #hFile, "[" & oRef.Guid & "]"
    If Not pfOnlyListFiles Then Print #hFile, "BuiltIn=" & oRef.BuiltIn
    If Not pfOnlyListFiles Then Print #hFile, "FullPath=" & oRef.FullPath
    If Not pfOnlyListFiles Then Print #hFile, "Guid=" & oRef.Guid
    If Not pfOnlyListFiles Then Print #hFile, "IsBroken=" & oRef.IsBroken
    If Not pfOnlyListFiles Then Print #hFile, "Major=" & oRef.Major
    If Not pfOnlyListFiles Then Print #hFile, "Minor=" & oRef.Minor
    If Not pfOnlyListFiles Then Print #hFile, "Name=" & oRef.Name
    Set oRef = Nothing
  Next i
  
  ExportProjectSettings = True
  
ExportProjectSettings_Exit:
  If fIsOpen Then
    Close hFile
  End If
  Set oRef = Nothing
  Exit Function
ExportProjectSettings_Err:
  LocalShowError "ExportProjectSettings", Err.Number, "[" & sFilename & "] " & Err.Description
  Resume ExportProjectSettings_Exit
  Resume
End Function

Public Sub DoExportModules()
  Dim sExportDir  As String
  Dim sIniSection As String
  
  'sExportDir = CombinePath(CurrentProject.Path, StripFileExt(StripFilePath(CurrentDb.Name)) & "_sources")
  sExportDir = CombinePath(CurrentProject.Path, "sources" & PATH_SEP & APP_VERSION & PATH_SEP & Year(Now) & Format$(Month(Now), "00") & Format$(Day(Now), "00"))
  
  'Debug.Print "ExportDir="; sExportDir
  If Len(sExportDir) = 0 Then
    MsgBox "Please set parameter in registry / ini: topic=Dev, name=ModExportDir", vbCritical
    Exit Sub
  End If
  
  Dim oProj       As VBProject
  Dim sProjects   As String
  Dim iProject    As Long
  Dim iRet        As VbMsgBoxResult
  Dim i           As Long
  Dim fOK         As Boolean
  Dim fListOnly   As Boolean
  
  For i = 1 To VBE.VBProjects.Count
    If i > 1 Then
      sProjects = sProjects & vbCrLf
    End If
    sProjects = sProjects & i & ": " & VBE.VBProjects(i).Name
  Next i
  sProjects = "Enter VBA project index:" & vbCrLf & vbCrLf & sProjects
  iProject = LocalLongChooseBox(sProjects, "Choose project", "1", VBE.VBProjects.Count, 1)
  If iProject = 0 Then Exit Sub
  Set oProj = VBE.VBProjects(iProject)
  
  iRet = MsgBox("Accept Export Path: " & vbCrLf & vbCrLf & sExportDir, vbQuestion + vbOKCancel + vbDefaultButton2)
  If iRet = vbOK Then
    If Not ExistDir(sExportDir) Then
      If Not CreatePath(sExportDir) Then
        LocalShowError "DoExportModules", -1&, "Can't create path [" & sExportDir & "]"
        GoTo DoExportModules_Exit
      End If
    End If
  Else
    GoTo DoExportModules_Exit
  End If
    
  iRet = MsgBox("Test mode ?" & vbCrLf & vbCrLf & "WARNING: this DOES NOT export anything", vbExclamation + vbYesNoCancel + vbDefaultButton2)
  If iRet = vbCancel Then
    GoTo DoExportModules_Exit
  Else
    If iRet = vbYes Then
      fListOnly = True
    End If
  End If
  
  Debug.Print "[" & Now & "] Exporting..."
  
  fOK = ExportAllModulesAsText(oProj, iProject, sExportDir, False, fListOnly)
  If Not fOK Then GoTo DoExportModules_Exit
    
  fOK = ExportAllObjects(oProj, iProject, sExportDir, fListOnly)
  If Not fOK Then GoTo DoExportModules_Exit
  
  fOK = ExportProjectSettings(oProj, iProject, sExportDir, fListOnly)
  If Not fOK Then GoTo DoExportModules_Exit

DoExportModules_Exit:
  Debug.Print "[" & Now & "] Done."
  Set oProj = Nothing
End Sub

#End If
