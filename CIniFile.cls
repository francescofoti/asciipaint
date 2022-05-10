VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private msFilename      As String
Private msSection       As String
Private msAsciiZero     As String

Private Const klMaxLongLength       As Long = 11&
Private Const ksOptionsSectionName  As String = "OPTIONS"

#If Win64 Then
Private Declare PtrSafe Function apiGetPrivateProfileSections Lib "kernel32" Alias "GetPrivateProfileStringA" _
                         (ByVal lpApplicationName As LongPtr, ByVal lpKeyName As Any, _
                          ByVal lpDefault As String, ByVal lpReturnedString As String, _
                          ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare PtrSafe Function apiGetPrivateProfileEntries Lib "kernel32" Alias "GetPrivateProfileStringA" _
                         (ByVal lpApplicationName As String, ByVal lpKeyName As LongPtr, _
                          ByVal lpDefault As String, ByVal lpReturnedString As String, _
                          ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare PtrSafe Function apiGetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
                         (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
                          ByVal lpDefault As String, ByVal lpReturnedString As String, _
                          ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare PtrSafe Function apiWritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
                         (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
                          ByVal lpString As Any, ByVal lpFileName As String) As Long
#Else
Private Declare Function apiGetPrivateProfileSections Lib "kernel32" Alias "GetPrivateProfileStringA" _
                         (ByVal lpApplicationName As Long, ByVal lpKeyName As String, _
                          ByVal lpDefault As String, ByVal lpReturnedString As String, _
                          ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function apiGetPrivateProfileEntries Lib "kernel32" Alias "GetPrivateProfileStringA" _
                         (ByVal lpApplicationName As String, ByVal lpKeyName As Long, _
                          ByVal lpDefault As String, ByVal lpReturnedString As String, _
                          ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function apiGetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
                         (ByVal lpApplicationName As String, ByVal lpKeyName As String, _
                          ByVal lpDefault As String, ByVal lpReturnedString As String, _
                          ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function apiWritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
                         (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
                          ByVal lpString As Any, ByVal lplFileName As String) As Long
#End If

Private Sub Class_Initialize()
  msAsciiZero = Chr$(0)
End Sub

Public Property Get Filename() As String
  Filename = msFilename
End Property

Public Property Let Filename(ByVal psFilename As String)
  msFilename = psFilename
End Property

Public Property Get Section() As String
  Section = msSection
End Property

Public Property Let Section(ByVal psSection As String)
  msSection = psSection
End Property

Public Property Get OptionsSectionName() As String
  OptionsSectionName = ksOptionsSectionName
End Property

'
' Public methods
'

Public Function GetString(ByVal psKeyName As String, Optional ByVal plMaxLength As Long = 1024&) As String
  GetString = GetPrivateProfileString(msFilename, msSection, psKeyName, plMaxLength)
End Function

Public Sub SetString(ByVal psKeyName As String, ByVal psValue As String)
  Dim szApplicationName     As String
  Dim szString              As String
  Dim szFileName            As String
  Dim lAPIRet               As Long
    
  szApplicationName = msSection & msAsciiZero
  szFileName = msFilename & msAsciiZero
  psKeyName = psKeyName & msAsciiZero
  psValue = psValue & msAsciiZero
  'Yes, it can fail, but we'll ignore it; you can check by calling GetString() if it matters
  lAPIRet = apiWritePrivateProfileString(szApplicationName, psKeyName, psValue, szFileName)
End Sub

Public Function GetFlag(ByVal psKeyName As String, Optional ByVal pfDefault As Boolean = False) As Boolean
  Dim sIniValue   As String
  
  sIniValue = GetPrivateProfileString(msFilename, msSection, psKeyName, klMaxLongLength)
  If Len(sIniValue) Then
    GetFlag = IsTrueString(sIniValue)
  Else
    GetFlag = pfDefault
  End If
End Function

Public Function GetLong(ByVal psKeyName As String, Optional ByVal plDefault As Long = 0&) As Long
  Dim sIniValue   As String
  
  sIniValue = GetPrivateProfileString(msFilename, msSection, psKeyName, klMaxLongLength)
  If Len(sIniValue) Then
    If CheckLong(sIniValue) = 0& Then
      GetLong = CLng(Val(sIniValue))
    Else
      GetLong = plDefault
    End If
  Else
    GetLong = plDefault
  End If
End Function

Public Function GetInt(ByVal psKeyName As String, Optional ByVal piDefault As Integer = 0&) As Integer
  Dim lValue    As Long
  lValue = GetLong(psKeyName, piDefault)
  If (lValue >= -32768) And (lValue <= 32767) Then
    GetInt = CInt(lValue)
  Else
    GetInt = piDefault
  End If
End Function

Public Function GetDouble(ByVal psKeyName As String, Optional ByVal pdblDefault As Double = 0#) As Double
  Dim sIniValue   As String
  
  sIniValue = GetPrivateProfileString(msFilename, msSection, psKeyName)
  If Len(sIniValue) Then
    GetDouble = CDbl(Val(sIniValue))
  Else
    GetDouble = pdblDefault
  End If
End Function

Public Sub GetOption(ByVal psKeyName As String, ByRef pvRetValue As Variant, Optional ByVal pvDefault As Variant)
  Dim sSaveSection  As String
  Dim sIniValue     As String
  Dim sHexPrefix    As String
  
  sSaveSection = msSection
  msSection = ksOptionsSectionName
  sIniValue = GetString(psKeyName)
  On Error Resume Next
  Select Case VarType(pvRetValue)
  Case vbInteger
    If IsMissing(pvDefault) Then pvDefault = CInt(0)
    pvRetValue = GetInt(psKeyName, pvDefault)
  Case vbLong
    If IsMissing(pvDefault) Then pvDefault = 0&
    pvRetValue = GetLong(psKeyName, pvDefault)
  Case vbByte
    If Len(sIniValue) Then
      pvRetValue = Val(sIniValue)
    Else
      If Not IsMissing(pvDefault) Then
        pvRetValue = pvDefault
      Else
        pvRetValue = False
      End If
    End If
  Case vbBoolean
    If Len(sIniValue) Then
      If IsTrueString(sIniValue) Then
        pvRetValue = True
      Else
        ' parse numeric value to boolean
        pvRetValue = CBool(Val(sIniValue))
      End If
    Else
      If Not IsMissing(pvDefault) Then
        pvRetValue = CBool(pvDefault)
      Else
        pvRetValue = False
      End If
    End If
  Case vbSingle
    If Len(sIniValue) Then
      pvRetValue = Val(sIniValue)
    Else
      If Not IsMissing(pvDefault) Then
        pvRetValue = CSng(pvDefault)
      Else
        pvRetValue = CSng(0)
      End If
    End If
  Case vbDouble
    If IsMissing(pvDefault) Then pvDefault = 0#
    pvRetValue = GetLong(psKeyName, pvDefault)
  Case Else
    pvRetValue = sIniValue
    If Not IsMissing(pvDefault) Then
      If Len(sIniValue) = 0 Then
        pvRetValue = pvDefault
      End If
    End If
  End Select
  msSection = sSaveSection
End Sub

Public Sub SetOption(ByVal sKeyword As String, ByVal vSetValue As Variant)
  Dim sSaveSection  As String
  Dim sSetValue     As String
  
  sSaveSection = msSection
  msSection = ksOptionsSectionName
  On Error Resume Next
  Select Case VarType(vSetValue)
  Case vbByte, vbInteger, vbLong
    sSetValue = CStr(vSetValue)
  Case vbBoolean
    sSetValue = CStr(CInt(vSetValue))
  Case vbString
    sSetValue = vSetValue
  Case Else
    sSetValue = vSetValue & ""
  End Select
  SetString sKeyword, sSetValue
  msSection = sSaveSection
End Sub

Public Function IsOptionSet(ByVal psKeyName As String) As Boolean
  Dim sSaveSection As String
  sSaveSection = msSection
  msSection = ksOptionsSectionName
  IsOptionSet = CBool(Len(GetString(psKeyName)))
  msSection = sSaveSection
End Function

'Returns number of section names, asRetNames Dims are 1 based
Public Function GetSectionNames(ByRef asRetNames() As String, Optional ByVal plMaxBufLen As Long = 4192&) As Long
  Dim szReturnedString      As String
  Dim lAPIRet               As Long
  
  On Error Resume Next
  Erase asRetNames
  szReturnedString = String$(plMaxBufLen, 0)
  lAPIRet = apiGetPrivateProfileSections(0&, "", "", szReturnedString, plMaxBufLen, msFilename)
  If lAPIRet > 0& Then
    GetSectionNames = MakeArray(szReturnedString, asRetNames())
  End If
End Function

'Returns number of section names, asRetNames Dims are 1 based
Public Function GetSectionEntries(ByRef asRetNames() As String, Optional ByVal plMaxBufLen As Long = 4192&) As Long
  Dim szReturnedString      As String
  Dim lAPIRet               As Long
  
  On Error Resume Next
  Erase asRetNames
  szReturnedString = String$(plMaxBufLen, 0)
  lAPIRet = apiGetPrivateProfileEntries(msSection, 0&, "", szReturnedString, plMaxBufLen, msFilename)
  If lAPIRet > 0& Then
    GetSectionEntries = MakeArray(szReturnedString, asRetNames())
  End If
End Function

Private Function MakeArray(ByRef pszString As String, ByRef asRetNames() As String) As Long
  Dim lPos        As Long
  Dim lCount      As Long
  Dim sValue      As String
  
  lPos = InStr(1, pszString, msAsciiZero)
  If lPos Then
    While lPos
      sValue = left$(pszString, lPos - 1)
      pszString = Right$(pszString, Len(pszString) - lPos)
      If Len(sValue) Then
        lCount = lCount + 1&
        ReDim Preserve asRetNames(1 To lCount)
        asRetNames(lCount) = sValue
      End If
      lPos = InStr(1, pszString, msAsciiZero)
    Wend
  Else
    lCount = 1&
    ReDim Preserve asRetNames(1& To lCount)
    asRetNames(lCount) = CtoVB(pszString)
  End If
  
  MakeArray = lCount
End Function

'
' Private methods
'

Private Function GetPrivateProfileString(ByRef psFilename As String, ByRef psSection As String, ByRef psEntryName As String, Optional ByVal lMaxLength As Long = 512&) As String
  Dim szReturnedString      As String
  Dim lAPIRet               As Long
    
  szReturnedString = String$(lMaxLength, " ") & msAsciiZero
  lAPIRet = apiGetPrivateProfileString(psSection, psEntryName, "", szReturnedString, lMaxLength, psFilename)
  If lAPIRet > -1 Then GetPrivateProfileString = Trim$(CtoVB(szReturnedString))
End Function

'a True string is either "True", "Yes" or any value <> 0
Private Function IsTrueString(ByVal sTrueString As String) As Boolean
  If StrComp(sTrueString, "true", vbTextCompare) = 0 Then
    IsTrueString = True
  ElseIf StrComp(sTrueString, "yes", vbTextCompare) = 0 Then
    IsTrueString = True
  Else
    IsTrueString = CBool(Val(sTrueString))
  End If
End Function

'Returns:
'0& : if sValue represents a valid long integer
'1& : if sValue is empty
'2& : String is too long to represent a signed long
'3& : bad character in string
'4& : overflow (or underflow)
Private Function CheckLong(ByVal sValue As String) As Long
  Dim iLen      As Integer
  Dim i         As Integer
  Dim sChar     As String
  Dim iAsc      As Integer
  Dim iAscZero  As Integer
  Dim iAscNine  As Integer
  Dim iAscPlus  As Integer
  Dim iAscMinus As Integer
  Dim lDummy    As Long
  
  iAscZero = Asc("0")
  iAscNine = Asc("9")
  iAscPlus = Asc("+")
  iAscMinus = Asc("-")
  
  iLen = Len(sValue)
  If iLen = 0 Then
    CheckLong = 1&  'string is empty
    Exit Function
  End If
  If iLen > 11 Then
    CheckLong = 2&  'string too long
    Exit Function
  End If
  
  For i = 1 To iLen
    sChar = Mid$(sValue, i, 1)
    iAsc = Asc(sChar)
    If i = 1 Then
      '((iAsc<iAscZero) or (iAsc>iAscNine)) : char is not a number
      '(iAsc<>iAscPlus) and (iAsc<>iAscMinus) but can be + or -
      If ((iAsc < iAscZero) Or (iAsc > iAscNine)) And _
         (iAsc <> iAscPlus) And (iAsc <> iAscMinus) Then
        CheckLong = 3&  'bad character
        Exit Function
      End If
    Else
      '((iAsc<iAscZero) or (iAsc>iAscNine)) : char is not a number
      If ((iAsc < iAscZero) Or (iAsc > iAscNine)) Then
        CheckLong = 3&  'bad character
        Exit Function
      End If
    End If
  Next i
  
  'A long can be between – 2,147,483,648 and 2,147,483,647
  'If we assign its val to a long and there is an error, then its an overflow
  On Error Resume Next
  lDummy = CLng(sValue)
  If Err.Number Then CheckLong = 4&
End Function

'Cut string before trailing msAsciiZero
Private Function CtoVB(ByRef pszString As String) As String
  Dim i   As Long
  i = InStr(pszString, msAsciiZero)
  If i Then
    CtoVB = left$(pszString, i - 1&)
  Else
    CtoVB = pszString
  End If
End Function



