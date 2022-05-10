Attribute VB_Name = "MStringsLite"
Option Compare Database
Option Explicit

Public Const PATH_SEP      As String = "\"
Public Const PATH_SEP_INV  As String = "/"
Public Const EXT_SEP       As String = "."
Public Const DRIVE_SEP     As String = ":"

Public Function CombinePath(ByVal psPath1 As String, ByVal psFilename As String) As String
  If left$(psFilename, 1) <> PATH_SEP Then
    CombinePath = NormalizePath(psPath1) & psFilename
  Else
    CombinePath = DenormalizePath(psPath1) & psFilename
  End If
End Function

' Make sure path ends in a backslash.
Public Function NormalizePath(ByVal spath As String) As String
  If Right$(spath, 1) <> PATH_SEP Then
    NormalizePath = spath & PATH_SEP
  Else
    NormalizePath = spath
  End If
End Function

' Make sure path doesn't end in a backslash
Private Function DenormalizePath(ByVal spath As String) As String
  If Right$(spath, 1) = PATH_SEP Then
    spath = left$(spath, Len(spath) - 1)
  End If
  DenormalizePath = spath
End Function

Public Function GetFileExt(ByRef psFilename As String) As String
  Dim lLen    As Long
  Dim i       As Long
  Dim sChar   As String
 
  'Going backwards to find the first EXT_SEP char (or any other path separator)
  lLen = Len(psFilename): i = lLen
  If i Then
    sChar = Mid$(psFilename, i, 1&)
    Do While (i > 0&) And (sChar <> PATH_SEP) And (sChar <> EXT_SEP) And (sChar <> PATH_SEP_INV)
      i = i - 1&: If i = 0& Then Exit Do
      sChar = Mid$(psFilename, i, 1&)
    Loop
    If (i > 0&) And (sChar = EXT_SEP) Then
      GetFileExt = Right$(psFilename, lLen - i)
    End If
  End If
End Function
 
'StripFileExt() returns the left part of a filename (and path), without the
' file extension part.
'ie:
'  StripFileExt("test.txt") gives "test".
'  StripFileExt("C:\mypath\test.txt") gives "C:\mypath\test".
Public Function StripFileExt(ByRef psFilename As String) As String
  Dim lLen    As Long
  Dim i       As Long
  Dim sChar   As String
  
  lLen = Len(psFilename): i = lLen
  If i Then
    sChar = Mid$(psFilename, i, 1&)
    Do While (i > 0&) And (sChar <> PATH_SEP) And (sChar <> EXT_SEP) And (sChar <> PATH_SEP_INV)
      i = i - 1&: If i = 0& Then Exit Do
      sChar = Mid$(psFilename, i, 1&)
    Loop
    If (i& > 0) And (sChar = EXT_SEP) Then
      StripFileExt = left$(psFilename, i - 1&)
    Else
      StripFileExt = psFilename
    End If
  End If
End Function

'StripFilePath() returns only the filename of a full or partial filename and path.
'ie:
'  StripFilePath("C:\mypath\test.txt") gives "test.txt"
Function StripFilePath(ByVal psFilename As String) As String
  Dim i           As Long
  Dim sChar       As String
  
  i = Len(psFilename)
  If i Then
    sChar = Mid$(psFilename, i, 1)
    While (sChar <> DRIVE_SEP) And (sChar <> PATH_SEP) And (i > 0)
      i = i - 1&
      If i Then
        sChar = Mid$(psFilename, i, 1)
      Else
        sChar = PATH_SEP
      End If
    Wend
    If i Then
      StripFilePath = Right$(psFilename, Len(psFilename) - i)
    Else
      StripFilePath = psFilename
    End If
  End If
End Function

'StripFileName() returns only the path of a full or partial filename and path.
'  StripFileName("C:\mypath\test.txt") gives "C:\mypath"
'but, be advised that:
'  StripFileName("\a.txt") gives "" (root directory will be returned as empty string)
Function StripFileName(ByVal psFilename As String) As String
  Dim i           As Long
  Dim fLoop     As Long
  Dim sChar       As String * 1
  
  i = Len(psFilename)
  If i Then fLoop = True
  While fLoop
    If i > 0 Then
      sChar = Mid$(psFilename, i, 1)
      If (sChar = PATH_SEP) Or (sChar = DRIVE_SEP) Then fLoop = False
    End If
    If i > 1& Then
      i = i - 1&
    Else
      i = 0&
      fLoop = False
    End If
  Wend
  If i& Then
    StripFileName = left$(psFilename, i)
  Else
    StripFileName = ""
  End If
End Function

'ChangeExt() returns the given filename (and/or path) with the next
'given file extension.
'ie:
' ChangeExt("c:\temp\test.pdf", "txt") gives c:\temp\test.txt
Function ChangeExt(ByVal sFile As String, ByVal sNewExt As String) As String
  Dim iLen        As Integer
  Dim i           As Integer
  Dim sChar       As String
  
  If left$(sNewExt, 1) = EXT_SEP Then sNewExt = Right$(sNewExt, Len(sNewExt) - 1) 'be forgiving
  iLen = Len(sFile)
  i = iLen
  If i Then
    sChar = Mid$(sFile, i, 1)
    Do While (i > 0) And (sChar <> PATH_SEP) And (sChar <> ".") And (sChar <> PATH_SEP_INV)
      i = i - 1
      If i > 0 Then
        sChar = Mid$(sFile, i, 1)
      End If
    Loop
    If (i > 0) And (sChar = EXT_SEP) Then
      sFile = left$(sFile, i - 1)
    End If
  End If
  ChangeExt = sFile & EXT_SEP & sNewExt
End Function

Public Function IsWhite(ByVal ch$) As Boolean
  If (ch$ > "") Then
    IsWhite = (ch$ = " ") Or (ch$ = vbTab) Or (ch$ = Chr$(13)) Or (ch$ = Chr$(10)) Or (Asc(ch$) = 160)
  End If
End Function

Public Function IsSpace(ByVal ch$) As Boolean
  If (ch$ > "") Then
    IsSpace = (ch$ = " ") Or (ch$ = vbTab) Or (Asc(ch$) = 160)
  End If
End Function

Public Function IsOneOf(ByRef pavValues As Variant, ByVal psPretending As String, Optional ByVal peCompareMethod As VbCompareMethod = VbCompareMethod.vbBinaryCompare) As Boolean
  If IsArray(pavValues) Then
    Dim vItem   As Variant
    For Each vItem In pavValues
      If StrComp(vItem, psPretending, peCompareMethod) = 0 Then
        IsOneOf = True
        Exit Function
      End If
    Next
  Else
    IsOneOf = CBool(StrComp(pavValues & "", psPretending, peCompareMethod) = 0)
  End If
End Function

Function LoByte(ByVal w As Integer) As Byte
    LoByte = w And &HFF
End Function

Function HiByte(ByVal w As Integer) As Byte
    HiByte = (w And &HFF00&) \ 256
End Function

Public Function StringToHexBytes(ByVal psString As String) As String
  Dim i     As Long
  Dim iAscW As Integer
  Dim sRet  As String
  Dim sByte As String
  
  For i = 1 To Len(psString)
    iAscW = AscW(Mid$(psString, i, 1))
    sByte = Hex$(HiByte(iAscW))
    If Len(sByte) < 2 Then sByte = "0" & sByte
    sRet = sRet & sByte
    sByte = Hex$(LoByte(iAscW))
    If Len(sByte) < 2 Then sByte = "0" & sByte
    sRet = sRet & sByte
  Next i
  
  StringToHexBytes = sRet
End Function

Public Function ContainsOneOf(ByRef psString As String, ByVal psChars As String, Optional ByVal peCompareMethod As VbCompareMethod = VbCompareMethod.vbBinaryCompare) As Boolean
  Dim i   As Long
  Dim k   As Long
  
  If Len(psChars) = 0 Then Exit Function
  For i = 1& To Len(psString)
    For k = 1& To Len(psChars)
      If StrComp(Mid$(psString, i, 1), Mid$(psChars, k, 1), peCompareMethod) = 0 Then
        ContainsOneOf = i '20171510 - FFO - Compatible but more useful
        Exit Function
      End If
    Next k
  Next i
End Function

Public Function IsValidFileName(ByVal psFilename As String) As Boolean
  Dim avForbidden   As Variant
  avForbidden = Array("CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9")
  If ContainsOneOf(psFilename, "<>:/\|?*" & Chr$(34)) Or IsOneOf(avForbidden, psFilename, vbTextCompare) Then
    Exit Function
  End If
  IsValidFileName = True
End Function

