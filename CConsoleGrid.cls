VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CConsoleGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Const NATIVE_FILE_MAGIC_TAG_V1 As Long = 3021968 'written in the native file format
Private Const NATIVE_FILE_MAGIC_TAG_V2 As Long = &H41534350
'V02.00.06 File Format Version 3, keeps V2 file magic tag
'Dropping support for file format less than previous version (< 2)
Private Const CLASS_VERSION As Long = 3&              'written in the native file format
Private Const MIN_SUPPORTED_VERSION As Long = 2&

'Rows and Cols for buffers
Private miRows      As Integer
Private miCols      As Integer

'Buffers
Private mfBuffersCreated  As Boolean
Private mfDirty           As Boolean  'if something changes, exposed via Dirty property
'Characters
Private masChars()    As String
'Two dimesions (rows*cols)
Private malForeCol()  As Long
Private mabForeCol()  As Byte 'indicates if malForeCol item is used
Private malBackCol()  As Long
Private mabBackCol()  As Byte
Private maiAttribs()  As Integer
' 11 bits for "attributes" (repeated declarations from MSupport, but private for easier integration)
Private Const ATTRIB_BOLDON      As Integer = 1
Private Const ATTRIB_ITALICON    As Integer = 2
Private Const ATTRIB_UNDLON      As Integer = 4
Private Const ATTRIB_STRIKEON    As Integer = 8
Private Const ATTRIB_INVERSEON   As Integer = 16
Private Const ATTRIB_BOLDOFF     As Integer = 32
Private Const ATTRIB_ITALICOFF   As Integer = 64
Private Const ATTRIB_UNDLOFF     As Integer = 128
Private Const ATTRIB_STRIKEOFF   As Integer = 256
Private Const ATTRIB_INVERSEOFF  As Integer = 512
Private Const ATTRIB_RESET       As Integer = 1024
Private Const ATTRIB_SPILLON     As Integer = 2048
Private Const ATTRIB_SPILLOFF    As Integer = 4096
Private Const ATTRIB_BITS As Integer = 12 'do not include the reset "attribute", its not a toggle
Private maiAttribMask(0 To ATTRIB_BITS - 1) As Integer 'Will contain powers of 2 to avoid recalc
Private masAttribInject(0 To ATTRIB_BITS - 1) As String 'Initialized with VT100 escape codes per attribmask

' DrawText() Format Flags
Private Const DT_TOP = &H0
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_RIGHT = &H2
Private Const DT_VCENTER = &H4
Private Const DT_BOTTOM = &H8
Private Const DT_WORDBREAK = &H10
Private Const DT_SINGLELINE = &H20
Private Const DT_EXPANDTABS = &H40
Private Const DT_TABSTOP = &H80
Private Const DT_NOCLIP = &H100
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_CALCRECT = &H400
Private Const DT_NOPREFIX = &H800
Private Const DT_INTERNAL = &H1000

Private mfFragmentText As Boolean
Private miFragmentCharWidth As Integer

'We'll cache fully rendered VT100 lines, so we don't have to rebuild them
'on a subsequent call in GetLineVT100() method.
Private moVT100Cache    As New CVT100Cache

'For rendering a line
Private Type TInjection
  ID          As Long 'needed to sort multiple entries at .iPos
  iPos        As Integer
  sInject     As String
End Type
Private matInjection() As TInjection
Private miInjectionCount  As Integer
Private miInjectionSize   As Integer

'Whether we want to show line numbers
Private mfShowBoundaries  As Boolean  'whether we add "|" at end of lines and "---...--+" at end of output

'Selected item
Private miSelCol    As Integer
Private miSelRow    As Integer
Private miSelZone   As Integer

Public Enum eClearDirection
  eToCursorPosition = 0
  eFromCursorPosition
End Enum

'search and replace char
Private miLastFindRow   As Integer
Private miLastFindCol   As Integer
Private msLastFindChar  As String
Public Enum eReplaceCharScope
  eToEndOfLine
  eFullLine
  eToEndOfText
  eFullText
  eSelection
End Enum
Public Enum eReplaceColorScope
  eForeColor
  eBackColor
  eBoth
End Enum

' For VT100 parsing
Private Const LSCMD_NOOP          As Integer = 0
Private Const LSCMD_TEXTOUT       As Integer = 1
Private Const LSCMD_INVERSEON     As Integer = 2
Private Const LSCMD_INVERSEOFF    As Integer = 3
Private Const LSCMD_BOLDON        As Integer = 4
Private Const LSCMD_ITALICON      As Integer = 5
Private Const LSCMD_UNDERLINEON   As Integer = 6
Private Const LSCMD_BOLDOFF       As Integer = 7
Private Const LSCMD_ITALICOFF     As Integer = 8
Private Const LSCMD_UNDERLINEOFF  As Integer = 9
Private Const LSCMD_CROSSEDOUTON  As Integer = 10
Private Const LSCMD_CROSSEDOUTOFF As Integer = 11
Private Const LSCMD_SETFORECOLOR  As Integer = 20
Private Const LSCMD_SETBACKCOLOR  As Integer = 21
Private Const LSCMD_BEGINZONE     As Integer = 90
Private Const LSCMD_ENDZONE       As Integer = 91
Private Const LSCMD_RESET         As Integer = 99
Private Const LSCMD_SETWIDTH      As Integer = 100
Private Const LSCMD_ADVANCEABS    As Integer = 101
Private Const LSCMD_ADVANCEREL    As Integer = 102
Private Const LSCMD_SAVEGPOS      As Integer = 103
Private Const LSCMD_RESTOREGPOS   As Integer = 104
Private Const LSCMD_DTFLAGS       As Integer = 110
Private Const LSCMD_BKCOLSPILL    As Integer = 111 'V01.02.04

Private Const VT100_RESET           As String = "0"
Private Const VT100_INVERSEON       As String = "7"
Private Const VT100_INVERSEOFF      As String = "27"
Private Const VT100_BOLDON          As String = "1"
Private Const VT100_BOLDOFF         As String = "21"
Private Const VT100_ITALICON        As String = "3"
Private Const VT100_ITALICOFF       As String = "23"
Private Const VT100_UNDERLINEON     As String = "4"
Private Const VT100_UNDERLINEOFF    As String = "24"
Private Const VT100_CROSSEDOUT_ON   As String = "9"
Private Const VT100_CROSSEDOUT_OFF  As String = "29"
Private Const VT100_SETFGCOLOR      As String = "38"
Private Const VT100_SETBKCOLOR      As String = "48"
Private Const VT100X_BEGINZONE      As String = "98"
Private Const VT100X_ENDZONE        As String = "99"
Private Const VT100X_SETWIDTH       As String = "100" 'V01.02.00 Adding support for setting width of next TEXT
Private Const VT100X_ADVANCEABS     As String = "101"
Private Const VT100X_ADVANCEREL     As String = "102"
Private Const VT100X_SAVEGPOS       As String = "103"
Private Const VT100X_RESTOREGPOS    As String = "104"
Private Const VT100X_DTFLAGS        As String = "110" 'V01.02.03
Private Const VT100X_BKCOLSPILL     As String = "111" 'V01.02.04

'V02.00.00 easier to raise an event than to chain code for dirty state
Event OnDirtyChange(ByVal pfDirty As Boolean)

Private Const SECTIONID_FORECOLOR_COLOR As Integer = 1
Private Const SECTIONID_FORECOLOR_ONOFF As Integer = 2
Private Const SECTIONID_BACKCOLOR_COLOR As Integer = 3
Private Const SECTIONID_BACKCOLOR_ONOFF As Integer = 4
Private Const SECTIONID_ATTRIBUTES      As Integer = 5
Private Const SECTIONID_CHARWDATA       As Integer = 6

'API
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

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

' Class events

Private Sub Class_Initialize()
  Dim i As Integer
  'Precompute powers of 2
  maiAttribMask(0) = ATTRIB_BOLDON
  maiAttribMask(1) = ATTRIB_ITALICON
  maiAttribMask(2) = ATTRIB_UNDLON
  maiAttribMask(3) = ATTRIB_STRIKEON
  maiAttribMask(4) = ATTRIB_INVERSEON
  maiAttribMask(5) = ATTRIB_BOLDOFF
  maiAttribMask(6) = ATTRIB_ITALICOFF
  maiAttribMask(7) = ATTRIB_UNDLOFF
  maiAttribMask(8) = ATTRIB_STRIKEOFF
  maiAttribMask(9) = ATTRIB_INVERSEOFF
  maiAttribMask(10) = ATTRIB_SPILLON
  maiAttribMask(11) = ATTRIB_SPILLOFF
  
  masAttribInject(0) = VT_BOLD_ON 'ATTRIB_BOLDON
  masAttribInject(1) = VT_ITAL_ON 'ATTRIB_ITALICON
  masAttribInject(2) = VT_UNDL_ON 'ATTRIB_UNDLON
  masAttribInject(3) = VT_STRIKE_ON 'ATTRIB_STRIKEON
  masAttribInject(4) = VT_INV_ON 'ATTRIB_INVERSEON
  masAttribInject(5) = VT_BOLD_OFF 'ATTRIB_BOLDOFF
  masAttribInject(6) = VT_ITAL_OFF 'ATTRIB_ITALICOFF
  masAttribInject(7) = VT_UNDL_OFF 'ATTRIB_UNDLOFF
  masAttribInject(8) = VT_STRIKE_OFF 'ATTRIB_STRIKEOFF
  masAttribInject(9) = VT_INV_OFF 'ATTRIB_INVERSEOFF
  masAttribInject(10) = VTX_SPILL(1) 'ATTRIB_SPILLON
  masAttribInject(11) = VTX_SPILL(0) 'ATTRIB_SPILLOFF
  'default rows/cols
  miRows = 40
  miCols = 80
  mfShowBoundaries = True
End Sub

Private Sub Class_Terminate()
  'nothing to do for now
End Sub

Public Function GetMemoryFootprint() As Double
  Dim dblByteCt     As Double
  
  dblByteCt = dblByteCt + LenB(miRows)
  dblByteCt = dblByteCt + LenB(miCols)
  dblByteCt = dblByteCt + LenB(mfBuffersCreated)
  dblByteCt = dblByteCt + LenB(mfDirty)
  
  On Error Resume Next
  Dim iLB1    As Integer
  Dim iUB1    As Integer
  Dim iLB2    As Integer
  Dim iUB2    As Integer
  Dim i       As Long
  Dim j       As Long
  
  'masChars
  iLB1 = LBound(masChars)
  iUB1 = UBound(masChars)
  For i = iLB1 To iUB1
    dblByteCt = dblByteCt + LenB(masChars(i))
  Next i
  'Debug.Print "After masChars(), bytes=" & dblByteCt
  'Debug.Print "After masChars(), Kbytes=" & Format$(dblByteCt / 1024, "0.####")
  'Debug.Print "After masChars(), Mbytes=" & Format$((dblByteCt / 1024) / 1024, "0.####")
  
  'malForeCol
  iLB1 = LBound(malForeCol, 1)
  iUB1 = UBound(malForeCol, 1)
  iLB2 = LBound(malForeCol, 2)
  iUB2 = UBound(malForeCol, 2)
  For i = iLB1 To iUB1
    For j = iLB2 To iUB2
      dblByteCt = dblByteCt + LenB(malForeCol(i, j))
    Next j
  Next i
  'Debug.Print "After malForeCol(), bytes=" & dblByteCt
  'Debug.Print "After malForeCol(), Kbytes=" & Format$(dblByteCt / 1024, "0.####")
  'Debug.Print "After malForeCol(), Mbytes=" & Format$((dblByteCt / 1024) / 1024, "0.####")
  
  'mabForeCol
  iLB1 = LBound(mabForeCol, 1)
  iUB1 = UBound(mabForeCol, 1)
  iLB2 = LBound(mabForeCol, 2)
  iUB2 = UBound(mabForeCol, 2)
  For i = iLB1 To iUB1
    For j = iLB2 To iUB2
      dblByteCt = dblByteCt + LenB(mabForeCol(i, j))
    Next j
  Next i
  'Debug.Print "After mabForeCol(), bytes=" & dblByteCt
  'Debug.Print "After mabForeCol(), Kbytes=" & Format$(dblByteCt / 1024, "0.####")
  'Debug.Print "After mabForeCol(), Mbytes=" & Format$((dblByteCt / 1024) / 1024, "0.####")
  
  'malBackCol
  iLB1 = LBound(malBackCol, 1)
  iUB1 = UBound(malBackCol, 1)
  iLB2 = LBound(malBackCol, 2)
  iUB2 = UBound(malBackCol, 2)
  For i = iLB1 To iUB1
    For j = iLB2 To iUB2
      dblByteCt = dblByteCt + LenB(malBackCol(i, j))
    Next j
  Next i
  'Debug.Print "After malBackCol(), bytes=" & dblByteCt
  'Debug.Print "After malBackCol(), Kbytes=" & Format$(dblByteCt / 1024, "0.####")
  'Debug.Print "After malBackCol(), Mbytes=" & Format$((dblByteCt / 1024) / 1024, "0.####")
  
  'mabBackCol
  iLB1 = LBound(mabBackCol, 1)
  iUB1 = UBound(mabBackCol, 1)
  iLB2 = LBound(mabBackCol, 2)
  iUB2 = UBound(mabBackCol, 2)
  For i = iLB1 To iUB1
    For j = iLB2 To iUB2
      dblByteCt = dblByteCt + LenB(mabBackCol(i, j))
    Next j
  Next i
  'Debug.Print "After mabBackCol(), bytes=" & dblByteCt
  'Debug.Print "After mabBackCol(), Kbytes=" & Format$(dblByteCt / 1024, "0.####")
  'Debug.Print "After mabBackCol(), Mbytes=" & Format$((dblByteCt / 1024) / 1024, "0.####")
  
  'maiAttribs
  iLB1 = LBound(maiAttribs, 1)
  iUB1 = UBound(maiAttribs, 1)
  iLB2 = LBound(maiAttribs, 2)
  iUB2 = UBound(maiAttribs, 2)
  For i = iLB1 To iUB1
    For j = iLB2 To iUB2
      dblByteCt = dblByteCt + LenB(maiAttribs(i, j))
    Next j
  Next i
  'Debug.Print "After maiAttribs(), bytes=" & dblByteCt
  'Debug.Print "After maiAttribs(), Kbytes=" & Format$(dblByteCt / 1024, "0.####")
  'Debug.Print "After maiAttribs(), Mbytes=" & Format$((dblByteCt / 1024) / 1024, "0.####")
  
  'maiAttribMask
  iLB1 = LBound(maiAttribMask)
  iUB1 = UBound(maiAttribMask)
  For i = iLB1 To iUB1
    dblByteCt = dblByteCt + LenB(maiAttribMask(i))
  Next i
  'Debug.Print "After maiAttribMask(), bytes=" & dblByteCt
  'Debug.Print "After maiAttribMask(), Kbytes=" & Format$(dblByteCt / 1024, "0.####")
  'Debug.Print "After maiAttribMask(), Mbytes=" & Format$((dblByteCt / 1024) / 1024, "0.####")
  
  'masAttribInject
  iLB1 = LBound(masAttribInject)
  iUB1 = UBound(masAttribInject)
  For i = iLB1 To iUB1
    dblByteCt = dblByteCt + LenB(masAttribInject(i))
  Next i
  'Debug.Print "After masAttribInject(), bytes=" & dblByteCt
  'Debug.Print "After masAttribInject(), Kbytes=" & Format$(dblByteCt / 1024, "0.####")
  'Debug.Print "After masAttribInject(), Mbytes=" & Format$((dblByteCt / 1024) / 1024, "0.####")
  
  'matInjection
  iLB1 = LBound(matInjection)
  iUB1 = UBound(matInjection)
  For i = iLB1 To iUB1
    dblByteCt = dblByteCt + LenB(matInjection(i))
  Next i
  'Debug.Print "After matInjection(), bytes=" & dblByteCt
  'Debug.Print "After matInjection(), Kbytes=" & Format$(dblByteCt / 1024, "0.####")
  'Debug.Print "After matInjection(), Mbytes=" & Format$((dblByteCt / 1024) / 1024, "0.####")
  
  dblByteCt = dblByteCt + LenB(miInjectionCount)
  dblByteCt = dblByteCt + LenB(miInjectionSize)
  dblByteCt = dblByteCt + LenB(mfShowBoundaries)
  dblByteCt = dblByteCt + LenB(miSelCol)
  dblByteCt = dblByteCt + LenB(miSelRow)
  dblByteCt = dblByteCt + LenB(miSelZone)
  
  dblByteCt = dblByteCt + LenB(miLastFindRow)
  dblByteCt = dblByteCt + LenB(miLastFindCol)
  dblByteCt = dblByteCt + LenB(msLastFindChar)
  
  dblByteCt = dblByteCt + LenB(mlErrNo)
  dblByteCt = dblByteCt + LenB(msErrCtx)
  dblByteCt = dblByteCt + LenB(msErrDesc)
  
  'Debug.Print "Finally, bytes=" & dblByteCt
  'Debug.Print "Finally, Kbytes=" & Format$(dblByteCt / 1024, "0.####")
  'Debug.Print "Finally, Mbytes=" & Format$((dblByteCt / 1024) / 1024, "0.####")
  
  GetMemoryFootprint = dblByteCt
End Function

'
' Public properties
'

Public Property Get Dirty() As Boolean
  Dirty = mfDirty
End Property

Private Sub SetDirty(ByVal pfDirty As Boolean)
  mfDirty = pfDirty
  On Error Resume Next
  RaiseEvent OnDirtyChange(mfDirty)
End Sub

Public Property Get Rows() As Integer
  Rows = miRows
End Property

Public Property Get Cols() As Integer
  Cols = miCols
End Property

Public Property Get CharAt(ByVal piRow As Integer, ByVal piCol As Integer) As String
  If (piRow > 0) And (piRow <= miRows) And (piCol > 0) And (piCol <= miCols) Then
    CharAt = Mid$(masChars(piRow), piCol, 1)
  End If
End Property

Public Property Let CharAt(ByVal piRow As Integer, ByVal piCol As Integer, ByVal psChar As String)
  If (Len(psChar) = 1) And (piRow > 0) And (piRow <= miRows) And (piCol > 0) And (piCol <= miCols) Then
    Mid$(masChars(piRow), piCol, 1) = left$(psChar, 1)
    SetDirty True
    moVT100Cache.InvalidateLine piRow
  End If
End Property

Public Property Get TextAt(ByVal piRow As Integer, ByVal piCol As Integer) As String
  If (piRow > 0) And (piRow <= miRows) And (piCol > 0) And (piCol <= miCols) Then
    TextAt = Right$(masChars(piRow), miCols - piCol + 1)
  End If
End Property

Public Property Let TextAt(ByVal piRow As Integer, ByVal piCol As Integer, ByVal psText As String)
  If (Len(psText) > 0) And (piRow > 0) And (piRow <= miRows) And (piCol > 0) And (piCol <= miCols) Then
    Mid$(masChars(piRow), piCol, Len(psText)) = psText
    If Len(masChars(piRow)) > miCols Then 'truncate to max number of columns
      masChars(piRow) = left$(masChars(piRow), miCols)
    End If
    SetDirty True
    moVT100Cache.InvalidateLine piRow
  End If
End Property

Public Property Get CharForeCol(ByVal piRow As Integer, ByVal piCol As Integer) As Long
  If (piRow > 0) And (piRow <= miRows) And (piCol > 0) And (piCol <= miCols) Then
    CharForeCol = malForeCol(piRow, piCol)
  End If
End Property

Public Property Let CharForeCol(ByVal piRow As Integer, ByVal piCol As Integer, ByVal plForeCol As Long)
  If (piRow > 0) And (piRow <= miRows) And (piCol > 0) And (piCol <= miCols) Then
    malForeCol(piRow, piCol) = plForeCol
    mabForeCol(piRow, piCol) = 1
    SetDirty True
    moVT100Cache.InvalidateLine piRow
  End If
End Property

Public Property Get HasForeCol(ByVal piRow As Integer, ByVal piCol As Integer) As Byte
  If (piRow > 0) And (piRow <= miRows) And (piCol > 0) And (piCol <= miCols) Then
    HasForeCol = mabForeCol(piRow, piCol)
  End If
End Property

Public Property Let HasForeCol(ByVal piRow As Integer, ByVal piCol As Integer, ByVal pbHasForeCol As Byte)
  If (piRow > 0) And (piRow <= miRows) And (piCol > 0) And (piCol <= miCols) Then
    mabForeCol(piRow, piCol) = pbHasForeCol
    SetDirty True
    moVT100Cache.InvalidateLine piRow
  End If
End Property

Public Property Get CharBackCol(ByVal piRow As Integer, ByVal piCol As Integer) As Long
  If (piRow > 0) And (piRow <= miRows) And (piCol > 0) And (piCol <= miCols) Then
    CharBackCol = malBackCol(piRow, piCol)
  End If
End Property

Public Property Let CharBackCol(ByVal piRow As Integer, ByVal piCol As Integer, ByVal plBackCol As Long)
  If (piRow > 0) And (piRow <= miRows) And (piCol > 0) And (piCol <= miCols) Then
    malBackCol(piRow, piCol) = plBackCol
    mabBackCol(piRow, piCol) = 1
    SetDirty True
    moVT100Cache.InvalidateLine piRow
  End If
End Property

Public Property Get HasBackCol(ByVal piRow As Integer, ByVal piCol As Integer) As Byte
  If (piRow > 0) And (piRow <= miRows) And (piCol > 0) And (piCol <= miCols) Then
    HasBackCol = mabBackCol(piRow, piCol)
  End If
End Property

Public Property Let HasBackCol(ByVal piRow As Integer, ByVal piCol As Integer, ByVal pbHasBackCol As Byte)
  If (piRow > 0) And (piRow <= miRows) And (piCol > 0) And (piCol <= miCols) Then
    mabBackCol(piRow, piCol) = pbHasBackCol
    SetDirty True
    moVT100Cache.InvalidateLine piRow
  End If
End Property

Public Property Get CharAttribs(ByVal piRow As Integer, ByVal piCol As Integer) As Integer
  If (piRow > 0) And (piRow <= miRows) And (piCol > 0) And (piCol <= miCols) Then
    CharAttribs = maiAttribs(piRow, piCol)
  End If
End Property

Public Property Let CharAttribs(ByVal piRow As Integer, ByVal piCol As Integer, ByVal piAttribs As Integer)
  If (piRow > 0) And (piRow <= miRows) And (piCol > 0) And (piCol <= miCols) Then
    maiAttribs(piRow, piCol) = piAttribs
    SetDirty True
    moVT100Cache.InvalidateLine piRow
  End If
End Property

Public Property Get ShowBoundaries() As Boolean
  ShowBoundaries = mfShowBoundaries
End Property

Public Property Let ShowBoundaries(ByVal pfShowBoundaries As Boolean)
  If mfShowBoundaries <> pfShowBoundaries Then
    mfShowBoundaries = pfShowBoundaries
    moVT100Cache.Clear
  End If
End Property

Public Property Get FragmentText() As Boolean
  FragmentText = mfFragmentText
End Property

Public Sub SetFragmentText(ByVal pfAutoFragmentText As Boolean, ByVal piFragmentCharWidth As Integer)
  mfFragmentText = pfAutoFragmentText
  miFragmentCharWidth = piFragmentCharWidth
End Sub

'
' Public methods
'

Public Function Resize(ByVal piRows As Integer, ByVal piCols As Integer) As Boolean
  Const LOCAL_ERR_CTX As String = "Resize"
  On Error GoTo Resize_Err
  Dim grdTemp   As CConsoleGrid
  Dim fOK       As Boolean
  Dim fCopyData As Boolean
  
  ClearErr

  If (piRows <= 0) Or (piCols <= 0) Then
    SetErr "Resize", -1&, "Invalid size"
    Exit Function
  End If
  
  'We have to copy the grid to a temporary new one
  'then paste it back to the new canvas
  
  'create a copy
  fCopyData = mfBuffersCreated
  If fCopyData Then
    Set grdTemp = New CConsoleGrid
    fOK = grdTemp.CreateFrom(Me, 1, miRows, 1, miCols)
    If Not fOK Then
      SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
      GoTo Resize_exit
    End If
  End If
  
  'Create new grid in this instance
  miRows = piRows
  miCols = piCols
  fOK = CreateBuffers()
  If Not fOK Then
    GoTo Resize_exit
  End If
  
  'Paste back
  If fCopyData Then
    fOK = Me.Paste(grdTemp, 1, 1, False)
  End If
  
Resize_exit:
  Resize = fOK
  Set grdTemp = Nothing
  Exit Function

Resize_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  Resume Resize_exit
End Function

Private Function CreateBuffers() As Boolean
  Dim i     As Integer
  
  On Error GoTo CreateBuffers_Err
  
  '(re)size buffers
  ReDim masChars(1 To miRows) As String
  ReDim malForeCol(1 To miRows, 1 To miCols) As Long
  ReDim mabForeCol(1 To miRows, 1 To miCols) As Byte
  ReDim malBackCol(1 To miRows, 1 To miCols) As Long
  ReDim mabBackCol(1 To miRows, 1 To miCols) As Byte
  ReDim maiAttribs(1 To miRows, 1 To miCols) As Integer
  For i = 1 To miRows
    masChars(i) = Space$(miCols)
  Next i
  
  moVT100Cache.Clear
  mfBuffersCreated = True
  CreateBuffers = True
  
CreateBuffers_exit:
  Exit Function

CreateBuffers_Err:
  SetErr "CreateBuffers", Err.Number, Err.Description
  Resume CreateBuffers_exit
End Function

Public Function LoadConsole(poConsole As CConsoul, ByRef iiProgress As IProgressIndicator) As Boolean
  Dim iRows   As Integer
  
  On Error GoTo LoadConsole_Err
  ClearErr

  poConsole.RenderMode = rmByLine
  If Not mfBuffersCreated Then
    If Not CreateBuffers() Then
      Exit Function
    End If
  End If
    
  Dim i           As Integer
  Dim sRender     As String
  Dim iExistingLines  As Integer
  
  If Not mfShowBoundaries Then
    poConsole.Clear
  End If
  
  iRows = miRows
  If poConsole.MaxCapacity < miRows Then
    iRows = poConsole.MaxCapacity - 1
  End If
  iExistingLines = poConsole.LineCount
  
  If Not iiProgress Is Nothing Then iiProgress.SetMax iRows
  
  poConsole.AutoRedraw = False
  For i = 1 To iRows
    sRender = Me.GetLineVT100(i, mfShowBoundaries)
    If i <= iExistingLines Then
      poConsole.SetLine i, sRender
    Else
      poConsole.OutputLn sRender
      iExistingLines = iExistingLines + 1
    End If
    If Not iiProgress Is Nothing Then iiProgress.SetValue i
    'DbgWait 1
  Next i
  
  If Not iiProgress Is Nothing Then iiProgress.SetCaption "Finishing"
  If mfShowBoundaries Then
    sRender = String$(miCols, "-") & "+"
    If iExistingLines < iRows + 1 Then
      poConsole.OutputLn sRender
      poConsole.OutputLn Space$(miCols + 1)
    Else
      poConsole.SetLine iRows + 1, sRender
    End If
  End If
  
  LoadConsole = True

LoadConsole_exit:
  poConsole.AutoRedraw = True
  Exit Function

LoadConsole_Err:
  SetErr "LoadConsole", Err.Number, Err.Description
  Resume LoadConsole_exit
End Function

Private Sub InitInjectionsArray()
  miInjectionSize = miCols * 2
  ReDim matInjection(1 To miInjectionSize) As TInjection
  miInjectionCount = 0
End Sub

Private Sub AddInjection(ByVal piPos As Integer, ByVal psInject As String)
  Static lNextID As Long
  
  lNextID = lNextID + 1
  If lNextID > 2147483647 Then
    lNextID = 1&  '/**/this will break the sort order; no solution at this time.
  End If
  
  If miInjectionCount = miInjectionSize Then
    ReDim Preserve matInjection(1 To miInjectionSize * 2) As TInjection
    miInjectionSize = miInjectionSize * 2
  End If
  miInjectionCount = miInjectionCount + 1
  With matInjection(miInjectionCount)
    .ID = lNextID
    .iPos = piPos
    .sInject = psInject
  End With
End Sub

Private Sub GenForeColInjections(ByVal piLine As Integer)
  Dim i               As Integer
  Dim fColorChanged   As Boolean
  Dim lLastColor      As Long
  Dim iAttribs        As Integer
  Dim fReset          As Boolean
  Dim fFirstTime      As Boolean
  
  i = 1
  lLastColor = malForeCol(piLine, 1)
  fFirstTime = True
  Do
    iAttribs = maiAttribs(piLine, i)
    fReset = CBool(iAttribs And ATTRIB_RESET)
    If fReset Then fFirstTime = True
    
    If mabForeCol(piLine, i) Then
      fColorChanged = fFirstTime Or CBool(malForeCol(piLine, i) <> lLastColor)
      If fColorChanged Then
        AddInjection i, VT_FCOLOR(malForeCol(piLine, i))
        lLastColor = malForeCol(piLine, i)
        fFirstTime = False
      End If
    End If
    i = i + 1
  Loop Until i > miCols
End Sub

Private Sub GenBackColInjections(ByVal piLine As Integer)
  Dim i               As Integer
  Dim fColorChanged   As Boolean
  Dim lLastColor      As Long
  Dim iAttribs        As Integer
  Dim fReset          As Boolean
  Dim fFirstTime      As Boolean
  
  i = 1
  lLastColor = malForeCol(piLine, 1)
  fFirstTime = True
  Do
    iAttribs = maiAttribs(piLine, i)
    fReset = CBool(iAttribs And ATTRIB_RESET)
    If fReset Then fFirstTime = True
    
    If mabBackCol(piLine, i) Then
      fColorChanged = fFirstTime Or CBool(malBackCol(piLine, i) <> lLastColor)
      If fColorChanged Then
        AddInjection i, VT_BCOLOR(malBackCol(piLine, i))
        lLastColor = malBackCol(piLine, i)
        fColorChanged = False
        fFirstTime = False
      End If
    End If
    i = i + 1
  Loop Until i > miCols
End Sub

Private Sub GenAttribsInjections(ByVal piLine As Integer)
  Dim i               As Integer
  Dim k               As Integer
  Dim iLastAttribs    As Integer
  Dim iAttribs        As Integer
  Dim iAtribsToSet    As Integer
  Dim fReset          As Boolean
  
  i = 1
  Do
    iAttribs = maiAttribs(piLine, i)
    fReset = CBool(iAttribs And ATTRIB_RESET)
    If fReset Then
      'clear reset attribute
      iAttribs = iAttribs And (Not ATTRIB_RESET)
      iLastAttribs = 0
    End If
    
    If iAttribs > 0 Then
      If iAttribs <> iLastAttribs Then
        'only an attrib that wasn't set and that is now set generates an injection
        'except for the reset attribute, handled separately
        For k = 0 To ATTRIB_BITS - 1
          If iAttribs And maiAttribMask(k) Then 'if the attrib is set now
            If Not (iLastAttribs And maiAttribMask(k)) Then 'and it wasn't set before
              AddInjection i, masAttribInject(k)
            End If
          End If
        Next k
        iLastAttribs = iAttribs
      End If
    End If 'iAttribs>0
    
    i = i + 1
  Loop Until i > miCols
End Sub

Private Sub GenResetsInjections(ByVal piLine As Integer)
  Dim i               As Integer
  Dim iAttribs        As Integer
  Dim fReset          As Boolean
  
  i = 1
  Do
    iAttribs = maiAttribs(piLine, i)
    fReset = CBool(iAttribs And ATTRIB_RESET)
    If fReset Then
      AddInjection i, VT_RESET()
    End If
    i = i + 1
  Loop Until i > miCols
End Sub

Private Function GetSortKey(ByVal piPos As Integer, ByVal plID As Long) As String
  'Debug.Print Format$(piPos, "0000000000") & Format$(plID, "0000000000")
  GetSortKey = Format$(piPos, "0000000000") & Format$(plID, "0000000000")
End Function

Private Sub SortInjectionsArray(ByVal lower As Integer, ByVal upper As Integer)
  Dim pivot   As TInjection
  Dim Temp    As TInjection
  Dim first   As Integer
  Dim last    As Integer
  Dim middle  As Integer
  
  'Locate pivot
  first = lower
  last = upper
  middle = (first + last) / 2
  pivot = matInjection(middle)
  
  Do  'Move pointers against each other
    While GetSortKey(matInjection(first).iPos, matInjection(first).ID) < GetSortKey(pivot.iPos, pivot.ID)
      first = first + 1
    Wend
    While GetSortKey(matInjection(last).iPos, matInjection(last).ID) > GetSortKey(pivot.iPos, pivot.ID)
      last = last - 1
    Wend
    
    If first <= last Then
      Temp = matInjection(first)
      matInjection(first) = matInjection(last)
      matInjection(last) = Temp
      first = first + 1&
      last = last - 1&
    End If
  Loop Until first > last
  
  If lower < last Then
    Call SortInjectionsArray(lower, last)
  End If
  If first < upper Then
    Call SortInjectionsArray(first, upper)
  End If
End Sub

Public Function GetLineVT100(ByVal piLine As Integer, _
  Optional ByVal pfAddBoundary As Boolean = True) As String
  Dim i       As Integer
  Dim sLine   As String
  Dim iPos    As Integer
  Dim k       As Integer
  Dim sRender As String
  Dim fInject As Boolean
  
  Dim iLenLine        As Integer
  Dim sFragmentPre    As String
  Dim iFragmentSepLen As Integer
  Dim sFragmentSep    As String
  
  Dim iLenInjects   As Integer
  Dim iBufLen       As Integer
  Dim iPoke         As Integer
  Dim iInjectLen    As Integer
  Dim iLenNOOP      As Integer
  Dim sTemp         As String
  
  On Error GoTo GetLineVT100_Err
  
  If moVT100Cache.IsCached(piLine) Then
    GetLineVT100 = moVT100Cache.GetLine(piLine)
    Exit Function
  End If
  
  sFragmentPre = VTX_SETWIDTH(miFragmentCharWidth) & VTX_DTFLAGS(DT_SINGLELINE Or DT_NOPREFIX Or DT_CENTER)
  sFragmentSep = VT_NOOP()
  iFragmentSepLen = Len(sFragmentSep)
  
  sLine = masChars(piLine)  'we start with row text

  InitInjectionsArray
  GenResetsInjections piLine
  GenForeColInjections piLine
  GenBackColInjections piLine
  GenAttribsInjections piLine
  If miInjectionCount > 1 Then
    SortInjectionsArray 1, miInjectionCount
  End If
  'V02.00.06 optimisation, build srender with spaces then poke with mid$
  'Get the total length of injections
  If miInjectionCount > 0 Then
    For i = 1 To miInjectionCount
      iLenInjects = iLenInjects + Len(matInjection(i).sInject)
    Next i
  End If
  iLenLine = Len(sLine)
  iBufLen = Len(sLine) + iLenInjects
  If mfFragmentText Then
    iBufLen = iBufLen + (iLenLine - 1) * iFragmentSepLen
  End If
  sRender = Space$(iBufLen)
  
  iPoke = 1
  k = 1 'index in injections array
  For i = 1 To iLenLine
    fInject = False
    If k <= miInjectionCount Then
      If i = matInjection(k).iPos Then
        fInject = True
      End If
    End If
    
    If fInject Then
      Do While fInject
        'sRender = sRender & matInjection(k).sInject
        iInjectLen = Len(matInjection(k).sInject)
        Mid$(sRender, iPoke, iInjectLen) = matInjection(k).sInject
        iPoke = iPoke + iInjectLen
        If (k + 1) > miInjectionCount Then Exit Do
        k = k + 1
        fInject = CBool(matInjection(k).iPos = i)
      Loop
    End If
    
    'sRender = sRender & Mid$(sLine, i, 1)
    Mid$(sRender, iPoke, 1) = Mid$(sLine, i, 1)
    iPoke = iPoke + 1
    If mfFragmentText Then
      If i < iLenLine Then
        Mid$(sRender, iPoke, iFragmentSepLen) = sFragmentSep
        iPoke = iPoke + iFragmentSepLen
      End If
    End If
  Next i
  'remaining injections after the last char
  If k < miInjectionCount Then
    For i = k To miInjectionCount
      sRender = sRender & matInjection(k).sInject
    Next i
  End If
  'boundary indicator
  If pfAddBoundary Then
    sRender = sRender & VT_RESET() & VTX_SETWIDTH(0) & "¦"
  End If
  If mfFragmentText Then
    sRender = sFragmentPre & sRender
  End If
  
GetLineVT100_Exit:
  GetLineVT100 = sRender
  Exit Function

GetLineVT100_Err:
  SetErr "GetLineVT100", Err.Number, Err.Description
  Resume GetLineVT100_Exit
End Function

Public Sub ClearChar(ByVal piRow As Integer, ByVal piCol As Integer, Optional ByVal pfClearAttribs As Boolean = True, Optional ByVal pfClearColors As Boolean = True)
  If (piRow > 0) And (piRow <= miRows) And (piCol > 0) And (piCol <= miCols) Then
    Mid$(masChars(piRow), piCol, 1) = " "
    If pfClearColors Then
      mabForeCol(piRow, piCol) = 0
      mabBackCol(piRow, piCol) = 0
      malForeCol(piRow, piCol) = 0
      malBackCol(piRow, piCol) = 0
    End If
    If pfClearAttribs Then
      maiAttribs(piRow, piCol) = 0
    End If
    moVT100Cache.InvalidateLine piRow
    SetDirty True
  End If
End Sub

Public Sub Clear()
  CreateBuffers
End Sub

Public Sub Reset()
  Clear
  SetDirty False
End Sub

Public Sub ShiftLeft( _
  ByVal pfShiftText As Boolean, _
  ByVal pfShiftColors As Boolean, _
  ByVal pfShiftAttribs As Boolean, _
  Optional ByVal piStartCol As Integer = 1)
  
  Dim iRow    As Integer
  Dim iCol    As Integer
  Dim sReste As String
  
  For iRow = 1 To miRows
    If pfShiftText Then
      If piStartCol > 1 Then
        sReste = Right$(masChars(iRow), Len(masChars(iRow)) - piStartCol)
        'sReste = Left$(sReste, Len(sReste) - 1)
        'masChars(iRow) = left$(masChars(iRow), piStartCol - 1) & sReste & " "
        Mid$(masChars(iRow), piStartCol, miCols) = sReste & " "
      Else
        'masChars(iRow) = Right$(masChars(iRow), Len(masChars(iRow)) - 1) & " "
        sReste = Right$(masChars(iRow), miCols - 1) & " "
        Mid$(masChars(iRow), 1, miCols) = sReste
      End If
    End If
      
    For iCol = piStartCol To miCols - 1
      If pfShiftColors Then
        malForeCol(iRow, iCol) = malForeCol(iRow, iCol + 1)
        malBackCol(iRow, iCol) = malBackCol(iRow, iCol + 1)
        mabForeCol(iRow, iCol) = mabForeCol(iRow, iCol + 1)
        mabBackCol(iRow, iCol) = mabBackCol(iRow, iCol + 1)
      End If
      If pfShiftAttribs Then
        maiAttribs(iRow, iCol) = maiAttribs(iRow, iCol + 1)
      End If
    Next iCol
  Next iRow
  ClearCol 1, miCols, 1, pfShiftText, pfShiftColors, pfShiftAttribs
  moVT100Cache.Clear
End Sub

Public Sub ShiftRight( _
  ByVal pfShiftText As Boolean, _
  ByVal pfShiftColors As Boolean, _
  ByVal pfShiftAttribs As Boolean, _
  Optional ByVal piStartCol As Integer = 1)
  Dim iRow    As Integer
  Dim iCol    As Integer
  
  For iRow = 1 To miRows
    If pfShiftText Then
      If piStartCol > 1 Then
        masChars(iRow) = left$(masChars(iRow), piStartCol - 1) & " " & Mid$(masChars(iRow), piStartCol, Len(masChars(iRow)) - piStartCol)
      Else
        masChars(iRow) = " " & left$(masChars(iRow), Len(masChars(iRow)) - 1)
      End If
    End If
    For iCol = miCols To piStartCol + 1 Step -1
      If pfShiftColors Then
        malForeCol(iRow, iCol) = malForeCol(iRow, iCol - 1)
        mabForeCol(iRow, iCol) = mabForeCol(iRow, iCol - 1)
        malBackCol(iRow, iCol) = malBackCol(iRow, iCol - 1)
        mabBackCol(iRow, iCol) = mabBackCol(iRow, iCol - 1)
      End If
      If pfShiftAttribs Then
        maiAttribs(iRow, iCol) = maiAttribs(iRow, iCol - 1)
      End If
    Next iCol
  Next iRow
  ClearCol 1, piStartCol, 1, pfShiftText, pfShiftColors, pfShiftAttribs
  moVT100Cache.Clear
End Sub

Public Sub ShiftUp( _
  ByVal pfShiftText As Boolean, _
  ByVal pfShiftColors As Boolean, _
  ByVal pfShiftAttribs As Boolean, _
  Optional ByVal piStartRow As Integer = 1)
  Dim iRow    As Integer
  Dim iCol    As Integer
  
  For iRow = piStartRow To miRows - 1
    If pfShiftText Then
      masChars(iRow) = masChars(iRow + 1)
    End If
    
    For iCol = 1 To miCols
      If pfShiftColors Then
        malForeCol(iRow, iCol) = malForeCol(iRow + 1, iCol)
        mabForeCol(iRow, iCol) = mabForeCol(iRow + 1, iCol)
        malBackCol(iRow, iCol) = malBackCol(iRow + 1, iCol)
        mabBackCol(iRow, iCol) = mabBackCol(iRow + 1, iCol)
      End If
      If pfShiftAttribs Then
        maiAttribs(iRow, iCol) = maiAttribs(iRow + 1, iCol)
      End If
    Next iCol
    moVT100Cache.InvalidateLine iRow
  Next iRow
  ClearLine miRows, 1, 1, pfShiftText, pfShiftColors, pfShiftAttribs 'sets dirty flag
End Sub

Public Sub ShiftDown( _
  ByVal pfShiftText As Boolean, _
  ByVal pfShiftColors As Boolean, _
  ByVal pfShiftAttribs As Boolean, _
  Optional ByVal piStartRow As Integer = 1)
  Dim iRow    As Integer
  Dim iCol    As Integer
  
  For iRow = miRows To piStartRow + 1 Step -1
    If pfShiftText Then
      masChars(iRow) = masChars(iRow - 1)
    End If
    
    For iCol = 1 To miCols
      If pfShiftColors Then
        malForeCol(iRow, iCol) = malForeCol(iRow - 1, iCol)
        mabForeCol(iRow, iCol) = mabForeCol(iRow - 1, iCol)
        malBackCol(iRow, iCol) = malBackCol(iRow - 1, iCol)
        mabBackCol(iRow, iCol) = mabBackCol(iRow - 1, iCol)
      End If
      If pfShiftAttribs Then
        maiAttribs(iRow, iCol) = maiAttribs(iRow - 1, iCol)
      End If
    Next iCol
    moVT100Cache.InvalidateLine iRow
  Next iRow
  ClearLine piStartRow, 1, 1, pfShiftText, pfShiftColors, pfShiftAttribs 'sets dirty flag
End Sub

Public Function ClearLine( _
  ByVal piRow As Integer, _
  ByVal piCol As Integer, _
  ByVal pbDirection As eClearDirection, _
  ByVal pfClearText As Boolean, _
  ByVal pfClearColors As Boolean, _
  ByVal pfClearAttribs As Boolean, _
  Optional ByVal piEndCol As Integer = 0) As Boolean
  If (piRow < 1) Or (piRow > miRows) Then
    Exit Function
  End If

  Dim i       As Integer
  Dim j       As Integer
  Dim iEndCol As Integer
  
  iEndCol = miCols
  If pbDirection = eFromCursorPosition Then
    If piEndCol > 0 Then
      If (piEndCol >= piCol) And (piEndCol <= miCols) Then
        iEndCol = piEndCol
      End If
    End If
  End If
  
  If pfClearText Then
    If pbDirection = eToCursorPosition Then 'clear to col
      If piCol < miCols Then
        Mid$(masChars(piRow), 1, piCol) = Space$(piCol)
      Else
        masChars(piRow) = Space$(miCols)
      End If
    Else  'clear from col
      If (piCol > 1) Or (iEndCol < miCols) Then
        Mid$(masChars(piRow), piCol, iEndCol - piCol + 1) = Space$(iEndCol - piCol + 1)
      Else
        masChars(piRow) = Space$(miCols)
      End If
    End If
  End If
  
  If pfClearColors Or pfClearAttribs Then
    i = piRow
    If pbDirection = eToCursorPosition Then 'to col
      For j = 1 To piCol
        If pfClearColors Then
          malForeCol(i, j) = 0
          mabForeCol(i, j) = False
          malBackCol(i, j) = 0
          mabBackCol(i, j) = False
        End If
        If pfClearAttribs Then
          maiAttribs(i, j) = 0
        End If
      Next j
    Else  'from col
      For j = piCol To iEndCol
        If pfClearColors Then
          malForeCol(i, j) = 0
          mabForeCol(i, j) = False
          malBackCol(i, j) = 0
          mabBackCol(i, j) = False
        End If
        If pfClearAttribs Then
          maiAttribs(i, j) = 0
        End If
      Next j
    End If
  End If
  
  moVT100Cache.InvalidateLine piRow
  SetDirty True
  ClearLine = True
End Function

Public Function ClearCol( _
  ByVal piRow As Integer, _
  ByVal piCol As Integer, _
  ByVal pbDirection As eClearDirection, _
  ByVal pfClearText As Boolean, _
  ByVal pfClearColors As Boolean, _
  ByVal pfClearAttribs As Boolean) As Boolean
  
  If (piCol < 1) Or (piCol > miCols) Then
    Exit Function
  End If

  Dim i       As Integer
  Dim j       As Integer
  
  j = piCol
  If pbDirection = eToCursorPosition Then 'to row
    For i = 1 To piRow
      If pfClearColors Then
        malForeCol(i, j) = 0
        mabForeCol(i, j) = False
        malBackCol(i, j) = 0
        mabBackCol(i, j) = False
      End If
      If pfClearAttribs Then
        maiAttribs(i, j) = 0
      End If
      If pfClearText Then
        Mid$(masChars(i), j, 1) = " "
      End If
    Next i
  Else  'from row
    For i = piRow To miRows
      If pfClearColors Then
        malForeCol(i, j) = 0
        mabForeCol(i, j) = False
        malBackCol(i, j) = 0
        mabBackCol(i, j) = False
      End If
      If pfClearAttribs Then
        maiAttribs(i, j) = 0
      End If
      If pfClearText Then
        Mid$(masChars(i), j, 1) = " "
      End If
    Next i
  End If
  moVT100Cache.Clear
  SetDirty True
  ClearCol = True
End Function

Public Function InsertLine( _
  ByVal pfClearText As Boolean, _
  ByVal pfClearColors As Boolean, _
  ByVal pfClearAttribs As Boolean, _
  ByVal piRow As Integer) As Boolean
  If (piRow < 1) Or (piRow > miRows) Then
    Exit Function
  End If
  
  ShiftDown pfClearText, pfClearColors, pfClearAttribs, piRow
  InsertLine = True
End Function

Public Function DeleteLine( _
  ByVal pfClearText As Boolean, _
  ByVal pfClearColors As Boolean, _
  ByVal pfClearAttribs As Boolean, _
  ByVal piRow As Integer) As Boolean
  If (piRow < 1) Or (piRow > miRows) Then
    Exit Function
  End If
  
  ShiftUp pfClearText, pfClearColors, pfClearAttribs, piRow
  DeleteLine = True
End Function

Public Function InsertCol( _
  ByVal pfClearText As Boolean, _
  ByVal pfClearColors As Boolean, _
  ByVal pfClearAttribs As Boolean, _
  ByVal piCol As Integer) As Boolean
  If (piCol < 1) Or (piCol > miCols) Then
    Exit Function
  End If
  ShiftRight pfClearText, pfClearColors, pfClearAttribs, piCol
  InsertCol = True
End Function

Public Function DeleteCol( _
  ByVal pfClearText As Boolean, _
  ByVal pfClearColors As Boolean, _
  ByVal pfClearAttribs As Boolean, _
  ByVal piCol As Integer) As Boolean
  If (piCol < 1) Or (piCol > miCols) Then
    Exit Function
  End If
  ShiftLeft pfClearText, pfClearColors, pfClearAttribs, piCol
  DeleteCol = True
End Function

'
' Files
'

Private Sub SaveNativeArray_Long(ByVal fh As Integer, ByVal piSectionID As Integer, ByRef palData() As Long, ByVal piRows As Integer, ByVal piCols As Integer)
  Dim iRowLB  As Integer
  Dim iRowUB  As Integer
  Dim iColLB  As Integer
  Dim iColUB  As Integer
  Dim iRow    As Integer
  Dim iCol    As Integer
  Dim lData   As Long
  Dim lMagic  As Long
  
  lMagic = NATIVE_FILE_MAGIC_TAG_V2
  
  'tag with magic long and section id
  Put #fh, , lMagic
  Put #fh, , piSectionID
  'Write the (real vb) dimensions
  iRowLB = LBound(palData, 1)
  iRowUB = UBound(palData, 1)
  iColLB = LBound(palData, 2)
  iColUB = UBound(palData, 2)
  Put #fh, , iRowLB
  Put #fh, , iRowUB
  Put #fh, , iColLB
  Put #fh, , iColUB
  'Write data, one by one
  For iRow = iRowLB To iRowUB
    For iCol = iColLB To iColUB
      lData = palData(iRow, iCol)
      Put #fh, , lData
    Next iCol
  Next iRow
End Sub

Private Sub SaveNativeArray_Byte(ByVal fh As Integer, ByVal piSectionID As Integer, ByRef pabData() As Byte, ByVal piRows As Integer, ByVal piCols As Integer)
  Dim iRowLB  As Integer
  Dim iRowUB  As Integer
  Dim iColLB  As Integer
  Dim iColUB  As Integer
  Dim iRow    As Integer
  Dim iCol    As Integer
  Dim bData   As Byte
  Dim lMagic  As Long
  
  lMagic = NATIVE_FILE_MAGIC_TAG_V2
  
  'tag with magic long and section id
  Put #fh, , lMagic
  Put #fh, , piSectionID
  'Write the (real vb) dimensions
  iRowLB = LBound(pabData, 1)
  iRowUB = UBound(pabData, 1)
  iColLB = LBound(pabData, 2)
  iColUB = UBound(pabData, 2)
  Put #fh, , iRowLB
  Put #fh, , iRowUB
  Put #fh, , iColLB
  Put #fh, , iColUB
  'Write data, one by one
  For iRow = iRowLB To iRowUB
    For iCol = iColLB To iColUB
      bData = pabData(iRow, iCol)
      Put #fh, , bData
    Next iCol
  Next iRow
End Sub

Private Sub SaveNativeArray_Intg(ByVal fh As Integer, ByVal piSectionID As Integer, ByRef paiData() As Integer, ByVal piRows As Integer, ByVal piCols As Integer)
  Dim iRowLB  As Integer
  Dim iRowUB  As Integer
  Dim iColLB  As Integer
  Dim iColUB  As Integer
  Dim iRow    As Integer
  Dim iCol    As Integer
  Dim iData   As Integer
  Dim lMagic  As Long
  
  lMagic = NATIVE_FILE_MAGIC_TAG_V2
  
  'tag with magic long and section id
  Put #fh, , lMagic
  Put #fh, , piSectionID
  'Write the (real vb) dimensions
  iRowLB = LBound(paiData, 1)
  iRowUB = UBound(paiData, 1)
  iColLB = LBound(paiData, 2)
  iColUB = UBound(paiData, 2)
  Put #fh, , iRowLB
  Put #fh, , iRowUB
  Put #fh, , iColLB
  Put #fh, , iColUB
  'Write data, one by one
  For iRow = iRowLB To iRowUB
    For iCol = iColLB To iColUB
      iData = paiData(iRow, iCol)
      Put #fh, , iData
    Next iCol
  Next iRow
End Sub

Private Sub LoadNativeArray_Long( _
    ByVal fh As Integer, _
    ByVal piSectionID As Integer, _
    ByRef palData() As Long, _
    ByVal piRows As Integer, _
    ByVal piCols As Integer, _
    ByVal piSkipRows As Integer, _
    ByVal piSkipCols As Integer _
  )
  Dim iRowLB  As Integer
  Dim iRowUB  As Integer
  Dim iColLB  As Integer
  Dim iColUB  As Integer
  Dim iRow    As Integer
  Dim iCol    As Integer
  Dim lData   As Long
  Dim lMagic  As Long
  Dim iSect   As Integer
  Dim iCalc   As Integer
  Dim iReadRowStop  As Integer
  Dim iReadColStop  As Integer
  
  Const LOCAL_ERR_CTX As String = "LoadNativeArray_Long"
  
  'magic and section id
  Get #fh, , lMagic
  If lMagic <> NATIVE_FILE_MAGIC_TAG_V2 Then
    Err.Raise 52&, LOCAL_ERR_CTX, "Invalid magic marker (at location #" & Seek(fh) & ") for section id #" & piSectionID
    Exit Sub
  End If
  Get #fh, , iSect
  If iSect <> piSectionID Then
    Err.Raise LOCAL_ERR_CTX, 52&, "Invalid section id #" & iSect & " (expected #" & piSectionID & ") (at location #" & Seek(fh) & ")" '52 in vba is "bad filename or number"
    Exit Sub
  End If
  'get array bounds
  Get #fh, , iRowLB
  Get #fh, , iRowUB
  Get #fh, , iColLB
  Get #fh, , iColUB
  'Check expected bounds
  iCalc = iRowUB - iRowLB + 1
  If iCalc <> piRows Then
    If iCalc < piRows Then
      Err.Raise 52&, LOCAL_ERR_CTX, "Row count mismatch, expected #" & piRows & ", got #" & iCalc
      Exit Sub
    Else
      'should log/trace/warn
    End If
  End If
  iCalc = iColUB - iColLB + 1
  If iCalc <> piCols Then
    If iCalc < piCols Then
      Err.Raise 52&, LOCAL_ERR_CTX, "Column count mismatch, expected #" & piCols & ", got #" & iCalc
      Exit Sub
    Else
      'should log/trace/warn
    End If
  End If
  
  'Read data, one by one
  iReadRowStop = piRows - piSkipRows
  iReadColStop = piCols - piSkipCols
  For iRow = iRowLB To iRowUB
    For iCol = iColLB To iColUB
      Get #fh, , lData
      If (iRow <= iReadRowStop) And (iCol <= iReadColStop) Then
        palData(iRow, iCol) = lData
      End If
    Next iCol
  Next iRow
End Sub

Private Sub LoadNativeArray_Byte( _
    ByVal fh As Integer, _
    ByVal piSectionID As Integer, _
    ByRef pabData() As Byte, _
    ByVal piRows As Integer, _
    ByVal piCols As Integer, _
    ByVal piSkipRows As Integer, _
    ByVal piSkipCols As Integer _
  )
  Dim iRowLB  As Integer
  Dim iRowUB  As Integer
  Dim iColLB  As Integer
  Dim iColUB  As Integer
  Dim iRow    As Integer
  Dim iCol    As Integer
  Dim bData   As Byte
  Dim lMagic  As Long
  Dim iSect   As Integer
  Dim iCalc   As Integer
  Dim iReadRowStop  As Integer
  Dim iReadColStop  As Integer
  
  Const LOCAL_ERR_CTX As String = "LoadNativeArray_Byte"
  
  'magic and section id
  Get #fh, , lMagic
  If lMagic <> NATIVE_FILE_MAGIC_TAG_V2 Then
    Err.Raise 52&, LOCAL_ERR_CTX, "Invalid magic marker at file pointer location #" & Seek(fh) & " for section id #" & piSectionID
    Exit Sub
  End If
  Get #fh, , iSect
  If iSect <> piSectionID Then
    Err.Raise LOCAL_ERR_CTX, 52&, "Invalid section id #" & iSect & " (expected #" & piSectionID & ") at file pointer location #" & Seek(fh) '52 in vba is "bad filename or number"
    Exit Sub
  End If
  'get array bounds
  Get #fh, , iRowLB
  Get #fh, , iRowUB
  Get #fh, , iColLB
  Get #fh, , iColUB
  'Check expected bounds
  iCalc = iRowUB - iRowLB + 1
  If iCalc <> piRows Then
    If iCalc < piRows Then
      Err.Raise 52&, LOCAL_ERR_CTX, "Row count mismatch, expected #" & piRows & ", got #" & iCalc
      Exit Sub
    Else
      'should log/trace/warn
    End If
  End If
  iCalc = iColUB - iColLB + 1
  If iCalc <> piCols Then
    If iCalc < piCols Then
      Err.Raise 52&, LOCAL_ERR_CTX, "Column count mismatch, expected #" & piCols & ", got #" & iCalc
      Exit Sub
    Else
      'should log/trace/warn
    End If
  End If
  
  'Read data, one by one
  iReadRowStop = piRows - piSkipRows
  iReadColStop = piCols - piSkipCols
  For iRow = iRowLB To iRowUB
    For iCol = iColLB To iColUB
      Get #fh, , bData
      If (iRow <= iReadRowStop) And (iCol <= iReadColStop) Then
        pabData(iRow, iCol) = bData
      End If
    Next iCol
  Next iRow
End Sub

Private Sub LoadNativeArray_Intg( _
    ByVal fh As Integer, _
    ByVal piSectionID As Integer, _
    ByRef paiData() As Integer, _
    ByVal piRows As Integer, _
    ByVal piCols As Integer, _
    ByVal piSkipRows As Integer, _
    ByVal piSkipCols As Integer _
  )
  Dim iRowLB  As Integer
  Dim iRowUB  As Integer
  Dim iColLB  As Integer
  Dim iColUB  As Integer
  Dim iRow    As Integer
  Dim iCol    As Integer
  Dim iData   As Integer
  Dim lMagic  As Long
  Dim iSect   As Integer
  Dim iCalc   As Integer
  Dim iReadRowStop  As Integer
  Dim iReadColStop  As Integer
  
  Const LOCAL_ERR_CTX As String = "LoadNativeArray_Intg"
  
  'magic and section id
  Get #fh, , lMagic
  If lMagic <> NATIVE_FILE_MAGIC_TAG_V2 Then
    Err.Raise 52&, LOCAL_ERR_CTX, "Invalid magic marker at file pointer location #" & Seek(fh) & " for section id #" & piSectionID
    Exit Sub
  End If
  Get #fh, , iSect
  If iSect <> piSectionID Then
    Err.Raise LOCAL_ERR_CTX, 52&, "Invalid section id #" & iSect & " (expected #" & piSectionID & ") at file pointer location #" & Seek(fh) '52 in vba is "bad filename or number"
    Exit Sub
  End If
  'get array bounds
  Get #fh, , iRowLB
  Get #fh, , iRowUB
  Get #fh, , iColLB
  Get #fh, , iColUB
  'Check expected bounds
  iCalc = iRowUB - iRowLB + 1
  If iCalc <> piRows Then
    If iCalc < piRows Then
      Err.Raise 52&, LOCAL_ERR_CTX, "Row count mismatch, expected #" & piRows & ", got #" & iCalc
      Exit Sub
    Else
      'should log/trace/warn
    End If
  End If
  iCalc = iColUB - iColLB + 1
  If iCalc <> piCols Then
    If iCalc < piCols Then
      Err.Raise 52&, LOCAL_ERR_CTX, "Column count mismatch, expected #" & piCols & ", got #" & iCalc
      Exit Sub
    Else
      'should log/trace/warn
    End If
  End If
  
  'Read data, one by one
  iReadRowStop = piRows - piSkipRows
  iReadColStop = piCols - piSkipCols
  For iRow = iRowLB To iRowUB
    For iCol = iColLB To iColUB
      Get #fh, , iData
      If (iRow <= iReadRowStop) And (iCol <= iReadColStop) Then
        paiData(iRow, iCol) = iData
      End If
    Next iCol
  Next iRow
End Sub

'V02.00.06 Adding optional max rows/cols for ability to get "snapshot"
'of the file, to use it in a library browser.
Public Function LoadFromNativeStream( _
    ByVal fh As Integer, _
    Optional ByVal piMaxRows As Integer = 0, _
    Optional ByVal piMaxCols As Integer = 0, _
    Optional ByVal pfAutoResize As Boolean = True _
  ) As Boolean
  Dim iSavedRows    As Integer
  Dim iSavedCols    As Integer
  Dim i             As Integer
  Dim lMagicTag     As Long
  Dim lClassVersion As Long
  Dim iRows         As Integer
  Dim iCols         As Integer
  Dim lMagic        As Long
  Dim iSectionID    As Integer
  Dim iSkipRows     As Integer
  Dim iSkipCols     As Integer
  Dim iSizeRows     As Integer
  Dim iSizeCols     As Integer
  Dim iReadRowStop  As Integer
  Dim iReadColStop  As Integer
  Dim sBuffer       As String
  
  Const LOCAL_ERR_CTX As String = "LoadFromNativeStream"
  
  On Error GoTo LoadFromNativeStream_Err
  ClearErr
  
  iSavedRows = miRows
  iSavedCols = miCols
  
  Get #fh, , lMagicTag 'NATIVE_FILE_MAGIC_TAG
  'We keep reading unsupported V1 tags, just to get to the unsupported version error
  If (lMagicTag <> NATIVE_FILE_MAGIC_TAG_V1) And (lMagicTag <> NATIVE_FILE_MAGIC_TAG_V2) Then
    SetErr LOCAL_ERR_CTX, 52&, "The file is not a native " & APP_NAME & " data file" '52 in vba is "bad filename or number"
    GoTo LoadFromNativeStream_Exit
  End If
  
  Get #fh, , lClassVersion 'CLASS_VERSION
  If lClassVersion > CLASS_VERSION Then
    SetErr LOCAL_ERR_CTX, 52&, "The data file version is greater than this application supports" '52 in vba is "bad filename or number"
    GoTo LoadFromNativeStream_Exit
  End If
  If lClassVersion < MIN_SUPPORTED_VERSION Then 'V02.00.06 dropping support for older versions
    SetErr LOCAL_ERR_CTX, 52&, "This data file version (" & lClassVersion & ") is not supported by this application" '52 in vba is "bad filename or number"
    GoTo LoadFromNativeStream_Exit
  End If
  
  Get #fh, , iRows
  Get #fh, , iCols
  If (iRows <= 0) Or (iCols <= 0) Then
    SetErr LOCAL_ERR_CTX, 52&, "Incorrect grid dimensions (rows=" & iRows & ",cols=" & iCols & ") (at position #" & Seek(fh) & ")" '52 in vba is "bad filename or number"
    GoTo LoadFromNativeStream_Exit
  End If
  
  If pfAutoResize Or (Not mfBuffersCreated) Then
    iSkipRows = 0
    iSizeRows = iRows
    If piMaxRows > 0 Then
      If iRows > piMaxRows Then
        iSkipRows = iRows - piMaxRows
        iSizeRows = piMaxRows
      End If
    End If
    
    iSkipCols = 0
    iSizeCols = iCols
    If piMaxCols > 0 Then
      If iCols > piMaxCols Then
        iSkipCols = iCols - piMaxCols
        iSizeCols = piMaxCols
      End If
    End If
    
    If Not Me.Resize(iSizeRows, iSizeCols) Then 'will set the error context if ever
      GoTo LoadFromNativeStream_Exit
    End If
  Else
    If piMaxRows > 0 Then
      If miRows > piMaxRows Then
        iSkipRows = miRows - piMaxRows
      End If
    End If
    If piMaxCols > 0 Then
      If miCols > piMaxCols Then
        iSkipCols = miCols - piMaxCols
      End If
    End If
  End If
  
  LoadNativeArray_Long fh, SECTIONID_FORECOLOR_COLOR, malForeCol(), iRows, iCols, iSkipRows, iSkipCols
  LoadNativeArray_Byte fh, SECTIONID_FORECOLOR_ONOFF, mabForeCol(), iRows, iCols, iSkipRows, iSkipCols
  LoadNativeArray_Long fh, SECTIONID_BACKCOLOR_COLOR, malBackCol(), iRows, iCols, iSkipRows, iSkipCols
  LoadNativeArray_Byte fh, SECTIONID_BACKCOLOR_ONOFF, mabBackCol(), iRows, iCols, iSkipRows, iSkipCols
  LoadNativeArray_Intg fh, SECTIONID_ATTRIBUTES, maiAttribs(), iRows, iCols, iSkipRows, iSkipCols
  
  'Debug.Print "LoadFromNativeStream: reading char buffer at pos #"; Seek(fh)
  Get #fh, , lMagic
  If lMagic <> NATIVE_FILE_MAGIC_TAG_V2 Then
    SetErr LOCAL_ERR_CTX, 52&, "Invalid magic marker (at position #" & Seek(fh) & ")" '52 in vba is "bad filename or number"
    GoTo LoadFromNativeStream_Exit
  End If
  Get #fh, , iSectionID
  If iSectionID <> SECTIONID_CHARWDATA Then
    SetErr LOCAL_ERR_CTX, 52&, "Invalid section id #" & iSectionID & " (expected #" & SECTIONID_CHARWDATA & ") (at position #" & Seek(fh) & ")" '52 in vba is "bad filename or number"
    GoTo LoadFromNativeStream_Exit
  End If
  
  'read char data lines
  iReadRowStop = iRows - iSkipRows
  iReadColStop = iCols - iSkipCols
  For i = 1 To iRows
    sBuffer = FileGetUnicodeString(fh, iReadColStop)
    If i <= iReadRowStop Then
      'We have to "poke" the thing. NO: masChars(i) = sBuffer
      Mid$(masChars(i), 1, Len(sBuffer)) = sBuffer
      If Len(sBuffer) = 0 Then
        SetErr LOCAL_ERR_CTX, 52&, "Canvas data row #" & i & " has been read as empty (at position #" & Seek(fh) & ")" '52 in vba is "bad filename or number"
        GoTo LoadFromNativeStream_Exit
      End If
    End If
  Next i

  SetDirty False
  moVT100Cache.Clear
  LoadFromNativeStream = True
  
LoadFromNativeStream_Exit:
  Exit Function

LoadFromNativeStream_Err:
  SetErr LOCAL_ERR_CTX, Err.Number, Err.Description
  On Error Resume Next
  miRows = iSavedRows: miCols = iSavedCols: Clear
End Function

Public Function LoadNative( _
    ByVal psFilename As String, _
    Optional ByVal piMaxRows As Integer = 0, _
    Optional ByVal piMaxCols As Integer = 0, _
    Optional ByVal pfAutoResize As Boolean = True _
  ) As Boolean
  Dim fh      As Integer
  Dim fIsOpen As Boolean
  Dim fOK     As Boolean
  
  On Error GoTo LoadNative_Err
  ClearErr
  
  fh = FreeFile
  Open psFilename For Binary Access Read Lock Write As #fh
  fIsOpen = True
  
  fOK = LoadFromNativeStream(fh, piMaxRows, piMaxCols, pfAutoResize)
  
  If pfAutoResize Then
    SetDirty False
  Else
    SetDirty True
  End If
  
  LoadNative = fOK
  
LoadNative_Exit:
  If fIsOpen Then
    Close #fh
  End If
  Exit Function

LoadNative_Err:
  SetErr "LoadNative", Err.Number, Err.Description
  Resume LoadNative_Exit
  Resume
End Function

Public Function SaveToNativeStream(ByVal fh As Integer) As Boolean
  Dim i       As Integer
  Dim sLine   As String
  
  On Error GoTo SaveToNativeStream_Err
  ClearErr
  
  Put #fh, , NATIVE_FILE_MAGIC_TAG_V2
  Put #fh, , CLASS_VERSION
  Put #fh, , miRows
  Put #fh, , miCols
  SaveNativeArray_Long fh, SECTIONID_FORECOLOR_COLOR, malForeCol(), miRows, miCols
  SaveNativeArray_Byte fh, SECTIONID_FORECOLOR_ONOFF, mabForeCol(), miRows, miCols
  SaveNativeArray_Long fh, SECTIONID_BACKCOLOR_COLOR, malBackCol(), miRows, miCols
  SaveNativeArray_Byte fh, SECTIONID_BACKCOLOR_ONOFF, mabBackCol(), miRows, miCols
  SaveNativeArray_Intg fh, SECTIONID_ATTRIBUTES, maiAttribs(), miRows, miCols
  'Debug.Print "SaveToNativeStream: char buffer at pos #" & Seek(fh)
  'put markers for string
  Put #fh, , NATIVE_FILE_MAGIC_TAG_V2
  i = SECTIONID_CHARWDATA
  Put #fh, , i
  For i = 1 To miRows
    FilePutUnicodeString fh, masChars(i)
  Next i
  
  SetDirty False
  SaveToNativeStream = True
  
SaveToNativeStream_Exit:
  Exit Function

SaveToNativeStream_Err:
  SetErr "SaveToNativeStream", Err.Number, Err.Description
  Resume SaveToNativeStream_Exit
End Function

Public Function SaveVT100(ByVal psFilename As String) As Boolean
  Dim fh      As Integer
  Dim fIsOpen As Boolean
  Dim i       As Integer
  Dim sLine   As String
  Dim sEscaped  As String
  
  On Error GoTo SaveVT100_Err
  ClearErr
  
  'Open file for output (overwrite)
  fh = FreeFile
  
  'Open psFilename For Binary Access Write Lock Read Write As #fh
  'fIsOpen = True
  'Put #fh, , (CByte(255))
  'Put #fh, , (CByte(254))
  'Close #fh
  'fIsOpen = False
  
  'Open psFilename For Append Access Write Lock Read Write As #fh
  Open psFilename For Output Access Write Lock Read Write As #fh
  fIsOpen = True
  Print #fh, StrConv(ChrW$(&HFEFF), vbUnicode);
  
  For i = 1 To miRows
    sLine = GetLineVT100(i, False)
    If i < miRows Then
      sLine = sLine & vbCrLf
    End If
    Print #fh, StrConv(sLine, vbUnicode);
    'FilePutUnicodeString fh, sLine, False
  Next i
  
  'Close the file
  Close fh
  fIsOpen = False
  
  'if saved, clear dirty flag
  SetDirty False

  SaveVT100 = True
  
SaveVT100_Exit:
  If fIsOpen Then
    Close fh
  End If
  Exit Function

SaveVT100_Err:
  SetErr "SaveVT100", Err.Number, Err.Description
  Resume SaveVT100_Exit
  Resume
End Function

Private Sub ParseEscSeq(ByVal psEscSeq As String, ByRef piRetCmdCode As Integer, ByRef psRetParams As String)
  Dim iLen        As Integer
  Dim sCode       As String
  Dim wRed        As Integer
  Dim wGreen      As Integer
  Dim wBlue       As Integer
  Dim iSepCt      As Integer
  Dim lColorValue As Long
  Dim i           As Integer
  
  piRetCmdCode = 0
  psRetParams = ""
  
  iLen = Len(psEscSeq)
  If iLen = 0 Then
    Exit Sub
  End If
  
  i = InStr(1, psEscSeq, ";")
  If i > 0 Then
    sCode = GetStringPart(1, ";", psEscSeq)
    psEscSeq = GetStringPart(2, ";", psEscSeq)
  Else
    sCode = left$(psEscSeq, 2)
  End If
  
  Select Case sCode
  Case VT100_RESET
    piRetCmdCode = LSCMD_RESET
  Case VT100_INVERSEON
    piRetCmdCode = LSCMD_INVERSEON
  Case VT100_INVERSEOFF
    piRetCmdCode = LSCMD_INVERSEOFF
  Case VT100_BOLDON
    piRetCmdCode = LSCMD_BOLDON
  Case VT100_BOLDOFF
    piRetCmdCode = LSCMD_BOLDOFF
  Case VT100_ITALICON
    piRetCmdCode = LSCMD_ITALICON
  Case VT100_ITALICOFF
    piRetCmdCode = LSCMD_ITALICOFF
  Case VT100_UNDERLINEON
    piRetCmdCode = LSCMD_UNDERLINEON
  Case VT100_UNDERLINEOFF
    piRetCmdCode = LSCMD_UNDERLINEOFF
  Case VT100_CROSSEDOUT_ON
    piRetCmdCode = LSCMD_CROSSEDOUTON
  Case VT100_CROSSEDOUT_OFF
    piRetCmdCode = LSCMD_CROSSEDOUTOFF
  Case VT100_SETFGCOLOR, VT100_SETBKCOLOR
    iSepCt = CountStringParts(psEscSeq, ";")
    If (iSepCt = 1) Or (iSepCt = 3) Then ' either color value (assumed in decimal not hex) or r;g;b
      If (iSepCt = 1) Then
        lColorValue = Val(Replace(psEscSeq, "$", "&H"))
      Else
        wRed = Val(GetStringPart(1, ";", Replace(psEscSeq, "$", "&H")))
        wGreen = Val(GetStringPart(2, ";", Replace(psEscSeq, "$", "&H")))
        wBlue = Val(GetStringPart(3, ";", Replace(psEscSeq, "$", "&H")))
        lColorValue = RGB(wRed, wGreen, wBlue)
      End If
      psRetParams = CStr(lColorValue)
      If sCode = VT100_SETFGCOLOR Then
        piRetCmdCode = LSCMD_SETFORECOLOR
      Else
        piRetCmdCode = LSCMD_SETBACKCOLOR
      End If
    End If
  Case VT100X_BEGINZONE
    psRetParams = CStr(Val(psEscSeq))
    piRetCmdCode = LSCMD_BEGINZONE
  Case VT100X_ENDZONE
    psRetParams = CStr(Val(psEscSeq))
    piRetCmdCode = LSCMD_ENDZONE
  Case VT100X_BKCOLSPILL
    psRetParams = CStr(Val(psEscSeq))
    piRetCmdCode = LSCMD_BKCOLSPILL
  End Select

End Sub

'it's all one based, this is internal, so caller obligation
Private Sub SetLineText( _
    ByVal piRow As Integer, _
    ByVal piCol As Integer, _
    ByVal psText As String, _
    Optional ByVal pfTransparent As Boolean = False _
  )
  Dim iLenText As Integer
  Dim iPokeLen As Integer
  iLenText = Len(psText)
  iPokeLen = miCols - piCol + 1
  If iPokeLen > iLenText Then
    iPokeLen = iLenText
  End If
  
  If Not pfTransparent Then
    Mid$(masChars(piRow), piCol, iPokeLen) = left$(psText, iPokeLen)
  Else
    Dim i As Integer
    Dim iChar As Integer
    For i = 1 To iPokeLen
      iChar = AscW(Mid$(psText, i, 1))
      If iChar <> 32 Then
        Mid$(masChars(piRow), piCol + i - 1, 1) = ChrW$(iChar)
      End If
    Next i
  End If
  
  SetDirty True
  moVT100Cache.InvalidateLine piRow
End Sub

Public Sub ParseVT100( _
    ByVal piRow As Integer, _
    ByVal piCol As Integer, _
    ByVal psText As String, _
    ByRef piRetCols As Integer, _
    Optional ByVal pfClear As Boolean = True, _
    Optional ByVal pfTransparent As Boolean = False _
  )
  If Len(psText) = 0 Then
    If pfClear Then
      ClearLine piRow, piCol, 1, True, True, True
    End If
    Exit Sub
  End If
  
  On Error GoTo ParseVT100_Err
  
  Dim CSI           As String
  Dim iFind         As Integer
  Dim iPos          As Integer
  Dim iEnd          As Integer
  Dim iLen          As Integer
  Dim sEscSeq       As String
  Dim iCmdCode      As Integer
  Dim sParams       As String
  Dim sText         As String
  Dim sBuffer       As String
  Dim k             As Integer
  Dim iAttribs      As Integer
  
  If pfClear Then
    ClearLine piRow, piCol, 1, True, True, True
  End If
  piRetCols = 0
  
  If piRow > miRows Then
    Exit Sub
  End If
  
  iLen = Len(psText)
  CSI = Chr$(27) & "["
  iFind = InStr(1, psText, CSI)
  If iFind > 0 Then
    iPos = 1
    Do While iFind > 0
      sEscSeq = ""
      iEnd = InStr(iFind + 1, psText, "m")
      If iEnd > 0 Then
        If (iEnd - (iFind + 2)) > 0 Then
          sEscSeq = Mid$(psText, iFind + 2, iEnd - (iFind + 2)) 'doesn't include the "m"
        Else
          sEscSeq = ""
        End If
        ParseEscSeq sEscSeq, iCmdCode, sParams
        If iCmdCode <> LSCMD_NOOP Then
          If iPos < (iFind - 1) Then
            sText = Mid(psText, iPos + 1, (iFind - iPos - 1))
            'push text out
            sBuffer = sBuffer & sText
          End If
          'push LSCMD+params
          k = Len(sBuffer) + 1 'column where the code applies
          If k <= miCols Then
            iAttribs = maiAttribs(piRow, k)
            Select Case iCmdCode
            Case LSCMD_INVERSEON
              iAttribs = iAttribs Or ATTRIB_INVERSEON
              maiAttribs(piRow, k) = iAttribs
            Case LSCMD_INVERSEOFF
              iAttribs = iAttribs Or ATTRIB_INVERSEOFF
              maiAttribs(piRow, k) = iAttribs
            Case LSCMD_BOLDON
              iAttribs = iAttribs Or ATTRIB_BOLDON
              maiAttribs(piRow, k) = iAttribs
            Case LSCMD_ITALICON
              iAttribs = iAttribs Or ATTRIB_ITALICON
              maiAttribs(piRow, k) = iAttribs
            Case LSCMD_UNDERLINEON
              iAttribs = iAttribs Or ATTRIB_UNDLON
              maiAttribs(piRow, k) = iAttribs
            Case LSCMD_BOLDOFF
              iAttribs = iAttribs Or ATTRIB_BOLDOFF
              maiAttribs(piRow, k) = iAttribs
            Case LSCMD_ITALICOFF
              iAttribs = iAttribs Or ATTRIB_ITALICOFF
              maiAttribs(piRow, k) = iAttribs
            Case LSCMD_UNDERLINEOFF
              iAttribs = iAttribs Or ATTRIB_UNDLOFF
              maiAttribs(piRow, k) = iAttribs
            Case LSCMD_CROSSEDOUTON
              iAttribs = iAttribs Or ATTRIB_STRIKEON
              maiAttribs(piRow, k) = iAttribs
            Case LSCMD_CROSSEDOUTOFF
              iAttribs = iAttribs Or ATTRIB_STRIKEOFF
              maiAttribs(piRow, k) = iAttribs
            Case LSCMD_SETFORECOLOR 'color is a long value in sText
              malForeCol(piRow, k) = Val(sParams)
              mabForeCol(piRow, k) = True
            Case LSCMD_SETBACKCOLOR
              malBackCol(piRow, k) = Val(sParams)
              mabBackCol(piRow, k) = True
            Case LSCMD_BEGINZONE  'Not supported, do nothing
            Case LSCMD_ENDZONE    'Not supported, do nothing
            Case LSCMD_RESET
              iAttribs = iAttribs Or ATTRIB_RESET
              maiAttribs(piRow, k) = iAttribs
            Case LSCMD_BKCOLSPILL
              If CBool(Val(sParams)) Then
                iAttribs = iAttribs Or ATTRIB_SPILLON
              Else
                iAttribs = iAttribs Or ATTRIB_SPILLOFF
              End If
              maiAttribs(piRow, k) = iAttribs
            End Select
          End If
          iPos = iEnd ' meaning we processed the text until iPos
          iFind = InStr(iEnd + 1, psText, CSI) ' find next escape sequence intro
        Else
          ' ignore escape sequence (although will be kept in output string)
          If (iFind - (iPos + 1)) > 0 Then
            sBuffer = sBuffer & Mid$(psText, iPos + 1, iFind - (iPos + 1))
          End If
          iPos = iEnd ' meaning we processed the text until iPos
          iFind = InStr(iEnd + 1, psText, CSI) ' find next escape sequence intro
        End If
      End If
    Loop
    If iPos < iLen Then
      sText = Mid(psText, iPos + 1, iLen - iPos)
      sBuffer = sBuffer & sText
    End If
    
    'Debug.Print Format$(piRow, "000"); " ["; sBuffer; "]"
    'sBuffer is/has the exact number of columns of text output w/o vt100 escapes
    piRetCols = Len(sBuffer)
    SetLineText piRow, piCol, sBuffer, pfTransparent
  Else
    'no escape codes, get just the text
    piRetCols = Len(psText)
    SetLineText piRow, piCol, psText, pfTransparent
  End If

ParseVT100_Exit:
  Exit Sub
  
ParseVT100_Err:
  Resume ParseVT100_Exit
  Resume
End Sub

Public Function LoadVT100( _
    ByVal psFilename As String, _
    ByRef piRetFileRows As Integer, _
    ByRef piRetFileCols As Integer, _
    Optional ByVal piStartRow As Integer = 1, _
    Optional ByVal piStartCol As Integer = 1, _
    Optional ByVal pfClear As Boolean = True, _
    Optional ByVal pfTransparent As Boolean = False _
  ) As Boolean
  Dim fh      As Integer
  Dim fIsOpen As Boolean
  Dim i       As Integer
  Dim sLine   As String
  
  Dim iArraySize  As Integer
  Dim iLineCt     As Integer
  Dim iColLen     As Integer
  Dim sBuffer     As String
  
  ReDim asLine(1 To 100) As String
  
  On Error GoTo LoadVT100_Err
  ClearErr
  
  'We read in binary mode (expecting a vb unicode file)
  fh = FreeFile
  Open psFilename For Binary Access Read Lock Write As #fh
  fIsOpen = True
  
  If pfClear Then
    Clear
  End If
  
  'Load into string array
  piRetFileRows = 0
  piRetFileCols = 0
  
  ReDim abBytes(1 To LOF(fh)) As Byte
  Get #fh, , abBytes()
  sBuffer = StrConv(abBytes(), vbUnicode)
  iLineCt = SplitString(asLine(), sBuffer, vbCrLf)
  piRetFileRows = iLineCt
  
  'We have to cycle thru all the lines,
  'just to avoid missing the longest one.
  For i = 1 To iLineCt
    sBuffer = asLine(i)
    'ParseVT100 will do all the bounds checking
    ParseVT100 piStartRow + i - 1, piStartCol, asLine(i), iColLen, pfClear, pfTransparent
    If iColLen > piRetFileCols Then
      piRetFileCols = iColLen
    End If
  Next i
  
  'Close the file
  Close fh
  fIsOpen = False
  
  'if loaded, clear dirty flag
  If pfClear Then
    SetDirty False
  Else
    SetDirty True
  End If
  
  LoadVT100 = True
  
LoadVT100_Exit:
  If fIsOpen Then
    Close fh
  End If
  Exit Function

LoadVT100_Err:
  SetErr "LoadVT100", Err.Number, Err.Description
  Resume LoadVT100_Exit
  Resume
End Function

Public Function ColorsToPalette(ByRef poPalette As CColorPalette) As Boolean
  Dim i         As Integer
  Dim j         As Integer
  
  For i = 1 To miRows
    For j = 1 To miCols
      If mabForeCol(i, j) Then
        poPalette.AddColor malForeCol(i, j)
        'Debug.Print "AddColor " & Hex$(malForeCol(i, j))
        'DoEvents
      End If
      If mabBackCol(i, j) Then
        poPalette.AddColor malBackCol(i, j)
        'Debug.Print "AddColor " & Hex$(malBackCol(i, j))
        'DoEvents
      End If
    Next j
  Next i

  ColorsToPalette = True
  
ColorsToPalette_Exit:
  Exit Function

ColorsToPalette_Err:
  SetErr "ColorsToPalette", Err.Number, Err.Description
  Resume ColorsToPalette_Exit
End Function

Public Function GetCharactersPlane( _
  ByRef pasRetData() As String, _
  ByVal piStartRow As Integer, _
  ByVal piEndRow As Integer, _
  ByVal piStartCol As Integer, _
  ByVal piEndCol As Integer) As Boolean
  
  On Error GoTo GetCharactersPlane_Err
  ClearErr

  Dim i               As Integer
  Dim j               As Integer
  Dim iDstRowCount    As Integer
  Dim iLength         As Integer
  
  iDstRowCount = (piEndRow - piStartRow) + 1
  If piEndCol > miCols Then
    piEndCol = miCols
  End If
  iLength = (piEndCol - piStartCol) + 1
  
  ReDim pasRetData(1 To iDstRowCount) As String
  If (piStartCol > 1) Or (piEndCol < miCols) Then
    'extract and copy the substrings
    j = 1
    For i = piStartRow To piEndRow
      pasRetData(j) = Mid$(masChars(i), piStartCol, iLength)
      j = j + 1
    Next i
  Else
    'we just copy the lines, so the loop is simpler
    j = 1
    For i = piStartRow To piEndRow
      pasRetData(j) = masChars(i)
      j = j + 1
    Next i
  End If
  
  GetCharactersPlane = True
  
GetCharactersPlane_Exit:
  Exit Function

GetCharactersPlane_Err:
  SetErr "GetCharactersPlane", Err.Number, Err.Description
  Resume GetCharactersPlane_Exit
  Resume
End Function

Public Function GetForecolorPlane( _
  ByRef palRetData() As Long, _
  ByRef pabRetUsed() As Byte, _
  ByVal piStartRow As Integer, _
  ByVal piEndRow As Integer, _
  ByVal piStartCol As Integer, _
  ByVal piEndCol As Integer) As Boolean
  
  On Error GoTo GetForecolorPlane_Err
  ClearErr

  Dim i               As Integer
  Dim k               As Integer
  Dim iLine           As Integer
  Dim iCol            As Integer
  Dim iDstRowCount    As Integer
  Dim iLength         As Integer
  
  iDstRowCount = (piEndRow - piStartRow) + 1
  If piEndCol > miCols Then
    piEndCol = miCols
  End If
  iLength = (piEndCol - piStartCol) + 1
  
  ReDim palRetData(1 To iDstRowCount, 1 To iLength) As Long
  ReDim pabRetUsed(1 To iDstRowCount, 1 To iLength) As Byte
  
  'extract and copy data
  iLine = 1 'result line index
  For i = piStartRow To piEndRow
    iCol = 1
    For k = piStartCol To piEndCol
      palRetData(iLine, iCol) = malForeCol(i, k)
      pabRetUsed(iLine, iCol) = mabForeCol(i, k)
      iCol = iCol + 1
    Next k
    iLine = iLine + 1
  Next i
  
  GetForecolorPlane = True
  
GetForecolorPlane_Exit:
  Exit Function

GetForecolorPlane_Err:
  SetErr "GetForecolorPlane", Err.Number, Err.Description
  Resume GetForecolorPlane_Exit
End Function

Public Function GetBackcolorPlane( _
  ByRef palRetData() As Long, _
  ByRef pabRetUsed() As Byte, _
  ByVal piStartRow As Integer, _
  ByVal piEndRow As Integer, _
  ByVal piStartCol As Integer, _
  ByVal piEndCol As Integer) As Boolean
  
  On Error GoTo GetBackcolorPlane_Err
  ClearErr

  Dim i               As Integer
  Dim k               As Integer
  Dim iLine           As Integer
  Dim iCol            As Integer
  Dim iDstRowCount    As Integer
  Dim iLength         As Integer
  
  iDstRowCount = (piEndRow - piStartRow) + 1
  If piEndCol > miCols Then
    piEndCol = miCols
  End If
  iLength = (piEndCol - piStartCol) + 1
  
  ReDim palRetData(1 To iDstRowCount, 1 To iLength) As Long
  ReDim pabRetUsed(1 To iDstRowCount, 1 To iLength) As Byte
  
  'extract and copy data
  iLine = 1 'result line index
  For i = piStartRow To piEndRow
    iCol = 1
    For k = piStartCol To piEndCol
      palRetData(iLine, iCol) = malBackCol(i, k)
      pabRetUsed(iLine, iCol) = mabBackCol(i, k)
      iCol = iCol + 1
    Next k
    iLine = iLine + 1
  Next i
  
  GetBackcolorPlane = True
  
GetBackcolorPlane_Exit:
  Exit Function

GetBackcolorPlane_Err:
  SetErr "GetBackcolorPlane", Err.Number, Err.Description
  Resume GetBackcolorPlane_Exit
End Function

Public Function GetAttributesPlane( _
  ByRef paiRetData() As Integer, _
  ByVal piStartRow As Integer, _
  ByVal piEndRow As Integer, _
  ByVal piStartCol As Integer, _
  ByVal piEndCol As Integer) As Boolean
  
  On Error GoTo GetAttributesPlane_Err
  ClearErr

  Dim i               As Integer
  Dim k               As Integer
  Dim iLine           As Integer
  Dim iCol            As Integer
  Dim iDstRowCount    As Integer
  Dim iLength         As Integer
  
  iDstRowCount = (piEndRow - piStartRow) + 1
  If piEndCol > miCols Then
    piEndCol = miCols
  End If
  iLength = (piEndCol - piStartCol) + 1
  
  ReDim paiRetData(1 To iDstRowCount, 1 To iLength) As Integer
  
  'extract and copy data
  iLine = 1 'result line index
  For i = piStartRow To piEndRow
    iCol = 1
    For k = piStartCol To piEndCol
      paiRetData(iLine, iCol) = maiAttribs(i, k)
      iCol = iCol + 1
    Next k
    iLine = iLine + 1
  Next i
  
  GetAttributesPlane = True
  
GetAttributesPlane_Exit:
  Exit Function

GetAttributesPlane_Err:
  SetErr "GetAttributesPlane", Err.Number, Err.Description
  Resume GetAttributesPlane_Exit
End Function

Public Function CreateFrom( _
  ByRef poSrcGrid As CConsoleGrid, _
  ByVal piStartRow As Integer, _
  ByVal piEndRow As Integer, _
  ByVal piStartCol As Integer, _
  ByVal piEndCol As Integer _
  ) As Boolean
  
  On Error GoTo CreateFrom_Err
  ClearErr

  Dim fOK         As Boolean
  
  'clear ourself without allocating arrays
  miRows = 0
  miCols = 0
  Erase masChars, malForeCol, mabForeCol, malBackCol, mabBackCol, maiAttribs
  mfBuffersCreated = False
  miSelCol = 0
  miSelRow = 0
  miSelZone = 0
  moVT100Cache.Clear
  
  'V02.00.01 - Abandoned for now : choose planes from which to copy
  fOK = poSrcGrid.GetCharactersPlane(masChars(), piStartRow, piEndRow, piStartCol, piEndCol)
  If fOK Then
    fOK = poSrcGrid.GetForecolorPlane(malForeCol(), mabForeCol(), piStartRow, piEndRow, piStartCol, piEndCol)
    If fOK Then
      fOK = poSrcGrid.GetBackcolorPlane(malBackCol(), mabBackCol(), piStartRow, piEndRow, piStartCol, piEndCol)
    End If
  End If
  If fOK Then
    fOK = poSrcGrid.GetAttributesPlane(maiAttribs(), piStartRow, piEndRow, piStartCol, piEndCol)
  End If
  
  If fOK Then
    miRows = piEndRow - piStartRow + 1
    miCols = piEndCol - piStartCol + 1
    mfBuffersCreated = True
    SetDirty False
  End If
  
  CreateFrom = fOK
  
CreateFrom_Exit:
  Exit Function

CreateFrom_Err:
  SetErr "CreateFrom", Err.Number, Err.Description
  Resume CreateFrom_Exit
End Function

Public Function Paste( _
  ByRef poSrcGrid As CConsoleGrid, _
  ByVal piDstRow As Integer, _
  ByVal piDstCol As Integer, _
  ByVal pfTransparent As Boolean _
  ) As Boolean

  On Error GoTo Paste_Err
  ClearErr
  
  If (piDstCol > miCols) Or (piDstRow > miRows) Then
    SetErr "Paste", 9, Error$(9)
    Exit Function
  End If
  
  Dim iRowsToCopy   As Integer
  Dim iColsToCopy   As Integer
  Dim iLine         As Integer
  Dim iCol          As Integer
  Dim iDstRow       As Integer
  Dim iDstCol       As Integer
  Dim fCopy         As Boolean
  
  'allocate temp arrays
  iRowsToCopy = poSrcGrid.Rows
  If (iRowsToCopy + piDstRow - 1) > miRows Then
    iRowsToCopy = iRowsToCopy - ((iRowsToCopy + piDstRow - 1) - miRows)
  End If
  
  iColsToCopy = poSrcGrid.Cols
  If (iColsToCopy + piDstCol - 1) > miCols Then
    iColsToCopy = iColsToCopy - ((iColsToCopy + piDstCol - 1) - miCols)
  End If
  
  For iLine = 1 To iRowsToCopy
    iDstRow = piDstRow + iLine - 1
    Mid$(masChars(iDstRow), piDstCol, iColsToCopy) = left$(poSrcGrid.TextAt(iLine, 1), iColsToCopy)
    For iCol = 1 To iColsToCopy
      iDstCol = piDstCol + iCol - 1
      If Not pfTransparent Then
        fCopy = True
      Else
        fCopy = False
        With poSrcGrid
          If (.CharAt(iLine, iCol) <> " ") Or _
             (.HasForeCol(iLine, iCol)) Or _
             (.HasBackCol(iLine, iCol)) Or _
             (.CharAttribs(iLine, iCol) > 0) Then
            fCopy = True
          End If
        End With
      End If
      If fCopy Then
        malForeCol(iDstRow, iDstCol) = poSrcGrid.CharForeCol(iLine, iCol)
        mabForeCol(iDstRow, iDstCol) = poSrcGrid.HasForeCol(iLine, iCol)
        malBackCol(iDstRow, iDstCol) = poSrcGrid.CharBackCol(iLine, iCol)
        mabBackCol(iDstRow, iDstCol) = poSrcGrid.HasBackCol(iLine, iCol)
        maiAttribs(iDstRow, iDstCol) = poSrcGrid.CharAttribs(iLine, iCol)
      End If
    Next iCol
    moVT100Cache.InvalidateLine iLine
  Next iLine
  
  SetDirty True
  Paste = True
  
Paste_Exit:
  Exit Function

Paste_Err:
  SetErr "Paste", Err.Number, Err.Description
  Resume Paste_Exit
End Function

'
' Search/Replace character
'

Public Property Get LastFindRow() As Integer
  LastFindRow = miLastFindRow
End Property

Public Property Get LastFindCol() As Integer
  LastFindCol = miLastFindCol
End Property

Public Function FindFirstChar(ByVal psChar As String, ByVal piStartRow As Integer, ByVal piStartCol As Integer) As Boolean
  If Len(psChar) <> 1 Then Exit Function
  
  Dim iRow      As Integer
  Dim iCol      As Integer
  Dim iFind     As Integer
  
  On Error GoTo FindFirstChar_Err
  ClearErr
  
  miLastFindRow = 0
  miLastFindCol = 0
  msLastFindChar = ""
  iCol = piStartCol
  
  For iRow = piStartRow To miRows
    If iRow > piStartRow Then
      iFind = InStr(1, masChars(iRow), psChar, vbBinaryCompare)
      If iFind > 0 Then
        miLastFindCol = iFind
        miLastFindRow = iRow
        FindFirstChar = True
        msLastFindChar = psChar
        Exit Function
      End If
    Else
      iFind = InStr(iCol, masChars(iRow), psChar, vbBinaryCompare)
      If iFind > 0 Then
        miLastFindCol = iFind
        miLastFindRow = iRow
        FindFirstChar = True
        msLastFindChar = psChar
        Exit Function
      End If
    End If
  Next iRow
  
FindFirstChar_Exit:
  Exit Function

FindFirstChar_Err:
  SetErr "FindFirstChar", Err.Number, Err.Description
  Resume FindFirstChar_Exit
End Function

Public Function FindNextChar() As Boolean
  Dim iCol    As Integer
  Dim iRow    As Integer
  
  If miLastFindRow > 0 Then
    iCol = miLastFindCol
    iRow = miLastFindRow
    If iCol < miCols Then
      iCol = iCol + 1
    Else
      iCol = 1
      If iRow < miRows Then
        iRow = iRow + 1
      Else
        'don't "cycle" for the next one
        Exit Function
      End If
    End If
    FindNextChar = FindFirstChar(msLastFindChar, iRow, iCol)
  Else
    If Len(msLastFindChar) > 0 Then
      FindNextChar = FindFirstChar(msLastFindChar, 1, 1)
    End If
  End If
End Function

Public Function ReplaceChar( _
  ByVal psSearchChar As String, _
  ByVal psReplaceWithChar As String, _
  ByVal piRefRow As Integer, _
  ByVal piRefCol As Integer, _
  ByVal peReplaceScope As eReplaceCharScope, _
  ByRef plRetReplaceCt As Long) As Boolean
  
  Dim fOK       As Boolean
  
  On Error GoTo ReplaceChar_Err
  ClearErr
  
  plRetReplaceCt = 0
  If Len(psSearchChar) = 0 Then
    ReplaceChar = True
    Exit Function
  End If
  
  If psReplaceWithChar = "" Then psReplaceWithChar = " "
  
  If peReplaceScope = eFullLine Then
    piRefCol = 1
  ElseIf peReplaceScope = eFullText Then
    piRefCol = 1
    piRefRow = 1
  End If

  fOK = FindFirstChar(psSearchChar, piRefRow, piRefCol)
  Do While fOK
  
    fOK = False
    
    Select Case peReplaceScope
    Case eReplaceCharScope.eToEndOfLine, eReplaceCharScope.eFullLine
      If miLastFindRow = piRefRow Then
        Mid$(masChars(piRefRow), miLastFindCol, 1) = psReplaceWithChar
        plRetReplaceCt = plRetReplaceCt + 1
        moVT100Cache.InvalidateLine piRefRow
      Else
        Exit Do
      End If
      piRefCol = piRefCol + 1
      fOK = CBool(piRefCol <= miCols)
      
    Case eReplaceCharScope.eToEndOfText, eReplaceCharScope.eFullText
      Mid$(masChars(miLastFindRow), miLastFindCol, 1) = psReplaceWithChar
      moVT100Cache.InvalidateLine miLastFindRow
      plRetReplaceCt = plRetReplaceCt + 1
      piRefRow = miLastFindRow
      piRefCol = miLastFindCol + 1
      If piRefCol > miCols Then
        piRefRow = piRefRow + 1
        If piRefRow <= miRows Then
          piRefCol = 1
          fOK = True
        End If
      Else
        fOK = True
      End If
    End Select
    
    If fOK Then
      fOK = FindFirstChar(psSearchChar, piRefRow, piRefCol)
    End If
    
  Loop
  
  ReplaceChar = True
ReplaceChar_Exit:
  Exit Function

ReplaceChar_Err:
  SetErr "ReplaceChar", Err.Number, Err.Description
  Resume ReplaceChar_Exit
End Function

Public Function FindFirstColor( _
    ByVal plColor As Long, _
    ByVal peColorScope As eReplaceColorScope, _
    ByVal piStartRow As Integer, _
    ByVal piStartCol As Integer, _
    ByRef piFoundRow As Integer, _
    ByRef piFoundCol As Integer _
  ) As Boolean
  Dim iRow      As Integer
  Dim iCol      As Integer
  Dim fFound    As Boolean
  Dim lColComp  As Long
  Dim R As Integer, G As Integer, B As Integer
  
  On Error GoTo FindFirstColor_Err
  ClearErr
  
  piFoundRow = 0
  piFoundCol = 0
  If peColorScope = eBoth Then
    'can't search on both
    Exit Function
  End If
  
  For iRow = piStartRow To miRows
    For iCol = piStartCol To miCols
      Select Case peColorScope
      Case eReplaceColorScope.eForeColor
        lColComp = malForeCol(iRow, iCol)
        R = GetRed(lColComp)
        G = GetGreen(lColComp)
        B = GetBlue(lColComp)
        If RGB(R, G, B) = plColor Then
          piFoundRow = iRow
          piFoundCol = iCol
          fFound = True
          GoTo FindFirstColor_Exit
        End If
      Case eReplaceColorScope.eBackColor
        lColComp = malBackCol(iRow, iCol)
        R = GetRed(lColComp)
        G = GetGreen(lColComp)
        B = GetBlue(lColComp)
        If RGB(R, G, B) = plColor Then
          piFoundRow = iRow
          piFoundCol = iCol
          fFound = True
          GoTo FindFirstColor_Exit
        End If
      End Select
    Next iCol
    'continue at start col of next row
    piStartCol = 1
  Next iRow
  
FindFirstColor_Exit:
  FindFirstColor = fFound
  Exit Function

FindFirstColor_Err:
  SetErr "FindFirstColor", Err.Number, Err.Description
  Resume FindFirstColor_Exit
End Function

Public Function ReplaceColors( _
  ByVal plFindColor As Long, _
  ByVal plReplaceColor As Long, _
  ByVal piRefRow As Integer, _
  ByVal piRefCol As Integer, _
  ByVal piRefEndRow As Integer, _
  ByVal piRefEndCol As Integer, _
  ByVal peColorScope As eReplaceColorScope, _
  ByRef plRetReplaceCt As Long) As Boolean
  
  Dim fOK         As Boolean
  Dim iFoundRow   As Integer
  Dim iFoundCol   As Integer
  Dim eLoopScope  As eReplaceColorScope
  Dim lReplCt     As Long
  Dim fDoIt       As Boolean
  
  On Error GoTo ReplaceColors_Err
  ClearErr
  
  plRetReplaceCt = 0
  
  Select Case peColorScope
  Case eReplaceColorScope.eForeColor, eReplaceColorScope.eBoth
    eLoopScope = eForeColor
  Case eReplaceColorScope.eBackColor
    eLoopScope = eBackColor
  Case Else
    'shouldn't happen
    ReplaceColors = True
    Exit Function
  End Select
  
  fOK = True
  Do While fOK
    lReplCt = 0&
    Do
      fOK = FindFirstColor(plFindColor, eLoopScope, piRefRow, piRefCol, iFoundRow, iFoundCol)
      fDoIt = fOK
      If piRefEndRow > 0 Then
        If iFoundRow > piRefEndRow Then
          fOK = False
        End If
      End If
      If piRefEndCol > 0 Then
        If iFoundCol > piRefEndCol Then
          fDoIt = False
        End If
      End If
      If fOK And fDoIt Then
        If eLoopScope = eForeColor Then
          malForeCol(iFoundRow, iFoundCol) = plReplaceColor
        Else
          malBackCol(iFoundRow, iFoundCol) = plReplaceColor
        End If
        lReplCt = lReplCt + 1
      End If
    Loop Until Not fOK
    
    plRetReplaceCt = plRetReplaceCt + lReplCt
    
    Select Case peColorScope
    Case eReplaceColorScope.eForeColor, eReplaceColorScope.eBackColor
      Exit Do
    Case eReplaceColorScope.eBoth
      eLoopScope = eBackColor
    Case Else
      'shouldn't happen
      Exit Do
    End Select
  Loop
  
  If plRetReplaceCt > 0& Then
    SetDirty True
  End If
  ReplaceColors = True
ReplaceColors_Exit:
  Exit Function

ReplaceColors_Err:
  SetErr "ReplaceColors", Err.Number, Err.Description
  Resume ReplaceColors_Exit
End Function

Public Function SwitchColors( _
  ByVal piRefRow As Integer, _
  ByVal piRefCol As Integer, _
  ByVal piRefEndRow As Integer, _
  ByVal piRefEndCol As Integer, _
  ByRef pconCanvas As CConsoul, _
  ByRef plRetSwitchCt As Long) As Boolean
  
  Dim fOK         As Boolean
  Dim iRow        As Integer
  Dim iCol        As Integer
  Dim lTempColor  As Long
  Dim lSwitchCt   As Long
  
  On Error GoTo SwitchColors_Err
  ClearErr
  
  plRetSwitchCt = 0
  
  For iRow = piRefRow To piRefEndRow
    For iCol = piRefCol To piRefEndCol
      If mabBackCol(iRow, iCol) And mabForeCol(iRow, iCol) Then
        lTempColor = malForeCol(iRow, iCol)
        malForeCol(iRow, iCol) = malBackCol(iRow, iCol)
        malBackCol(iRow, iCol) = lTempColor
        lSwitchCt = lSwitchCt + 1&
      End If
    Next iCol
  Next iRow
  
  plRetSwitchCt = lSwitchCt
  If lSwitchCt > 0& Then
    SetDirty True
  End If
  
  SwitchColors = True
SwitchColors_Exit:
  Exit Function

SwitchColors_Err:
  SetErr "SwitchColors", Err.Number, Err.Description
  Resume SwitchColors_Exit
End Function


