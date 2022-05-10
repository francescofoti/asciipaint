Attribute VB_Name = "MWin32File"
Option Compare Database
Option Explicit

'Public constants
Public Const FILE_BEGIN   As Long = 0&
Public Const FILE_CURRENT As Long = 1&
Public Const FILE_END     As Long = 2&

'Private types
Private Type OVERLAPPED
  Internal      As Long
  InternalHigh  As Long
  Offset        As Long
  OffsetHigh    As Long
  hEvent        As Long
End Type

#If Win64 Then
Private Type SECURITY_ATTRIBUTES
 nLength As Long
 lpSecurityDescriptor As LongPtr
 bInheritHandle As Long
End Type
#Else
Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type
#End If

'Public constants
Public Const HFILE_ERROR = -1&
Public Const GENERIC_READ& = &H80000000
Public Const GENERIC_WRITE& = &H40000000
Public Const FILE_SHARE_READ& = &H1&
Public Const FILE_SHARE_WRITE& = &H2&
Public Const CREATE_ALWAYS& = 2&
Public Const CREATE_NEW& = 1&
Public Const OPEN_ALWAYS& = 4&
Public Const OPEN_EXISTING& = 3&
Public Const TRUNCATE_EXISTING& = 5&
Public Const FILE_ATTRIBUTE_ARCHIVE& = &H20
Public Const FILE_ATTRIBUTE_COMPRESSED& = &H800
Public Const FILE_ATTRIBUTE_NORMAL& = &H80
Public Const FILE_ATTRIBUTE_HIDDEN& = &H2
Public Const FILE_ATTRIBUTE_READONLY& = &H1
Public Const FILE_ATTRIBUTE_SYSTEM& = &H4
Public Const FILE_FLAG_WRITE_THROUGH& = &H80000000
Public Const FILE_FLAG_OVERLAPPED& = &H40000000
Public Const FILE_FLAG_NO_BUFFERING& = &H20000000
Public Const FILE_FLAG_RANDOM_ACCESS& = &H10000000
Public Const FILE_FLAG_SEQUENTIAL_SCAN& = &H8000000
Public Const FILE_FLAG_DELETE_ON_CLOSE& = &H4000000
Public Const INVALID_HANDLE_VALUE& = -1&
Public Const ERROR_ALREADY_EXISTS& = 183&
'/**/ Adapt functions for 64 bits /**/
#If Win64 Then
Private Declare PtrSafe Function apiCreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
Private Declare PtrSafe Function apiCreateFileWithSecurity Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
Private Declare PtrSafe Function apiCloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function apiSetEndOfFile Lib "kernel32" Alias "SetEndOfFile" (ByVal hFile As LongPtr) As Long
Private Declare PtrSafe Function apiLockFile Lib "kernel32" Alias "LockFile" (ByVal hFile As LongPtr, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long) As Long
Private Declare PtrSafe Function apiUnlockFile Lib "kernel32" Alias "UnlockFile" (ByVal hFile As LongPtr, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long) As Long

Private Declare PtrSafe Function apiWriteFile Lib "kernel32" Alias "WriteFile" (ByVal hFile As LongPtr, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As LongPtr, lpOverlapped As LongPtr) As Long
'Private Declare PtrSafe Function apiWriteFile Lib "kernel32" Alias "WriteFile" (ByVal hFile As LongPtr, lpBuffer As LongPtr, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As LongPtr, lpOverlapped As LongPtr) As LongPtr

Private Declare PtrSafe Function apiWriteFileString Lib "kernel32" Alias "WriteFile" (ByVal hFile As LongPtr, ByVal lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As LongPtr, lpOverlapped As LongPtr) As Long
Private Declare PtrSafe Function apiReadFile Lib "kernel32" Alias "ReadFile" (ByVal hFile As LongPtr, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As LongPtr) As Long
Private Declare PtrSafe Function apiReadFileString Lib "kernel32" Alias "ReadFile" (ByVal hFile As LongPtr, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As LongPtr) As Long
Private Declare PtrSafe Function apiSetFilePointer Lib "kernel32" Alias "SetFilePointer" (ByVal hFile As LongPtr, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare PtrSafe Function apiGetFileSize Lib "kernel32" Alias "GetFileSize" (ByVal hFile As LongPtr, lpFileSizeHigh As Long) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Sub CopyMemoryToString Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpstrDest As String, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Sub CopyMemoryFromString Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, ByVal lpstrSource As String, ByVal Length As LongPtr)
#Else
Private Declare Function apiCreateFile& Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long)
Private Declare Function apiCreateFileWithSecurity& Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long)
Private Declare Function apiCloseHandle& Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long)
Private Declare Function apiSetEndOfFile& Lib "kernel32" Alias "SetEndOfFile" (ByVal hObject As Long)
Private Declare Function apiLockFile& Lib "kernel32" Alias "LockFile" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long)
Private Declare Function apiUnlockFile& Lib "kernel32" Alias "UnlockFile" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long)
Private Declare Function apiReadFile& Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long)
Private Declare Function apiReadFileString& Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, ByVal lpstrRetBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long)
Private Declare Function apiWriteFile& Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long)
Private Declare Function apiWriteFileBytes& Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long)
Private Declare Function apiWriteFileString& Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, ByVal lpszBuffer As String, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long)
Private Declare Function apiSetFilePointer& Lib "kernel32" Alias "SetFilePointer" (ByVal hFile As Long, ByVal lDistanceToMoveLow As Long, ByVal lDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long)
Private Declare Function apiGetFileSize Lib "kernel32" Alias "GetFileSize" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub CopyMemoryToString Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpstrDest As String, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub CopyMemoryFromString Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, ByVal lpstrSource As String, ByVal cbCopy As Long)
#End If

'Opens a file in shared and random access mode.
#If Win64 Then
Public Function Win32OpenFile(ByVal sFilename As String, ByVal fReadOnly As Boolean, ByVal fExclusive As Boolean, ByRef lRetErr As Long, ByRef sRetErrDesc As String, Optional ByVal fNoCache As Boolean = False, Optional ByVal fDeleteOnClose As Boolean = False) As LongPtr
  Dim lWinFileHandle    As LongPtr
  Dim lTemplateHandle   As LongPtr
#Else
Public Function Win32OpenFile(ByVal sFilename As String, ByVal fReadOnly As Boolean, ByVal fExclusive As Boolean, ByRef lRetErr As Long, ByRef sRetErrDesc As String, Optional ByVal fNoCache As Boolean = False, Optional ByVal fDeleteOnClose As Boolean = False) As Long
  Dim lWinFileHandle    As Long
  Dim lTemplateHandle   As Long
#End If
  Dim lAccessFlags      As Long
  Dim lShareMode        As Long
  Dim lDisposition      As Long
  Dim lAttributes       As Long
  Dim tSecurity         As SECURITY_ATTRIBUTES
  
  'Setup access flags
  lAccessFlags = GENERIC_READ
  If Not fReadOnly Then
    lAccessFlags = lAccessFlags Or GENERIC_WRITE
  End If
  
  If Not fExclusive Then
    lShareMode = FILE_SHARE_READ
    If Not fReadOnly Then
      lShareMode = lShareMode Or FILE_SHARE_WRITE
    End If
  End If
  lDisposition = OPEN_EXISTING
  If Not fNoCache Then
    lAttributes = FILE_FLAG_RANDOM_ACCESS
  Else
    lAttributes = FILE_FLAG_RANDOM_ACCESS Or FILE_FLAG_WRITE_THROUGH&
  End If
  If fDeleteOnClose Then
    lAttributes = lAttributes Or FILE_FLAG_DELETE_ON_CLOSE
  End If
  With tSecurity
    .nLength = Len(tSecurity)
    .lpSecurityDescriptor = 0
    .bInheritHandle = True 'Doesn't really matter
  End With
  lWinFileHandle = apiCreateFileWithSecurity(sFilename, lAccessFlags, lShareMode, tSecurity, lDisposition, lAttributes, lTemplateHandle)
  
  'test for errors
  If lWinFileHandle = INVALID_HANDLE_VALUE Then
    lRetErr = Err.LastDllError: sRetErrDesc = LastApiError()
  End If
  
  Win32OpenFile = lWinFileHandle
End Function

'Opens a file in shared and random access mode.
#If Win64 Then
Public Function Win32OpenFileRaw(ByVal sFilename As String, ByVal plAccessFlags As Long, ByVal plShareMode As Long, ByVal plDisposition As Long, ByVal plAttributes As Long, ByVal plTplHandle As LongLong, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As LongPtr
  Dim lWinFileHandle    As LongLong
  Dim lTemplateHandle   As LongLong
#Else
Public Function Win32OpenFileRaw(ByVal sFilename As String, ByVal plAccessFlags As Long, ByVal plShareMode As Long, ByVal plDisposition As Long, ByVal plAttributes As Long, ByVal plTplHandle As Long, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Long
  Dim lWinFileHandle    As Long
  Dim lTemplateHandle   As Long
#End If
  Dim tSecurity         As SECURITY_ATTRIBUTES
  
  With tSecurity
    .nLength = Len(tSecurity)
    .lpSecurityDescriptor = 0
    .bInheritHandle = True 'Doesn't really matter
  End With
  lWinFileHandle = apiCreateFileWithSecurity(sFilename, plAccessFlags, plShareMode, tSecurity, plDisposition, plAttributes, plTplHandle)
  
  'test for errors
  If lWinFileHandle = INVALID_HANDLE_VALUE Then
    lRetErr = Err.LastDllError: sRetErrDesc = LastApiError()
  End If
  
  Win32OpenFileRaw = lWinFileHandle
End Function

'Close a file previously opened by Win32OpenFile.
'Warning: All file locks must be previously released.
#If Win64 Then
Public Function Win32CloseFile(ByVal lWinFileHandle As LongPtr, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Boolean
#Else
Public Function Win32CloseFile(ByVal lWinFileHandle As Long, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Boolean
#End If
  If apiCloseHandle(lWinFileHandle) Then
    Win32CloseFile = True
  Else
    lRetErr = Err.LastDllError: sRetErrDesc = LastApiError()
  End If
End Function

'Set the end of file at the current file position.
#If Win64 Then
Public Function Win32SetEndOfFile(ByVal lWinFileHandle As LongPtr, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Boolean
#Else
Public Function Win32SetEndOfFile(ByVal lWinFileHandle As Long, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Boolean
#End If
  If apiSetEndOfFile(lWinFileHandle) Then
    Win32SetEndOfFile = True
  Else
    lRetErr = Err.LastDllError: sRetErrDesc = LastApiError()
  End If
End Function

#If Win64 Then
Public Function Win32GetFileSize(ByVal lWinFileHandle As LongPtr) As Long
#Else
Public Function Win32GetFileSize(ByVal lWinFileHandle As Long) As Long
#End If
  Dim lRetSize    As Long
  Dim lDummy      As Long
  lRetSize = apiGetFileSize(lWinFileHandle, lDummy)
  Win32GetFileSize = lRetSize
End Function

#If Win64 Then
'Locks a region of a previously opened file.
Public Function Win32LockFileRegion(ByVal lWinFileHandle As LongPtr, ByVal lOffset As Long, ByVal lNumberOfBytesToLock As Long, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Boolean
#Else
'Locks a region of a previously opened file.
Public Function Win32LockFileRegion(ByVal lWinFileHandle As Long, ByVal lOffset As Long, ByVal lNumberOfBytesToLock As Long, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Boolean
#End If
  If apiLockFile(lWinFileHandle, lOffset, 0&, lNumberOfBytesToLock, 0&) Then
    Win32LockFileRegion = True
  Else
    lRetErr = Err.LastDllError: sRetErrDesc = LastApiError()
  End If
End Function

'Read a byte block from an opened file. The size of the byte array is used
'to compute the number of bytes to read from the file and thus should be
'previously allocated.
#If Win64 Then
Public Function Win32ReadBytes(ByVal lWinFileHandle As LongPtr, abRetBuffer() As Byte, ByRef lRetNumberOfBytesRead As Long, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Boolean
#Else
Public Function Win32ReadBytes(ByVal lWinFileHandle As Long, abRetBuffer() As Byte, ByRef lRetNumberOfBytesRead As Long, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Boolean
#End If
  Dim lNumberOfBytesToRead As Long
  lNumberOfBytesToRead = (UBound(abRetBuffer) - LBound(abRetBuffer)) + 1&
  If apiReadFile(lWinFileHandle, abRetBuffer(LBound(abRetBuffer)), lNumberOfBytesToRead, lRetNumberOfBytesRead, 0&) Then
    Win32ReadBytes = True
  Else
    lRetErr = Err.LastDllError: sRetErrDesc = LastApiError()
  End If
End Function

'Read a byte block from an opened file. The size of the string is used
'to compute the number of bytes to read from the file and thus should be
'previously filled, for example with vb's String$().
#If Win64 Then
Public Function Win32ReadBytesString(ByVal lWinFileHandle As LongPtr, ByRef sRetBuffer As String, ByRef lRetNumberOfBytesRead As Long, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Boolean
#Else
Public Function Win32ReadBytesString(ByVal lWinFileHandle As Long, ByRef sRetBuffer As String, ByRef lRetNumberOfBytesRead As Long, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Boolean
#End If
  Dim lNumberOfBytesToRead As Long
  lNumberOfBytesToRead = Len(sRetBuffer)
  If lNumberOfBytesToRead = 0 Then
    Win32ReadBytesString = True
    Exit Function
  End If
  
  If apiReadFileString(lWinFileHandle, sRetBuffer, lNumberOfBytesToRead, lRetNumberOfBytesRead, 0&) Then
    Win32ReadBytesString = True
  Else
    lRetErr = Err.LastDllError: sRetErrDesc = LastApiError()
  End If
End Function

'Read a byte block from an opened file. The data is read into the buffer
'pointed by lpDataBuffer. Memory pointed by lpDataBuffer should have been previously allocated.
#If Win64 Then
Public Function Win32ReadData(ByVal lWinFileHandle As LongPtr, lpDataBuffer As Long, ByVal lNumberOfBytesToRead As Long, ByRef lRetNumberOfBytesRead As Long, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Boolean
#Else
Public Function Win32ReadData(ByVal lWinFileHandle As Long, lpDataBuffer As Long, ByVal lNumberOfBytesToRead As Long, ByRef lRetNumberOfBytesRead As Long, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Boolean
#End If
  If apiReadFile(lWinFileHandle, lpDataBuffer, lNumberOfBytesToRead, lRetNumberOfBytesRead, 0&) Then
    Win32ReadData = True
  Else
    lRetErr = Err.LastDllError: sRetErrDesc = LastApiError()
  End If
End Function

'Unlocks a previously locked region.
'Caution: when unlocking, specify the SAME number of bytes that were locked, otherwise
'locked leaks may be created.
#If Win64 Then
Public Function Win32UnLockFileRegion(ByVal lWinFileHandle As LongPtr, ByVal lOffset As Long, ByVal lNumberOfBytesToUnLock As Long, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Boolean
#Else
Public Function Win32UnLockFileRegion(ByVal lWinFileHandle As Long, ByVal lOffset As Long, ByVal lNumberOfBytesToUnLock As Long, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Boolean
#End If
  If apiUnlockFile(lWinFileHandle, lOffset, 0&, lNumberOfBytesToUnLock, 0&) Then
    Win32UnLockFileRegion = True
  Else
    lRetErr = Err.LastDllError: sRetErrDesc = LastApiError()
  End If
End Function

'Write a byte block to an opened file. The size of the byte array is used
'to compute the number of bytes to write to the file and thus should be
'previously allocated and filled.
#If Win64 Then

Public Function Win32WriteBytes(ByVal lWinFileHandle As LongPtr, abBuffer() As Byte, ByRef lRetNumberOfBytesWritten As LongPtr, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Boolean
#Else
Public Function Win32WriteBytes(ByVal lWinFileHandle As Long, abBuffer() As Byte, ByRef lRetNumberOfBytesWritten As Long, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Boolean
#End If
  Dim lNumberOfBytesToWrite As Long
  lNumberOfBytesToWrite = (UBound(abBuffer) - LBound(abBuffer)) + 1&
  If apiWriteFileBytes(lWinFileHandle, abBuffer(LBound(abBuffer)), lNumberOfBytesToWrite, lRetNumberOfBytesWritten, 0&) Then
    Win32WriteBytes = True
  Else
    lRetErr = Err.LastDllError: sRetErrDesc = LastApiError()
  End If
End Function

#If Win64 Then
Public Function Win32WriteBytesString(ByVal lWinFileHandle As LongPtr, ByRef sData As String, ByRef lRetNumberOfBytesWritten As LongPtr, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Boolean
#Else
Public Function Win32WriteBytesString(ByVal lWinFileHandle As Long, ByRef sData As String, ByRef lRetNumberOfBytesWritten As Long, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Boolean
#End If
  Dim lNumberOfBytesToWrite As Long
  Dim abData()              As Byte
  lNumberOfBytesToWrite = Len(sData)
  If apiWriteFileString(lWinFileHandle, sData, lNumberOfBytesToWrite, lRetNumberOfBytesWritten, 0&) Then
    Win32WriteBytesString = True
  Else
    lRetErr = Err.LastDllError: sRetErrDesc = LastApiError()
  End If
End Function

'Write a byte block to an opened file. The size of the byte array is used
'to compute the number of bytes to write to the file and thus should be
'previously allocated and filled.
#If Win64 Then
Public Function Win32WriteData(ByVal lWinFileHandle As LongPtr, lpDataBuffer As Long, ByVal lNumberOfBytesToWrite As Long, ByRef lRetNumberOfBytesWritten As LongPtr, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Boolean
#Else
Public Function Win32WriteData(ByVal lWinFileHandle As Long, lpDataBuffer As Long, ByVal lNumberOfBytesToWrite As Long, ByRef lRetNumberOfBytesWritten As Long, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Boolean
#End If
  If apiWriteFile(lWinFileHandle, lpDataBuffer, lNumberOfBytesToWrite, lRetNumberOfBytesWritten, 0&) Then
    Win32WriteData = True
  Else
    lRetErr = Err.LastDllError: sRetErrDesc = LastApiError()
  End If
End Function

#If Win64 Then
Public Function Win32SeekFile(ByVal lWinFileHandle As LongPtr, ByVal lDistance As Long, ByVal lMoveMethod As Long, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Boolean
#Else
'Move the file pointer (Seek in the file)
Public Function Win32SeekFile(ByVal lWinFileHandle As Long, ByVal lDistance As Long, ByVal lMoveMethod As Long, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Boolean
#End If
  Dim lRet      As Long
  lRet = apiSetFilePointer(lWinFileHandle, lDistance, 0&, lMoveMethod)
  If lRet = HFILE_ERROR Then
    lRetErr = Err.LastDllError: sRetErrDesc = LastApiError()
  Else
    Win32SeekFile = True
  End If
End Function

'get the file pointer position
#If Win64 Then
Public Function Win32GetFilePosition(ByVal lWinFileHandle As LongPtr, ByRef lRetPos As Long, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Boolean
#Else
Public Function Win32GetFilePosition(ByVal lWinFileHandle As Long, ByRef lRetPos As Long, ByRef lRetErr As Long, ByRef sRetErrDesc As String) As Boolean
#End If
  lRetPos = apiSetFilePointer(lWinFileHandle, 0&, 0&, FILE_CURRENT)
  If lRetPos = HFILE_ERROR Then
    lRetErr = Err.LastDllError: sRetErrDesc = LastApiError()
  Else
    Win32GetFilePosition = True
  End If
End Function


