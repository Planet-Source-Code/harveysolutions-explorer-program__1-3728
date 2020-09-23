Attribute VB_Name = "EplorerMod1"
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function SearchPath Lib "kernel32" Alias "SearchPathA" (ByVal lpPath As String, ByVal lpFileName As String, ByVal lpExtension As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" _
       (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long


Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" _
       (ByVal lpRootPathName As String, _
       lpSectorsPerCluster As Long, _
       lpBytesPerSector As Long, _
       lpNumberOfFreeClusters As Long, _
       lpTotalNumberOfClusters As Long) As Long
   

Public NbFile As Long
Public FileFSToOpen As String
Public StringToFind As String
Public ProgressCancel As Boolean
Public TypeView

Public Const MAX_PATH As Long = 260
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
   
Type FileTime
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
   
Type SaveF
   StingToSave As String
End Type
   
Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FileTime
    ftLastAccessTime As FileTime
    ftLastWriteTime As FileTime
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type


   Public Type RANDYS_OWN_DRIVE_INFO
       DrvSectors As Long
       DrvBytesPerSector As Long
       DrvFreeClusters As Long
       DrvTotalClusters As Long
       DrvSpaceFree As Long
       DrvSpaceUsed As Long
       DrvSpaceTotal As Long
  End Type
        

Public Function StripNull(ByVal WhatStr As String) As String
   Dim pos As Integer
    pos = InStr(WhatStr, Chr$(0))
    If pos > 0 Then
       StripNull = Left$(WhatStr, pos - 1)
    Else
       StripNull = WhatStr
    End If
End Function


           
