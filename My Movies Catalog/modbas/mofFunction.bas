Attribute VB_Name = "modFunction"
Option Explicit

'/// Init Common Controls
'/// *****************************************************************
Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public m_hMod As Long

'/// Verify if the File exist
'/// *****************************************************************
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Const GENERIC_READ As Long = &H80000000
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const OPEN_EXISTING As Long = 3
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const MAX_PATH = 260

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

'/// Get File size
'/// *****************************************************************
Private Const K_B = 1024#
Private Const M_B = (K_B * 1024#) ' MegaBytes
Private Const G_B = (M_B * 1024#) ' GigaBytes
Private Const T_B = (G_B * 1024#) ' TeraBytes
Private Const P_B = (T_B * 1024#) ' PetaBytes
Private Const E_B = (P_B * 1024#) ' ExaBytes
Private Const Z_B = (E_B * 1024#) ' ZettaBytes
Private Const Y_B = (Z_B * 1024#) ' YottaBytes

Public Enum DISP_BYTES_FORMAT
    DISP_BYTES_LONG
    DISP_BYTES_SHORT
    DISP_BYTES_ALL
End Enum

'/// Open File/Document or Browser Web
'/// *****************************************************************
Public Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)

Public Enum WINDOWSSTATE
    OPEN_HIDDEN = 0
    OPEN_NORMAL = 4
    OPEN_MINIMIZED = 2
    OPEN_MAXIMIZED = 3
End Enum
Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   m_hMod = LoadLibrary("shell32.dll")
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function

Public Function FileExists(sSource As String) As Boolean
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   hFile = FindFirstFile(sSource, WFD)
   FileExists = hFile <> INVALID_HANDLE_VALUE
   Call FindClose(hFile)
End Function

Public Function MakeProper(StringIn As Variant) As String
    On Error GoTo HandleErr
    Dim strBuild As String
    Dim intLength As Integer
    Dim intCounter As Integer
    Dim strChar As String
    Dim strPrevChar As String
    intLength = Len(StringIn)
    If intLength > 0 Then
        strBuild = UCase(Left(StringIn, 1))
        For intCounter = 1 To intLength
            strPrevChar = Mid$(StringIn, intCounter, 1)
            strChar = Mid$(StringIn, intCounter + 1, 1)
            Select Case strPrevChar
                Case Is = " ", ".", "/"
                    strChar = UCase(strChar)
                Case Else
            End Select
            strBuild = strBuild & strChar
        Next intCounter
        MakeProper = strBuild
        strBuild = MakeWordsLowerCase(strBuild, " and ", " or ", " the ", " a ", " To ")
        MakeProper = strBuild
    End If
ExitHere:
    Exit Function
HandleErr:
    Err.Clear
        Resume ExitHere
End Function

Private Function MakeWordsLowerCase(StringIn As String, ParamArray WordsToCheck()) As String
    'Looks for the words in the WordsToCheck
    '     Array within
    'the StringIn string and makes them lower case
    
    On Error GoTo HandleErr
    Dim strWordToFind As String
    Dim intWordStarts As Integer
    Dim intWordEnds As Integer
    Dim intStartLooking As Integer
    Dim strResult As String
    Dim intLength As Integer
    Dim intCounter As Integer
    strResult = StringIn
    intLength = Len(strResult)
    intStartLooking = 1
    For intCounter = LBound(WordsToCheck) To UBound(WordsToCheck)
        strWordToFind = WordsToCheck(intCounter)
        Do
        intWordStarts = InStr(intStartLooking, strResult, strWordToFind)
        If intWordStarts = 0 Then Exit Do
        intWordEnds = intWordStarts + Len(strWordToFind)
        strResult = Left(strResult, intWordStarts - 1) & LCase(strWordToFind) & Mid$(strResult, intWordEnds, (intLength - intWordEnds) + 1)
        intStartLooking = intWordEnds
        Loop While intWordStarts > 0
        intStartLooking = 1
    Next intCounter
    MakeWordsLowerCase = strResult
ExitHere:
    Exit Function
HandleErr:
    Err.Clear
        Resume ExitHere
End Function
Public Function MakeDirectory(szDirectory As String) As Boolean
Dim strFolder As String
Dim szRslt As String
On Error GoTo IllegalFolderName
If Right$(szDirectory, 1) <> "\" Then szDirectory = szDirectory & "\"
strFolder = szDirectory
szRslt = Dir(strFolder, 63)
While szRslt = ""
    DoEvents
    szRslt = Dir(strFolder, 63)
    strFolder = Left$(strFolder, Len(strFolder) - 1)
    If strFolder = "" Then GoTo IllegalFolderName
Wend
If Right$(strFolder, 1) <> "\" Then strFolder = strFolder & "\"
While strFolder <> szDirectory
    strFolder = Left$(szDirectory, Len(strFolder) + 1)
    If Right$(strFolder, 1) = "\" Then MkDir strFolder
Wend
MakeDirectory = True
Exit Function
IllegalFolderName:
        MakeDirectory = False
    Err.Clear
End Function

Public Function GetSizeBytes(Dec As Variant, Optional DispBytesFormat As DISP_BYTES_FORMAT = DISP_BYTES_ALL) As String
    Dim DispLong As String: Dim DispShort As String: Dim s As String
    If DispBytesFormat <> DISP_BYTES_SHORT Then DispLong = FormatNumber(Dec, 0) & " bytes" Else DispLong = ""
    If DispBytesFormat <> DISP_BYTES_LONG Then
        If Dec > Y_B Then
            DispShort = FormatNumber(Dec / Y_B, 2) & " Yb"
        ElseIf Dec > Z_B Then
            DispShort = FormatNumber(Dec / Z_B, 2) & " Zb"
        ElseIf Dec > E_B Then
            DispShort = FormatNumber(Dec / E_B, 2) & " Eb"
        ElseIf Dec > P_B Then
            DispShort = FormatNumber(Dec / P_B, 2) & " Pb"
        ElseIf Dec > T_B Then
            DispShort = FormatNumber(Dec / T_B, 2) & " Tb"
        ElseIf Dec > G_B Then
            DispShort = FormatNumber(Dec / G_B, 2) & " Gb"
        ElseIf Dec > M_B Then
            DispShort = FormatNumber(Dec / M_B, 2) & " Mb"
        ElseIf Dec > K_B Then
            DispShort = FormatNumber(Dec / K_B, 2) & " Kb"
        Else
            DispShort = FormatNumber(Dec, 0) & " bytes"
        End If
    Else
        DispShort = ""
    End If
    Select Case DispBytesFormat
        Case DISP_BYTES_SHORT:
            GetSizeBytes = DispShort
        Case DISP_BYTES_LONG:
            GetSizeBytes = DispLong
        Case Else:
            GetSizeBytes = DispLong & " (" & DispShort & ")"
    End Select
End Function

Public Function FormatFileSize(ByVal dblFileSize As Double, Optional ByVal strFormatMask As String) As String
On Error Resume Next
Select Case dblFileSize
    Case 0 To 1023 ' Bytes
        FormatFileSize = Format(dblFileSize) & " bytes"
    Case 1024 To 1048575 ' KB
        If strFormatMask = Empty Then strFormatMask = "###0"
        FormatFileSize = Format(dblFileSize / 1024#, strFormatMask) & " KB"
    Case 1024# ^ 2 To 1073741823 ' MB
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatFileSize = Format(dblFileSize / (1024# ^ 2), strFormatMask) & " MB"
    Case Is > 1073741823# ' GB
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatFileSize = Format(dblFileSize / (1024# ^ 3), strFormatMask) & " GB"
End Select
End Function

Public Function FormatTime(ByVal sglTime As Single) As String
On Error Resume Next
Select Case sglTime
    Case 0 To 59
        FormatTime = Format(sglTime, "0") & " sec"
    Case 60 To 3599
        FormatTime = Format(Int(sglTime / 60), "#0") & " min " & Format(sglTime Mod 60, "0") & " sec"
    Case Else
        FormatTime = Format(Int(sglTime / 3600), "#0") & " hr " & Format(sglTime / 60 Mod 60, "0") & " min"
End Select
End Function

Public Function FormatPercentage(ByVal dblFileSize As Double, Optional ByVal strFormatMask As String) As String
On Error Resume Next
Select Case dblFileSize
    Case 0 To 1023
        FormatPercentage = Format(dblFileSize)
    Case 1024 To 1048575
        If strFormatMask = Empty Then strFormatMask = "###0"
        FormatPercentage = Format(dblFileSize / 1024#, strFormatMask)
    Case 1024# ^ 2 To 1073741823
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatPercentage = Format(dblFileSize / (1024# ^ 2), strFormatMask)
    Case Is > 1073741823#
        If strFormatMask = Empty Then strFormatMask = "###0.0"
        FormatPercentage = Format(dblFileSize / (1024# ^ 3), strFormatMask)
End Select
End Function
