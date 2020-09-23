Attribute VB_Name = "modAdmin"
Option Explicit

' any public declarations are used in the main form

' used to see if a function exists in a specific dll
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long

' used to get special folders: system, my documents
Private Declare Function SHGetFolderPath Lib "shfolder" _
        Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, _
        ByVal nFolder As Long, ByVal hToken As Long, _
        ByVal dwFlags As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "Shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
    Private Type SHITEMID
        cb   As Long
        abID As Byte
    End Type
    Private Type ITEMIDLIST
        mkid As SHITEMID
    End Type
    Public Enum eSpecialFolders
        CSIDL_PERSONAL = &H5
        CSIDL_SYSTEM = &H25
        CSIDL_DESKTOP = &H0
    End Enum

' used to perfectly measure the available width of the listview control
Public Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
    Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type

' APIs and constants used to manipulate the desktop listview
Public Declare Function IsWindowVisible Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Public Declare Function SendNotifyMessage Lib "user32.dll" Alias "SendNotifyMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function RegisterWindowMessage Lib "user32.dll" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Public Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Private Const LVM_FIRST As Long = &H1000
    Private Const LVN_FIRST As Long = -100
    Private Const NM_FIRST As Long = 0
    Public Const WM_NOTIFY As Long = &H4E  ' flag for custom DLL, must be wParam of SendMessage/PostMessage
    Public Const WM_ENDSESSION As Long = &H16  ' terminate subclassing/reset
    Public Const NM_CUSTOMDRAW As Long = (NM_FIRST - 12) ' item is being drawn; get colors
    Public Const LVN_GETDISPINFO As Long = (LVN_FIRST - 50) ' something changed on desktop
    Public Const LVN_DELETEITEM As Long = (LVN_FIRST - 3)   ' something deleted from desktop
    Public Const LVN_INSERTITEM As Long = (LVN_FIRST - 2)   ' something added to desktop
    Public Const LVN_BEGINDRAG As Long = (LVN_FIRST - 9)    ' one or more icons being dragged
    Public Const WM_GETICON As Long = &H7F  ' we want the icon image
    Public Const WM_LBUTTONDBLCLK As Long = &H203
    Public Type POINTAPI
        x As Long
        y As Long
    End Type

' APIs used only for the GUI interface or updating host listview
Private Declare Function AttachThreadInput Lib "user32.dll" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Public Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function CopyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Public Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Public Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Public Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetLastActivePopup Lib "user32.dll" (ByVal hwndOwnder As Long) As Long
Private Declare Function GetNextWindow Lib "user32.dll" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long 'also used in clsBarColors
Public Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Public Const GWL_STYLE As Long = -16       ' Get/Set window styles
    Public Const COLOR_DESKTOP As Long = 1&
    Public Const WM_SYSCOLORCHANGE As Long = &H15        ' changed desktop colors/theme
    Public Const WM_SETTINGSCHANGE As Long = &H1A          ' something changed; can occur when
    Public Const LVM_GETTEXTCOLOR As Long = (LVM_FIRST + 35)
    Public Const LVM_REDRAWITEMS As Long = (LVM_FIRST + 21)
    Public Const LVM_GETITEMCOUNT As Long = (LVM_FIRST + 4)
    Public Const LVM_GETITEMTEXTA As Long = (LVM_FIRST + 45)
    Public Const LVM_SETITEMTEXTA As Long = (LVM_FIRST + 46)
    Public Const LVM_GETITEMA As Long = (LVM_FIRST + 5)
    Public Const LVM_GETITEMPOSITION As Long = (LVM_FIRST + 16)
    Public Const LVM_SETITEMPOSITION32 As Long = (LVM_FIRST + 49)
    Public Const LVIF_TEXT As Long = &H1
    Public Const LVIF_IMAGE As Long = &H2
    Public Const WM_COMMAND As Long = &H111
    Public Const WM_KILLFOCUS As Long = &H8
    Public Const WM_SETFOCUS As Long = &H7
    Public Const WM_ACTIVATE As Long = &H6
    Public Const WM_CLOSE As Long = &H10
    Public Const WM_DISPLAYCHANGE As Long = &H7E
    Public Const WM_MENUSELECT As Long = &H11F
    Public Const WM_USER As Long = &H400
    Public Const WM_LBUTTONDOWN As Long = &H201
    Public Const WM_RBUTTONUP As Long = &H205
    Public Const WM_NCHITTEST As Long = &H84
    Public Const HTCLOSE As Long = 20
    Public Const SC_CLOSE As Long = &HF060&
    Public Const SC_MINIMIZE As Long = &HF020&
    Public Const SW_SHOW As Long = 5
    Public Const SW_SHOWNORMAL As Long = 1
    Public Const WM_SYSCOMMAND As Long = &H112
    Public Const HWND_BROADCAST As Long = &HFFFF&
    Private Const GW_HWNDPREV As Long = 3
    Private Const GW_HWNDNEXT As Long = 2
    Public Const LVS_AUTOARRANGE As Long = &H100
    Public Const IDM_TOGGLEAUTOARRANGE9x = &H7041
    Public Const IDM_TOGGLEAUTOARRANGEnt = IDM_TOGGLEAUTOARRANGE9x Or &H10

' Following used primarily to create interprocess communications
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" (ByVal hFile As Long, ByVal lpFileMappigAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As String) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As Long, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As Long, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long
    Private Type OSVERSIONINFO
       dwOSVersionInfoSize As Long
       dwMajorVersion As Long
       dwMinorVersion As Long
       dwBuildNumber As Long
       dwPlatformId As Long
       szCSDVersion As String * 128
    End Type
   Private Const PAGE_READWRITE = &H4
   Private Const MEM_RESERVE = &H2000&
   Private Const MEM_RELEASE = &H8000&
   Private Const MEM_COMMIT = &H1000&
   Private Const PROCESS_VM_OPERATION = &H8
   Private Const PROCESS_VM_READ = &H10
   Private Const PROCESS_VM_WRITE = &H20
   Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
   Private Const SECTION_QUERY = &H1
   Private Const SECTION_MAP_WRITE = &H2
   Private Const SECTION_MAP_READ = &H4
   Private Const SECTION_MAP_EXECUTE = &H8
   Private Const SECTION_EXTEND_SIZE = &H10
   Private Const SECTION_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SECTION_QUERY Or SECTION_MAP_WRITE Or SECTION_MAP_READ Or SECTION_MAP_EXECUTE Or SECTION_EXTEND_SIZE
   Private Const FILE_MAP_ALL_ACCESS = SECTION_ALL_ACCESS


' Used to read INI files
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFilename As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFilename As String) As Long
Private Declare Function apiGetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Const maxLen As Long = 256 ' file lengths
Private m_sysNT As Long            ' NT based or not
Private crcTable() As Long         'crc32 lookup

Public Function ReadWriteINI(ByVal iniFile As String, ByVal Mode As String, _
            ByVal SectionName As String, ByVal Keyname As String, _
            Optional ByVal KeyValue As String = "*****", _
            Optional ByVal DeleteSection As Boolean = False, _
            Optional ByVal ReadBufferLen As Long = 512) As String

' Inserted by LaVolpe OnError Insertion Program.
On Error GoTo ReadWriteINI_General_ErrTrap

If DeleteSection = True Then
    WritePrivateProfileSection SectionName, "", iniFile
    ReadWriteINI = "ok"
    Exit Function
End If

Dim anInt As Long
Dim defaultValue As String

' ******* WRITE MODE *************************************
If UCase(Mode) = "WRITE" Then
    If KeyValue = "" Then KeyValue = vbNullString
    anInt = WritePrivateProfileString(SectionName, Keyname, KeyValue, iniFile)
    If anInt <> 0 Then ReadWriteINI = "ok"
ElseIf UCase(Mode) = "GET" Then
' *******  READ MODE *************************************
    defaultValue = KeyValue
    KeyValue = String$(ReadBufferLen + 1, 32) ' account for trailing nullchar
    anInt = GetPrivateProfileString(SectionName, Keyname, defaultValue, KeyValue, ReadBufferLen + 1, iniFile)
    If anInt <> 0 Then
        ReadWriteINI = Left$(KeyValue, anInt)
    Else
        ReadWriteINI = ""
    End If
End If

' Inserted by LaVolpe OnError Insertion Program.
ReadWriteINI_General_ErrTrap:
If Err Then
    MsgBox "Err: " & Err.Number & " - Procedure: ReadWriteINI" & vbCrLf & Err.Description, vbExclamation + vbOKOnly
    Err.Clear
End If
End Function


Public Function GetUserName() As String
    ' Returns the network login name
    Dim lngRtn As Long, nrChars As Long
    Dim strUserName As String
    strUserName = String$(maxLen, 0)
    nrChars = maxLen
    lngRtn = apiGetUserName(strUserName, nrChars)
    If lngRtn <> 0 Then
        GetUserName = Left$(strUserName, nrChars - 1)
    Else
        GetUserName = "Default"
    End If
End Function

Public Function CRC32(ByVal sFileName As String, Optional ByVal lcrc As Long = 0) As Long
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=38270&lngWId=1
  'The 3rd var is optional so if you passed
  'part of the array you can continue it by passing
  'the return value
  Dim lCurPos As Long
  Dim lTemp As Long
  Dim tArray() As Byte
  
  If Len(sFileName) = 0 Then Exit Function
  tArray() = StrConv(sFileName, vbFromUnicode)
  If IsArrayEmpty(Not crcTable) Then BuildTable
  
  lTemp = lcrc Xor &HFFFFFFFF 'lcrc is for current value from partial check on the partial array
  
  For lCurPos = 0 To UBound(tArray)
    lTemp = (((lTemp And &HFFFFFF00) \ &H100) And &HFFFFFF) Xor (crcTable((lTemp And 255) Xor tArray(lCurPos)))
  Next lCurPos
  
  CRC32 = lTemp Xor &HFFFFFFFF
  'Returns CRC value
End Function

Private Function BuildTable() As Boolean
  Dim I As Long, x As Long, crc As Long
  Const Limit = &HEDB88320 'usally its shown backward, cant remember what it was.
  'Its the same polynomial that PKZIP uses (I Think)
  ReDim crcTable(0 To 255)
  For I = 0 To 255
    crc = I
    For x = 0 To 7
      If crc And 1 Then
        crc = (((crc And &HFFFFFFFE) \ 2) And &H7FFFFFFF) Xor Limit
      Else
        crc = ((crc And &HFFFFFFFE) \ 2) And &H7FFFFFFF
      End If
    Next x
    crcTable(I) = crc
  Next I
End Function

Public Sub ClearCRCtable()
    Erase crcTable
End Sub

Private Function IsArrayEmpty(ByVal arrayPtr As Long) As Boolean
    IsArrayEmpty = (arrayPtr = -1)
End Function

'============================================
'  Mapping Allocate and Release functions
'============================================
Public Function CreateMapping(ByVal memSize As Long, ByRef fileHandle As Long, Optional ByVal mapName As String = vbNullString, _
            Optional ByVal hForeignWnd As Long, Optional ByVal UseWin9xMap As Boolean = False) As Long
   Dim mapHandle As Long
   Dim processID As Long
   If isNT = True And UseWin9xMap = False Then ' Win9x vs NT
        GetWindowThreadProcessId hForeignWnd, processID
        fileHandle = OpenProcess(PROCESS_VM_OPERATION Or PROCESS_VM_READ Or PROCESS_VM_WRITE, False, processID)
        If fileHandle <> 0 Then _
            mapHandle = VirtualAllocEx(fileHandle, ByVal 0&, ByVal memSize, MEM_RESERVE Or MEM_COMMIT, PAGE_READWRITE)
    Else
        fileHandle = CreateFileMapping(&HFFFFFFFF, 0, PAGE_READWRITE, 0, memSize, mapName)
        If fileHandle <> 0 Then _
            mapHandle = MapViewOfFile(fileHandle, FILE_MAP_ALL_ACCESS, 0, 0, 0)
    End If
    CreateMapping = mapHandle
End Function

Public Sub FreeMapping(ByVal fileHandle As Long, ByVal mapHandle As Long, _
        Optional ByVal UseWin9xMap As Boolean = False)
    
    If fileHandle <> 0 Then
        If isNT = True And UseWin9xMap = False Then
            Call VirtualFreeEx(fileHandle, ByVal mapHandle, 0&, MEM_RELEASE)
        Else
            UnmapViewOfFile mapHandle
        End If
        CloseHandle fileHandle
    End If
End Sub

Public Function isNT() As Boolean
    If m_sysNT = 0 Then
        Dim verinfo As OSVERSIONINFO
        verinfo.dwOSVersionInfoSize = Len(verinfo)
        Call GetVersionEx(verinfo)
        If verinfo.dwPlatformId > 1 Then m_sysNT = 2 Else m_sysNT = 1
    End If
    isNT = (m_sysNT = 2)
End Function

Public Function GetTargetWindow(Optional bProgMgr As Boolean) As Long

Dim lProgMgr As Long, hTarget As Long
lProgMgr = FindWindowEx(GetDesktopWindow(), 0, "Progman", "Program Manager")
If bProgMgr = True Then
    GetTargetWindow = lProgMgr
Else
    If lProgMgr <> 0 Then
        hTarget = FindWindowEx(lProgMgr, 0&, "SHELLDLL_DefView", vbNullString)
        GetTargetWindow = hTarget
    End If
End If

End Function

Public Function IsActiveDesktop() As Boolean
    
    Dim hDtop As Long, hTarget As Long
    hDtop = GetTargetWindow()
    If hDtop <> 0 Then
        hTarget = FindWindowEx(hDtop, 0&, "Internet Explorer_Server", vbNullString)
        IsActiveDesktop = (hTarget <> 0)
    End If

End Function

Public Function DeskTopMsgBox() As Boolean
    
Const clsLen As Long = 7
Dim hTarget As Long
Dim lProgMgr As Long
Dim sClass As String * clsLen
Dim dlgClass As String * clsLen

lProgMgr = GetTargetWindow(True)
dlgClass = "#32770" & Chr$(0)

' use this first; it has potential of being faster....
hTarget = GetLastActivePopup(lProgMgr)
If hTarget <> lProgMgr Then
    If Not IsWindowVisible(hTarget) = 0 Then
        Call GetClassName(hTarget, sClass, clsLen)
        If sClass = dlgClass Then
            DeskTopMsgBox = True
            Exit Function
        End If
    End If
End If

' if above didn't work, we will walk the zorder backwards
hTarget = lProgMgr
Do
    hTarget = GetNextWindow(hTarget, GW_HWNDPREV)
    If hTarget = 0 Then
        Exit Do
    Else
        If GetParent(hTarget) = lProgMgr Then
            If Not IsWindowVisible(hTarget) = 0 Then
                Call GetClassName(hTarget, sClass, clsLen)
                If sClass = dlgClass Then
                    DeskTopMsgBox = True
                    Exit Do
                End If
            End If
        End If
    End If
Loop

End Function


'[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced]
'"HideFileExt"=dword:00000000

'[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer]
'"ShellState"=hex:10,00,00,00,03,00,00,00,00,00,00,00,00,00,00,00

Public Function SpecialFolderPath(ByVal eFolder As eSpecialFolders) As String

        Dim sRtn As String
        Dim sPath As String
        Dim lpil As Long
        Dim lRtn As Long
        Dim IDL As ITEMIDLIST
        Dim hLib As Long
        
        sPath = String$(maxLen, 0)
        
        ' special folder path depending on versino of IE
        hLib = LoadLibrary("shfolder")
        If hLib Then
            If GetProcAddress(hLib, "SHGetFolderPathA") <> 0 Then
                lRtn = SHGetFolderPath(0, eFolder, 0, 0, sPath)
                If lRtn = 0 Then lRtn = 1 Else lRtn = 0
            End If
        End If
        FreeLibrary hLib
        If lRtn = 0 Then
            hLib = LoadLibrary("Shell32")
            If hLib Then
                If GetProcAddress(hLib, "SHGetSpecialFolderLocation") <> 0 Then
                    If SHGetSpecialFolderLocation(0&, eFolder, IDL) = 0 Then
                        lRtn = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
                    End If
                End If
            End If
            FreeLibrary hLib
        End If
        If lRtn = 0 Then
            sRtn = App.Path
        Else
            sRtn = Left$(sPath, InStr(sPath, Chr(0)) - 1)
        End If
        If Right$(sRtn, 1) <> "\" Then sRtn = sRtn & "\"
        
        SpecialFolderPath = sRtn

End Function

Public Function MakeDWord(lWord As Integer, hWord As Integer) As Long
      MakeDWord = (hWord * &H10000) Or (lWord And &HFFFF&)
   End Function
                
Public Function LoWord(DWord As Long) As Integer
      If DWord And &H8000& Then ' &H8000& = &H00008000
         LoWord = DWord Or &HFFFF0000
      Else
         LoWord = DWord And &HFFFF&
      End If
   End Function

Public Function HiWord(DWord As Long) As Integer
      HiWord = (DWord And &HFFFF0000) \ &H10000
End Function

Public Function ForceForeGround(ByVal hwnd As Long)

    Dim lActiveThread As Long, lMyThread As Long
            
    If IsIconic(hwnd) Then
        ShowWindow hwnd, SW_SHOWNORMAL
    ElseIf IsWindowVisible(hwnd) = 0 Then
        ShowWindow hwnd, SW_SHOWNORMAL
    End If
    
    ' source: http://www.thescarms.com/vbasic/alttab.asp. Slightly modified by me
    If hwnd <> GetForegroundWindow() Then
        lActiveThread = GetWindowThreadProcessId(GetForegroundWindow(), ByVal 0&)
        lMyThread = GetWindowThreadProcessId(hwnd, ByVal 0&)
        If lActiveThread <> lMyThread Then
            ' Attach the foreground thread to this window.
            If AttachThreadInput(lActiveThread, lMyThread, True) <> 0 Then
                ' Detach the foreground window's thread from this window.
                Call SetForegroundWindow(hwnd)
                Call AttachThreadInput(lActiveThread, lMyThread, False)
            End If
        End If
        Call SetForegroundWindow(hwnd)
    End If

End Function

