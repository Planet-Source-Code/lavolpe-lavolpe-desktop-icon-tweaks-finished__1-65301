Attribute VB_Name = "lvDTop"
Option Explicit

' Caution. The only portion of this DLL than can safely use all VB-related functions,
'   strings, and formulas is the HookDesktop routine.
' All other routines will most probably crash if strings, VB-functions, and arrays
'   are used as normal. This is because the DLL is being injected into a process
'   that does not have the VB runtimes loaded. If no runtimes, VB-objects crash process.

' Therefore, to paritally overcome lack of VB: constants, APIS, structures and other
'   nice-to-have items were included in the deskTop32sc.TLB file. TLBs allow APIs
'   and structures to be used without need for runtimes.

' The above were general terms for this DLL's compatibility with windows.
' The following statements are specifically for this project....

' This DLL may be final unless it evolves into unicode which would require a few
' extra lines, TLB change, and a new routine or two.

' All SendMessages destined for the host app use SendMessageTimeout to prevent the
'   desktop from freezing up if the host is not responding.

' Desktop Icons.
' 1. Making an icon have a transparent backcolor for the caption is simple
'   enough, you send a message telling the desktop to make item X backcolor
'   equal -1. However, it only applies one time and can be reset back to defaults
'   at almost any time: auto-arranging icons, adding/removing icons, and more.
' 2. For fun, I wanted to try and offer a per-icon custom setting. Each icon can
'   have its own forecolor and backcolor. Other options in future revisions?
' 3. I can find no way of obtaining the "key value" for each item in the desktop.
'   Am assuming this value is internally maintained by O/S. Therefore, the only
'   possible way to uniquely identify icons on the desktop is via their caption.
' 4. The above identification works well but has some minor flaws. For example,
'   you could rename "My Computer" to "Recycle Bin" and O/S has no problems. But
'   this project cannot determine which custom settings belong to which "Recycle Bin".
' 5. One Major Caution here. I have found that it is best not to directly interact with
'   the desktop with one exception: custom drawing icon captions only because the
'   system passes messages to do that. If you interrupt something (refreshing, updating,
'   and possibly drawing) the desktop will probably crash.

Private hOldWndProc As Long ' subclassed window's previous window procedure
Private hListView As Long   ' hWnd for the listview (used to get listitem information)
Private hHostWnd As Long    ' the window we will send messages to
Private msgPrivate As Long  ' a way for the host to talk to this DLL
Private hPopup As Long      ' submenu inserted into desktop context menu
Private hMenuSelect As Long ' result of user clicking contextmenu

Private hIcon As Long       ' imagelist icon passed to host/destroyed when reused or on exit
Private m_DblClck As Long   ' the desktop's classes don't process double clicks,
Private m_SysDblClickTime As Long 'but I want one so we create a simple algo for it

Private hFileRx As Long       ' interprocess comm using a mapped memory file
Private hMapRx As Long        ' the mapped portion of the file
Private hFileTx As Long       ' interprocess comm using a mapped memory file
Private hMapTx As Long        ' the mapped portion of the file
Private m_State As Long        ' 0=not subclassing, 1=full subclass, 2=minimal subclassing

Public Function DllMain(ByVal hInst As Long, ByVal fdwReason As Long, ByVal lpvReserved As Long) As Long
    'DLL Main Entry Point
    DllMain = 1
End Function

' the cusWndProcA and cusWndProcW are nearly identical. If the desktop is a
' unicode window, we will make the effort to use unicode APIs when possible

Private Function cusWndProcA(ByVal hWnd As Long, ByVal wMsg As Long, _
            ByVal wParam As Long, ByVal lParam As Long) As Long
            
    Dim bHandled As Boolean
    
    Select Case wMsg
    Case WM_NOTIFY
        cusWndProcA = On_Notify(hWnd, wParam, lParam, bHandled, False)
    Case msgPrivate
        cusWndProcA = On_PrivateMsg(hWnd, wParam, lParam, bHandled)
        If cusWndProcA = lvNULL Then PostMessage hWnd, WM_NULL, hWnd, -hWnd ' abort subclassing
    Case WM_DESTROY
        SetWindowLongA hWnd, GWL_WNDPROC, hOldWndProc
        cusWndProcA = CallWindowProcA(hOldWndProc, hWnd, wMsg, wParam, lParam)
        Call On_Destroy(bHandled)
    Case WM_NULL
        If wParam = hWnd And lParam = -hWnd Then
            SetWindowLongA hWnd, GWL_WNDPROC, hOldWndProc
            cusWndProcA = On_Destroy(bHandled)
        End If
    Case WM_ACTIVATEAPP
        If wParam = lvNULL Then
            Call On_Misc(WM_KILLFOCUS, bHandled)
            cusWndProcA = CallWindowProcA(hOldWndProc, hWnd, wMsg, wParam, lParam)
            Call On_LostFocus(bHandled)
        Else
            Call On_Misc(WM_SETFOCUS, True)
        End If
        
    Case WM_PARENTNOTIFY    ' dbl click algo
        If wParam = WM_LBUTTONDOWN Then
            Dim clkTime As Long
            clkTime = GetTickCount
            If clkTime - m_DblClck < m_SysDblClickTime Then _
                PostMessage hHostWnd, msgPrivate, WM_LBUTTONDBLCLK, lParam
            m_DblClck = clkTime
        End If
        
    Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN, WM_XBUTTONDOWN
        ' only received if the desktop is hidden
        Call On_Misc(WM_SETFOCUS, True)
    
    Case WM_INITMENU
        Call On_InitPopup(wParam, False)
    Case WM_MENUSELECT
        If lParam = hPopup Then
            hMenuSelect = wParam
        Else
            If Not lParam = lvNULL Then hMenuSelect = lvNULL
        End If
    Case WM_EXITMENULOOP
        Call On_ExitPopup
        
    End Select
    
    If Not bHandled Then cusWndProcA = CallWindowProcA(hOldWndProc, hWnd, wMsg, wParam, lParam)

End Function

Private Function cusWndProcW(ByVal hWnd As Long, ByVal wMsg As Long, _
            ByVal wParam As Long, ByVal lParam As Long) As Long
            
    Dim bHandled As Boolean
    
    Select Case wMsg
        Case WM_NOTIFY
            cusWndProcW = On_Notify(hWnd, wParam, lParam, bHandled, True)
        Case msgPrivate
            cusWndProcW = On_PrivateMsg(hWnd, wParam, lParam, bHandled)
            If cusWndProcW = lvNULL Then PostMessage hWnd, WM_NULL, hWnd, -hWnd ' abort subclassing
        Case WM_DESTROY
            SetWindowLongW hWnd, GWL_WNDPROC, hOldWndProc
            cusWndProcW = CallWindowProcW(hOldWndProc, hWnd, wMsg, wParam, lParam)
            Call On_Destroy(bHandled)
        Case WM_NULL
            If wParam = hWnd And lParam = -hWnd Then
                SetWindowLongW hWnd, GWL_WNDPROC, hOldWndProc
                cusWndProcW = On_Destroy(bHandled)
            End If
        Case WM_ACTIVATEAPP
            If wParam = lvNULL Then
                Call On_Misc(WM_KILLFOCUS, bHandled)
                cusWndProcW = CallWindowProcW(hOldWndProc, hWnd, wMsg, wParam, lParam)
                Call On_LostFocus(bHandled)
            Else
                Call On_Misc(WM_SETFOCUS, True)
            End If
        
        Case WM_PARENTNOTIFY ' dbl click algo
            If wParam = WM_LBUTTONDOWN Then
                Dim clkTime As Long
                clkTime = GetTickCount
                If clkTime - m_DblClck < m_SysDblClickTime Then _
                    PostMessage hHostWnd, msgPrivate, WM_LBUTTONDBLCLK, lParam
                m_DblClck = clkTime
            End If
        
        Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN, WM_XBUTTONDOWN
            ' only received if the desktop is hidden
            Call On_Misc(WM_SETFOCUS, True)
        
        Case WM_INITMENU
            Call On_InitPopup(wParam, True)
        Case WM_MENUSELECT
            If lParam = hPopup Then
                hMenuSelect = wParam
            Else
                If Not lParam = lvNULL Then hMenuSelect = lvNULL
            End If
        Case WM_EXITMENULOOP
            Call On_ExitPopup
            
    End Select

    If Not bHandled Then cusWndProcW = CallWindowProcW(hOldWndProc, hWnd, wMsg, wParam, lParam)

End Function

Private Function HookDesktop(ByVal hWndParent As Long, _
        ByRef mapFileTx As String, _
        ByRef mapFileRx As String, _
        ByRef PrivateMsg As Long) As Long
    
' This function is called from host to begin or resume hooking into the desktop
' The dll at this point is in the host process, a VB process

' Required Parameters:
' hWndParent :: the hWnd of host application; host must be subclassed to receive messages
' mapFileRx :: the file name the host must use to create a mapped file for receipt from this DLL
' mapFileTx :: the file name the host must use to create a mapped file for transmission to this DLL
' PrivateMsg :: the WM message used for Send/PostMessage calls between host & DLL

' Return value:  non-zero indicates successful injection of DLL into desktop process

    Dim hHook As Long
    Dim hThread As Long
    Dim hMod As Long
    Dim hTarget As Long
    Dim lHookMsg As Long
    
    If hWndParent = lvNULL Then Exit Function
    
    hTarget = FindDeskTop   ' find the desktop window we want to subclass
    If hTarget = lvNULL Then Exit Function
        
    ' get the thread that window resides in (we don't want a global hook)
    hThread = GetWindowThreadProcessId(hTarget, ByVal 0&)
    ' get instance handle to our DLL
    hMod = GetModuleHandle("DTopTweaker")   ' must be name of this DLL
    ' register a custom message to communicate with that DLL when it gets injected
    lHookMsg = RegisterWindowMessage(WM_cPrivate)
    
    If hMod <> lvNULL And hThread <> lvNULL And lHookMsg <> lvNULL Then
        ' set a thread-specific hook
        hHook = SetWindowsHookEx(WH_CBT, AddressOf hookProc, hMod, hThread)
        ' send a message to the target window. This will activate the hook
        SendMessage hTarget, WM_SYSCOMMAND, lvNULL, ByVal lHookMsg
        ' unhook the thread now
        UnhookWindowsHookEx hHook
        ' send a specific message to ensure it is responding
        ' This message also passes the host hWnd so the DLL knows who to talk to
        If SendMessage(hTarget, lHookMsg, -lHookMsg, ByVal hWndParent) = lHookMsg Then
            HookDesktop = lHookMsg
            mapFileTx = StrConv(WM_cPrivate, vbFromUnicode) ' VB expects ANSI from DLL
            mapFileRx = StrConv(WM_altPrivate, vbFromUnicode) ' VB expects ANSI from DLL
            PrivateMsg = lHookMsg ' system-unique message for use between DLL & host
        Else ' something is wrong; abort. Specific message to abort subclassing
            SendMessage hTarget, WM_NULL, hTarget, ByVal -hTarget
        End If
    End If


End Function

Private Function hookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' This is only called once and we begin subclassing when it is called.
' The DLL at this point now exists in the target process and usage of normal VB-related
'   functions, strings, arrays WILL crash the process unless handled very carefully.
    
    Dim hLib As Long, bUnicode As Long
    Dim hNewWndProc As Long, targetHwnd As Long
    
    If m_State = lvNULL Then         ' prevent recurrence (if even possible)
        
        ' the DLL is now in the target process, but we don't know if the target hWnd
        ' hooked this DLL or another window on the thread hooked it. Therefore, we
        ' shouldn't rely on the hWnd provided in the structure pointed to by lParam.
        m_State = 1
        targetHwnd = FindDeskTop            ' find the window we want subclassed
        If targetHwnd <> lvNULL Then
            ' create a private message the DLL can use to communicate with the host
            msgPrivate = RegisterWindowMessage(WM_cPrivate)
            
            ' Need to up DLL ref count otherwise DLL will be unmapped when this routine returns
            hLib = LoadLibrary("DTopTweaker")
            If Not hLib = lvNULL Then
                
                ' if our target is a unicode window, let's use a unicode subclass procedure
                bUnicode = IsWindowUnicode(targetHwnd)
                ' can't use AddressOf in non-VB process, do it the old fashioned way
                If bUnicode = lvNULL Then
                    hNewWndProc = GetProcAddress(hLib, "cusWndProcA")
                Else
                    hNewWndProc = GetProcAddress(hLib, "cusWndProcW")
                End If
                If Not hNewWndProc = lvNULL Then    ' subclass
                    If bUnicode = lvNULL Then
                        hOldWndProc = SetWindowLongA(targetHwnd, GWL_WNDPROC, hNewWndProc)
                    Else
                        hOldWndProc = SetWindowLongW(targetHwnd, GWL_WNDPROC, hNewWndProc)
                    End If
                End If
                If hNewWndProc = lvNULL Or hOldWndProc = lvNULL Then
                    ' getting here should be impossible, but because it is a foreign
                    ' process we checked anyway
                    FreeLibrary hLib
                    m_State = lvNULL
                End If
            Else
                m_State = lvNULL
            End If
        End If
    End If
    hookProc = 1
    
End Function

Private Function FindDeskTop() As Long
' short function whose only purpose is to identify the SysListView32 object
'   belonging to the desktop and that object's parent (which we will subclass)

    Dim lProgMgr As Long
    Dim lShellDLL As Long
    
    lProgMgr = FindWindowEx(GetDesktopWindow(), lvNULL, "Progman", "Program Manager")
    If lProgMgr <> lvNULL Then
        lShellDLL = FindWindowEx(lProgMgr, lvNULL, "SHELLDLL_DefView", vbNullString)
        hListView = FindWindowEx(lShellDLL, lvNULL, "SysListView32", vbNullString)
        FindDeskTop = lShellDLL
    End If

End Function

Private Function CreateMapping() As Long
    
    ' Interprocess communication between the host and this DLL will be handled
    ' via mapped memory files.  The communication will be initiated 99% of the
    ' time from the DLL. Create the mapping now. The On_PrivateMsg routine is
    ' our DLL answering the host initiated questions
    
    ' Note: WM_cPrivate & WM_altPrivate are GUIDs and should reduce the possibility
    '   of us trying to use an existing memory file or someone else using ours.
    
    ' Create 2 small 1kb memory files, large enough for what we want to do.
    
    ' Create the file to be used for Receipt from the DLL
    hFileRx = CreateFileMapping(-1, ByVal 0&, PAGE_READWRITE, lvNULL, 1024, WM_cPrivate)
    If hFileRx <> lvNULL Then
        ' now map the entire file into our process so we can read/write from it
        hMapRx = MapViewOfFile(hFileRx, FILE_MAP_WRITE, lvNULL, lvNULL, lvNULL)
        
        ' now create the file to be used for Transmission to the DLL
        If hMapRx <> lvNULL Then
            hFileTx = CreateFileMapping(-1, ByVal 0&, PAGE_READWRITE, lvNULL, 1024, WM_altPrivate)
            If hFileTx <> lvNULL Then
                ' now map the entire file into our process so we can read/write from it
                hMapTx = MapViewOfFile(hFileTx, FILE_MAP_WRITE, lvNULL, lvNULL, lvNULL)
            End If
        End If
    End If
    
    ' sanity checks
    If hMapTx = lvNULL Then ' failure
        If hFileTx <> lvNULL Then CloseHandle hFileTx
        If hMapRx <> lvNULL Then UnmapViewOfFile hMapRx
        If hFileRx <> lvNULL Then CloseHandle hFileRx
    Else
        CreateMapping = 1
    End If
    
End Function

Private Sub On_Misc(ByVal wMsg As Long, ByRef bHandled As Boolean)
    
    ' some messages are informative to the host. Simply pass it along
    If Not hHostWnd = lvNULL Then PostMessage hHostWnd, msgPrivate, wMsg, lvNULL
    bHandled = True

End Sub

Private Function On_Notify(ByVal hWnd As Long, ByVal wParam As Long, _
        ByVal lParam As Long, ByRef isHandled As Boolean, ByVal wideChar As Boolean) As Long

' called from the ANSI and UNICODE window procedures...
' this is where we will handle custom draw messages and edit-type messages

    If hMapTx = lvNULL Then Exit Function   ' do we have a mapped file yet?
                        
    Dim tCD As NMLVCUSTOMDRAW
        
    ' see if the message we got is custom-draw and from our listview
    CopyMemory tCD, ByVal lParam, 12 ' size of the header portion of the structure
    If tCD.nmcd.hdr.hwndFrom = hListView Then
        
        Dim lRtn As Long, lItemNr As Long
        
        If Not m_State = 1 Then Exit Function  ' we are not drawing
        
            Select Case tCD.nmcd.hdr.code
            Case NM_CUSTOMDRAW
        
                ' get all the pertinent information from the lParam pointer
                CopyMemory tCD, ByVal lParam, 56 ' full structure is 84 for newest version
                
                Select Case tCD.nmcd.dwDrawStage
                
                    Case CDDS_PREPAINT  ' initial message meaning: Do we want customdrawn messages?
                        
                        On_Notify = CDRF_NOTIFYITEMDRAW ' yes, we want paint notifications
                        isHandled = True
                    
                    Case CDDS_ITEMPREPAINT ' icon is about to be painted
                        
                        'FYI:
                        'tCD.nmcd.dwItemSpec  appears to be the item number (changes)
                        'tCD.nmcd.lItemlParam appears to be dynamic pointer
                        ' lItem.lParam when retrieved also appears to be dynamic
                        
                        ' The layout of the shared mapped files are as follows
                        ' bytes 0-3, caption fore color
                        ' bytes 4-7, caption back color
                        ' bytes 8-43, LVITEM structure
                        ' bytes 44-99, reserved/buffer
                        ' bytes 100+ , the caption
                        If Not GetItemData(tCD.nmcd.dwItemSpec, VarPtr(tCD.clrText)) = lvNULL Then
                            If Not SendMessageTimeout(hHostWnd, msgPrivate, NM_CUSTOMDRAW, ByVal tCD.nmcd.dwItemSpec, SMTO_ABORTIFHUNG, 100, lRtn) = lvNULL Then
                                If Not lRtn = lvNULL Then
                                    ' set the text forecolor & backcolor
                                    CopyMemory ByVal VarPtr(tCD.clrText), ByVal hMapRx, 8
                                    ' save the changes back to the LVITEM structure
                                    CopyMemory ByVal lParam, tCD, 56 ' full structure is 84 for newest version
                                    isHandled = True
                                End If
                            End If
                        End If
                        ' tell listview to finish drawing (with any changes we made)
                        On_Notify = CDRF_DODEFAULT
                End Select
            
            Case LVN_GETDISPINFOA, LVN_GETDISPINFOW
            '^^ accepted/edited icon caption
                
                Dim lvDI As LV_DISPINFO
                ' get all the pertinent information from the lParam pointer
                CopyMemory lvDI, ByVal lParam, 56
                
                If (lvDI.Item.mask And LVIF_TEXT) = LVIF_TEXT Then
                
                    ' let it go thru first, then pass info to the host
                    If wideChar Then
                        lRtn = CallWindowProcW(hOldWndProc, hWnd, WM_NOTIFY, wParam, lParam)
                    Else
                        lRtn = CallWindowProcA(hOldWndProc, hWnd, WM_NOTIFY, wParam, lParam)
                    End If
                    
                    ZeroMemory ByVal hMapTx + 100, cMaxLen ' clear any existing caption
                    If tCD.nmcd.hdr.code = LVN_GETDISPINFOA Then ' ANSI
                        lstrcpyA ByVal hMapTx + 100, ByVal lvDI.Item.pszText
                    Else    ' Convert from wide char to ANSI for memory mapped file
                        lvDI.Item.cchTextMax = lstrlenW(ByVal lvDI.Item.pszText)
                        WideCharToMultiByte 0, WC_NO_BEST_FIT_CHARS, ByVal lvDI.Item.pszText, lvDI.Item.cchTextMax, ByVal hMapTx + 100, cMaxLen, 0, 0
                    End If
    
                    ' inform host the item has changed.
                    SendMessageTimeout hHostWnd, msgPrivate, LVN_GETDISPINFOA, ByVal lvDI.Item.iItem, SMTO_ABORTIFHUNG, 100, ByVal 0&
                    On_Notify = lRtn
                    isHandled = True
                    
                End If
            
            Case LVN_DELETEITEM, LVN_INSERTITEM
                PostMessage hHostWnd, msgPrivate, tCD.nmcd.hdr.code, ByVal 0&
                '^^ trial & error shows interruption here can carsh explorer
            End Select
            
        End If
        
        ' prevents sending host a doubleclick if user double clicks on an icon
        ' vs doubleclicking on the desktop
        If tCD.nmcd.hdr.code = LVN_ITEMCHANGED Then
            If m_DblClck > lvNULL Then m_DblClck = lvNULL
        End If
    
End Function

Private Function On_PrivateMsg(ByVal hWnd As Long, ByVal wParam As Long, _
                ByVal lParam As Long, ByRef isHandled As Boolean) As Long
    
'   Message was sent from the host.
'   Future versions may communicate more with the DLL

' Special messages and combinations thereof are used. This DLL is not a generic DLL.
' It is specifically designed for the host, so the host knows how to talk to
' the DLL and the DLL knows how to talk to the host.

    Dim lRtn As Long '<< if lRtn is zero, subclassing is aborted
    Select Case wParam
    
        Case -msgPrivate ' initial comm btwn host and DLL
            If hHostWnd = lvNULL Then   ' first time thru
                hHostWnd = lParam       ' talk to host via this hWnd
                m_SysDblClickTime = GetDoubleClickTime() + 1
                If Not CreateMapping = lvNULL Then lRtn = msgPrivate
            Else ' 2nd time thru, if we are already set up, ensure a non-zero return value
                If Not hFileTx = lvNULL Then lRtn = msgPrivate
            End If
            
        Case WM_NOTIFY
            Select Case lParam
            Case WM_CLOSE   ' terminate subclassing
                SendMessage hWnd, WM_NULL, hWnd, ByVal -hWnd
                lRtn = 1
            Case WM_ENDSESSION ' stop custom drawing; minimal subclassing
                m_State = 2
                lRtn = 1
            Case WM_ACTIVATE    ' restart custom drawing
                If Not m_State = 1 Then
                    m_State = 1
                    lRtn = hHostWnd ' meaningful to the host app
                Else
                    lRtn = 1
                End If
            End Select
        
        Case Else               ' reserved for future expansion
            If LoWord(wParam) = WM_GETICON Then
                ' host wants the icon for the passed index in lParam
                Dim hIL As Long ' get the imagelist hWnd
                ' the HiWord is either 0=large icons, or 1=small icons
                hIL = SendMessage(hListView, LVM_GETIMAGELIST, HiWord(wParam), ByVal lvNULL)
                If Not hIL = lvNULL Then
                    ' destroy previously created icon, if any
                    If Not hIcon = lvNULL Then DestroyIcon hIcon
                    ' have imagelist create icon for us
                    lRtn = ImageList_GetIcon(hIL, lParam, ByVal 1&)
                    hIcon = lRtn
                End If
                If lRtn = 0 Then lRtn = 1
            Else
                lRtn = 1
            End If
    End Select
    isHandled = True
    On_PrivateMsg = lRtn
    
End Function

Private Function On_Destroy(ByRef isHandled As Boolean) As Long
    
' Clean up when subclassing is terminated
    m_State = lvNULL
    ' When we've finished with it, unmap the file-mapping objects
    ' from our address space and release the handles.
    If Not hMapTx = lvNULL Then UnmapViewOfFile ByVal hMapTx
    If Not hFileTx = lvNULL Then CloseHandle hFileTx
    If Not hMapRx = lvNULL Then UnmapViewOfFile ByVal hMapRx
    If Not hFileRx = lvNULL Then CloseHandle hFileRx
    If Not hIcon = lvNULL Then DestroyIcon hIcon
    hMapTx = lvNULL
    hFileTx = lvNULL
    hMapRx = lvNULL
    hFileRx = lvNULL
    'hOldWndProc = lvNULL << no can do; if nested messages, won't be able to continue forwarding to DefWndProc
    hHostWnd = lvNULL
    hListView = lvNULL
    msgPrivate = lvNULL
    hIcon = lvNULL
    isHandled = True
    On_Destroy = 1

End Function


Private Sub On_LostFocus(ByRef bHandled As Boolean)
    
    Dim lItem As tlbLVITEM
    Dim ItemNr As Long
    
    ' the desktop won't always refresh properly when user clicked on icon & then
    ' clicked off onto another application window... Redraw the selected item(s)
    For ItemNr = lvNULL To SendMessage(hListView, LVM_GETITEMCOUNT, lvNULL, ByVal lvNULL) - 1

        With lItem
            .iItem = ItemNr
            .mask = LVIF_STATE
            .stateMask = LVIS_SELECTED
            .state = lvNULL
        End With
        SendMessage hListView, LVM_GETITEM, lvNULL, ByVal VarPtr(lItem)
        If lItem.state = lItem.stateMask Then
            PostMessage hListView, LVM_REDRAWITEMS, ItemNr, ItemNr
        End If
    Next
    bHandled = True
    
End Sub

Private Sub On_InitPopup(ByVal wParam As Long, ByVal wideChar As Boolean)
                
' Add a popup menu into the desktop's context menu when it is displayed

    Dim lFlags As Long
    Dim lStates As Long
    '^^ 1=Active, 2=AutoHide,
    '   4=can restore, 8=can undo restore, ++ reserved for future options
    
    ' note the WM_USER+# below. Each of the # values relate 1:1 to the
    ' mnuPopup() array in the host application. This way when the host app gets
    ' a menuselection, it simply needs to subtract WM_USER for the appropriate mnuPopup
    
        hMenuSelect = lvNULL
        If SendMessageTimeout(hHostWnd, msgPrivate, WM_SETTINGCHANGE, ByVal 0&, SMTO_ABORTIFHUNG, 100, lStates) = 0 Then Exit Sub
        ' create a submenu to be inserted into the popup menu about to be displayed
        hPopup = CreateMenu() ' hpopup will be destroyed when menu closes
        AppendMenu hPopup, MF_STRING, WM_USER + 7, mnuSave
        If (lStates And 4) = lvNULL Then lFlags = MF_DISABLED Or MF_GRAYED Else lFlags = lvNULL
        AppendMenu hPopup, MF_STRING Or lFlags, WM_USER + 8, mnuRestore
        If (lStates And 8) = lvNULL Then lFlags = MF_DISABLED Or MF_GRAYED Else lFlags = lvNULL
        AppendMenu hPopup, MF_STRING Or lFlags, WM_USER + 9, mnuUnRestore
        AppendMenu hPopup, MF_SEPARATOR, lvNULL, ""
        AppendMenu hPopup, MF_STRING, WM_USER + 4, mnuHide ' the mnu... variables are defined in the TLB
        If (lStates And 2) = 2 Then lFlags = MF_CHECKED
        AppendMenu hPopup, MF_STRING Or lFlags, WM_USER + 5, mnuAuto
        AppendMenu hPopup, MF_SEPARATOR, lvNULL, ""
        If (lStates And 1) = 1 Then lFlags = MF_CHECKED Else lFlags = lvNULL
        AppendMenu hPopup, MF_STRING Or lFlags, WM_USER + 11, mnuActive
        AppendMenu hPopup, MF_SEPARATOR, lvNULL, ""
        AppendMenu hPopup, MF_STRING, WM_USER + 1, mnuApp
        ' now add this new submenu to the 2nd to last item in the popup menu
        InsertMenu wParam, GetMenuItemCount(wParam) - 1, MF_POPUP Or MF_STRING Or MF_BYPOSITION, hPopup, mnuMain
        ' now add separator to the 2nd to last item in the popup menu
        InsertMenu wParam, GetMenuItemCount(wParam) - 1, MF_SEPARATOR Or MF_BYPOSITION, lvNULL, ""
        
End Sub

Private Sub On_ExitPopup()

' Determine if our custom menu items were selected in the desktop's context menu

    If hPopup = lvNULL Then Exit Sub ' shouldn't happen
     
    ' clean up
    DestroyMenu hPopup
    hPopup = lvNULL
    ' send host app the result
    hMenuSelect = LoWord(hMenuSelect)
    If hMenuSelect > WM_USER Then PostMessage hHostWnd, msgPrivate, WM_MENUSELECT, ByVal hMenuSelect
    
End Sub

Private Function LoWord(ByVal xWord As Long) As Long
    
    ' the menu item selected is the LoWord of the hMenuSelect value
    If xWord And &H8000& Then
       LoWord = xWord Or &HFFFF0000
    Else
       LoWord = xWord And &HFFFF&
    End If
    
End Function

Private Function HiWord(ByVal xWord As Long) As Long

    HiWord = (xWord And &HFFFF0000) \ &H10000
    
End Function

Private Function GetItemData(ByVal ItemNr As Long, ByVal ptrColors As Long) As Long

    Dim lRtn As Long
    Dim lItem As tlbLVITEM
    '^^ LEN OF LVITEM varies depending on IE & O/S versions (from 36 to to 52 bytes)
    '   We will use the first 36 bytes for our purposes
    
    ' The layout of the shared mapped files are as follows
    ' bytes 0-3, caption fore color
    ' bytes 4-7, caption back color
    ' bytes 8-43, LVITEM structure
    ' bytes 44-99, reserved/buffer
    ' bytes 100+ , the caption
    With lItem
        .mask = LVIF_TEXT Or LVIF_IMAGE Or LVIF_STATE
        .iItem = ItemNr
        .cchTextMax = cMaxLen   ' max characters for caption
        .pszText = hMapTx + 100 ' use mapped file for the return text (ANSI)
    End With
    '**Note. A listview can have unlimited characters for the caption. However,
    ' per MSDN documentation, only 260 can be displayed. Therefore cMaxLen=260
    
    ZeroMemory ByVal hMapTx + 100, cMaxLen ' clear out any previous caption
    CopyMemory ByVal hMapTx, ByVal ptrColors, 8
    CopyMemory ByVal hMapTx + 8, ByVal VarPtr(lItem), 36
    lRtn = SendMessage(hListView, LVM_GETITEM, ItemNr, ByVal hMapTx + 8)
    GetItemData = lRtn
    
End Function
