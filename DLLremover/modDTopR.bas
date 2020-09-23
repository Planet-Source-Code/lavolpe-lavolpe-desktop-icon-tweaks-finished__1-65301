Attribute VB_Name = "modDTopR"
Option Explicit

' HYBRID DLL

' Its only purpose is to go in and remove the DTopTweaker hybrid DLL.

' When DTopTweaker gets injected into the desktop process (any process really). That
' DLL needs to up its reference count by one before the injection routine returns;
' otherwise it will be unmapped from the process.  However, that DLL cannot
' dereference itself from within itself without crashing the injected process.
' To get around this, this DLL is created which will get injected into the same
' target process, look for the DTopTweaker.dll and dereference it. This all occurs
' during the initial hook and this DLL does not up its reference count so it is
' automatically uninjected when the UnhookDesktop function returns.

' The only hitch is that the DTopTweaker.dll must already have stopped subclassing,
' otherwise, crashing the process is likely. Therefore a check is made.

Public Function DllMain(ByVal hInst As Long, ByVal fdwReason As Long, ByVal lpvReserved As Long) As Long
'DLL Main Entry Point
    DllMain = 1
End Function

Private Function UnhookDesktop() As Long
    
    Dim hHook As Long
    Dim hThread As Long
    Dim hMod As Long
    Dim hTarget As Long
    Dim lHookMsg As Long
    
    ' return value indicates success. But in reality a non-crash indicates success
    
    hTarget = FindDeskTop   ' find the desktop window we want to subclass
    If hTarget = lvNULL Then Exit Function
        
    ' get the thread that window resides in (we don't want a global hook)
    hThread = GetWindowThreadProcessId(hTarget, ByVal 0&)
    ' get instance handle to our DLL
    hMod = GetModuleHandle("DTopTwkR")   ' must be name of this DLL
    ' register a custom message to communicate with that DLL when it gets injected
    lHookMsg = RegisterWindowMessage(WM_cPrivate)
    
    If hMod <> lvNULL And hThread <> lvNULL And lHookMsg <> lvNULL Then
        ' specific to the DTopTweaker.dll to stop stubclassing
        SendMessage hTarget, WM_NULL, hTarget, ByVal -hTarget
        ' set a thread-specific hook
        hHook = SetWindowsHookEx(WH_CBT, AddressOf hookProc, hMod, hThread)
        ' send a message to the target window. This will activate the hook
        SendMessage hTarget, WM_SYSCOMMAND, lvNULL, ByVal lHookMsg
        ' unhook the thread now, the thread should have been hooked
        UnhookDesktop = UnhookWindowsHookEx(hHook)
    End If


End Function

Private Function hookProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    ' This function resides in the target process
    Dim hLib As Long
    ' see if the DTopTweaker is mapped into the process
    hLib = GetModuleHandle("DTopTweaker")
    ' if so unmap/uninject it now
    If Not hLib = lvNULL Then FreeLibrary hLib
    hookProc = 1
    
End Function

Private Function FindDeskTop() As Long
' short function whose only purpose is to identify the SysListView32 object's parent

    Dim lProgMgr As Long
    
    lProgMgr = FindWindowEx(GetDesktopWindow(), lvNULL, "Progman", "Program Manager")
    If lProgMgr <> lvNULL Then
        FindDeskTop = FindWindowEx(lProgMgr, lvNULL, "SHELLDLL_DefView", vbNullString)
    End If

End Function

