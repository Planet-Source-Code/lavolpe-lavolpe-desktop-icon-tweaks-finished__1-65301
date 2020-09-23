Attribute VB_Name = "modDLLInstall"
Option Explicit

Private Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
   dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
   dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
   dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
   dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
   dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
   dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
   dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
   dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
   dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
   dwFileFlagsMask As Long        '  = &h3F for version "0.42"
   dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
   dwFileType As Long             '  e.g. VFT_DRIVER
   dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long           '  e.g. 0
   dwFileDateLS As Long           '  e.g. 0
End Type
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long

Public Function InstallDLL(ByVal DLLname As String, ByVal resNumber As Long, _
            Optional minVerRqd As String = vbNullString, Optional ByVal resSection As String = "Custom") As Boolean

    ' function can install needed dlls from a resource file
    ' (unless privileges prevent writing to system folder)

' DLLname :: the name of the DLL  (i.e., DTopTweaker)
' resNumber :: the resource number for the dll (i.e., 101)
' minVerRqd (Optional) :: if passed, the existing DLL will be compared to that version
'           otherwise the dll in the res file will be extracted and the version will
'           be retrieved from that copy for comparison
' resSection (Optional) :: resource section for the RES dll, default is Custom

    Dim targetPath As String        ' system folder
    Dim dllVer(1 To 2) As String    ' version of target & source DLLs
    Dim bDat() As Byte              ' array for res file extraction
    Dim v1() As String, v2() As String  ' arrays for version comparision
    
    Dim fNr As Integer
    Dim I As Integer
    Dim bInstall As Boolean
    
    On Error GoTo EndInstall
    I = InStr(DLLname, "\") ' ensure a full path wasn't passed
    Do Until I = 0
        DLLname = Mid$(DLLname, I + 1)
    Loop                    ' ensure passed name has a .dll extension
    If LCase$(Right$(DLLname, 4)) <> ".dll" Then DLLname = DLLname & ".dll"
    
    targetPath = SpecialFolderPath(CSIDL_SYSTEM)    ' get system folder
    
    If Len(Dir$(targetPath & DLLname, vbHidden Or vbReadOnly Or vbSystem)) = 0 Then
        ' if the target file doesn't exist, no version comparison needed
        bInstall = True
        minVerRqd = ""  ' forces extraction from RES file
    Else
        If Len(minVerRqd) = 0 Then ' when minimal version required not passed
            ' get source dll from the resource file
            If Len(Dir$(App.Path & "\" & DLLname, vbHidden Or vbReadOnly Or vbSystem)) > 0 Then
                SetAttr App.Path & "\" & DLLname, vbNormal
                Kill App.Path & "\" & DLLname
            End If
            bDat() = LoadResData(resNumber, resSection)
            fNr = FreeFile()
            Open App.Path & "\" & DLLname For Binary As #fNr
            Put #fNr, , bDat
            Close #fNr
            ' get the version of the source dll
            dllVer(1) = DisplayVerInfo(App.Path & "\" & DLLname)
        Else ' use the passed version
            dllVer(1) = minVerRqd
        End If
        ' get the version of target dll
        dllVer(2) = DisplayVerInfo(targetPath & DLLname)
        ' compare two versions, fast check first
        If Not dllVer(1) = dllVer(2) Then
            v1() = Split(dllVer(1), ".")
            v2() = Split(dllVer(2), ".")
            If UBound(v1) < 3 Then ReDim Preserve v1(0 To 3)
            If UBound(v2) < 3 Then ReDim Preserve v2(0 To 3)
            For I = 0 To 3
                If Val(v1(I)) > Val(v2(I)) Then
                    bInstall = True
                    Exit For
                End If
            Next
        End If
    End If
        
    If bInstall Then
        If Len(minVerRqd) = 0 Then ' we'll need to extract dll from resource file
            bDat() = LoadResData(resNumber, resSection)
            fNr = FreeFile()
            Open App.Path & "\" & DLLname For Binary As #fNr
            Put #fNr, , bDat
            Close #fNr
        End If
        If Len(Dir$(targetPath & DLLname, vbHidden Or vbReadOnly Or vbSystem)) > 0 Then
            SetAttr targetPath & DLLname, vbNormal
            Kill targetPath & DLLname
        End If
        FileCopy App.Path & "\" & DLLname, targetPath & DLLname
    End If
    
EndInstall:
If Err Then
    InstallDLL = Err.Number
    Err.Clear
End If
On Error Resume Next
If bInstall = True Or Len(minVerRqd) = 0 Then
    ' Remove the RES extracted dll from our app.path
    If Len(Dir$(App.Path & "\" & DLLname, vbHidden Or vbReadOnly Or vbSystem)) > 0 Then
        Kill App.Path & "\" & DLLname
    End If
End If
If Err Then Err.Clear
End Function


Private Function DisplayVerInfo(FullFileName As String) As String
   
   Dim rc As Long, sBuffer() As Byte
   Dim lBufferLen As Long, lVerPointer As Long
   Dim lVerbufferLen As Long
   Dim udtVerBuffer As VS_FIXEDFILEINFO
   Dim sRtn As String

   '*** Get size ****
   lBufferLen = GetFileVersionInfoSize(FullFileName, rc)
   If lBufferLen > 1 Then

        '**** Store info to udtVerBuffer struct ****
        ReDim sBuffer(lBufferLen)
        rc = GetFileVersionInfo(FullFileName, 0&, lBufferLen, sBuffer(0))
        rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
        CopyMemory udtVerBuffer, ByVal lVerPointer, Len(udtVerBuffer)
    
        '**** Determine Product Version number ****
        sRtn = udtVerBuffer.dwProductVersionMSh & "." & udtVerBuffer.dwProductVersionMSl & "." & udtVerBuffer.dwProductVersionLSh & "." & udtVerBuffer.dwProductVersionLSl
   
   End If
   
   If Len(sRtn) = 0 Then sRtn = "0.0.0.0"
   
   DisplayVerInfo = sRtn

End Function

