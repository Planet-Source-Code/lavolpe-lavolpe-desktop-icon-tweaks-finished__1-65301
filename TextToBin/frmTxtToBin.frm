VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmTxtToBin 
   Caption         =   "Form1"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3525
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   3525
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2745
      Top             =   165
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   3645
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   165
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create DLL && TLB Files"
      Height          =   495
      Left            =   195
      TabIndex        =   0
      Top             =   165
      Width           =   3195
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1515
      Left            =   225
      TabIndex        =   2
      Top             =   810
      Width           =   3135
   End
End
Attribute VB_Name = "frmTxtToBin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Sub Command1_Click()

' nothing really fancy. Just get some string data & write it as binary

Dim fnr(1 To 2) As Integer
Dim s As String
Dim b(0 To 3) As Byte
Dim x As Long
Dim i As Integer

With CommonDialog1
    .CancelError = True
    .Flags = cdlOFNExplorer Or cdlOFNFileMustExist
    .DialogTitle = "Select either VBDLL.Txt or VBTLB.txt"
End With
On Error GoTo ExitRoutine
CommonDialog1.ShowOpen

Dim sPath As String
sPath = Left$(CommonDialog1.FileName, InStrRev(CommonDialog1.FileName, "\"))

Dim f(1 To 6) As String
f(1) = sPath & "VBDLL.TXT"
f(2) = sPath & "VBDLLREMOVE.TXT"
f(3) = sPath & "VBTLB.TXT"
f(4) = sPath & "DTopTweaker.dll"
f(5) = sPath & "DTopTwkR.dll"
f(6) = sPath & "deskTop32sc.tlb"

' verify exists
For i = 1 To 3
    If Len(Dir$(f(i), vbHidden Or vbReadOnly Or vbSystem)) > 0 Then
        SetAttr f(i), vbNormal
    Else
        MsgBox "Invalid path, the two files were not found. Try again", vbInformation + vbOKOnly, "Oops"
        Exit Sub
    End If
Next

' kill dest files if they already exist
For i = 4 To 6
    If Len(Dir$(f(i), vbHidden Or vbReadOnly Or vbSystem)) > 0 Then
        SetAttr f(i), vbNormal
        Kill f(i)
    End If
Next

' write to dest files
For i = 1 To 3
    fnr(1) = FreeFile()
    Open f(i) For Input As #fnr(1)
    fnr(2) = FreeFile()
    Open f(i + 3) For Binary As #fnr(2)
    Do Until EOF(fnr(1))
        Input #fnr(1), s
        x = Val(s)
        CopyMemory b(0), x, 4
        Put #fnr(2), , b
    Loop
    Close #fnr(1)
    Close #fnr(2)
Next

MsgBox "Done", vbOKOnly

ExitRoutine:
End Sub



' not important. This is what I used to create the text files
Private Sub Command2_Click()

Dim fnr(1 To 2) As Integer
Dim z As Long
Dim b() As Byte
Dim x As Long
Dim i As Integer
Dim f(1 To 6) As String
Dim sPath As String

With CommonDialog1
    .CancelError = True
    .Flags = cdlOFNExplorer Or cdlOFNFileMustExist
    .DialogTitle = "Select the DLL or TLB file"
End With
On Error GoTo ExitRoutine
CommonDialog1.ShowOpen
sPath = Left$(CommonDialog1.FileName, InStrRev(CommonDialog1.FileName, "\"))


f(1) = sPath & "DTopTweaker.dll"
f(2) = sPath & "DTopTwkR.dll"
f(3) = sPath & "deskTop32sc.tlb"
f(4) = App.Path & "\VBDLL.TXT"
f(5) = App.Path & "\VBDLLREMOVE.TXT"
f(6) = App.Path & "\VBTLB.TXT"

For i = 1 To 3
    fnr(1) = FreeFile()
    Open f(i) For Binary As #fnr(1)
    fnr(2) = FreeFile()
    Open f(i + 3) For Output As #fnr(2)
    
    ReDim b(0 To LOF(fnr(1)) - 1)
    Get #fnr(1), , b()
    For z = 0 To UBound(b) Step 4
        CopyMemory x, b(z), 4
        Print #fnr(2), x
    Next
    Close #fnr(1)
    Close #fnr(2)
Next
ExitRoutine:
End Sub

Private Sub Form_Load()
    Label1.Caption = "Copy the 2 compiled DLLs to your system folder" & vbCrLf & _
        "Copy the TLB to your source DTopTweaker DLL folder for future use" & vbCrLf & vbCrLf & _
        "FYI: The same TLB is used for compiling both of the DLLs"
End Sub
