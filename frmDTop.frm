VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDTop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desktop Icon Tweaker"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6645
   Icon            =   "frmDTop.frx":0000
   LinkTopic       =   "frmDtop"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   591
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   443
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrActiveDesktop 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   5280
      Top             =   1575
   End
   Begin VB.Timer tmrAutoHide 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   5265
      Top             =   1065
   End
   Begin VB.Timer tmrFlash 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5835
      Top             =   1590
   End
   Begin VB.Timer tmrRebuild 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5850
      Top             =   1080
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   5835
      Top             =   525
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ListView lvDtop 
      Height          =   2430
      Left            =   45
      TabIndex        =   17
      Top             =   15
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   4286
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ilDT"
      SmallIcons      =   "ilDT"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IconCaption"
         Object.Width           =   10583
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Mode"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "fColor"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "bColor"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Display"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "IconIndex"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "CurrentXY"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "RestoreXY"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CheckBox chkSort 
      Caption         =   "View Listing by Display Caption"
      Height          =   255
      Left            =   75
      TabIndex        =   0
      Top             =   2460
      Width           =   2550
   End
   Begin VB.CommandButton cmdEnd 
      Cancel          =   -1  'True
      Caption         =   "&Hide"
      Height          =   375
      Index           =   1
      Left            =   4500
      TabIndex        =   15
      Top             =   5595
      Width           =   1005
   End
   Begin VB.CommandButton cmdActivate 
      Caption         =   "&No Tweak"
      Height          =   375
      Index           =   1
      Left            =   1065
      TabIndex        =   14
      Top             =   5595
      Width           =   1005
   End
   Begin VB.PictureBox Picture1 
      Height          =   2820
      Index           =   1
      Left            =   105
      ScaleHeight     =   2760
      ScaleWidth      =   6390
      TabIndex        =   29
      Top             =   6075
      Visible         =   0   'False
      Width           =   6450
      Begin VB.ListBox lstMisc 
         Height          =   735
         ItemData        =   "frmDTop.frx":08CA
         Left            =   390
         List            =   "frmDTop.frx":08E0
         Style           =   1  'Checkbox
         TabIndex        =   41
         Top             =   1800
         Width           =   5550
      End
      Begin VB.ComboBox cboAutoHide 
         Height          =   315
         Left            =   4095
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1485
         Width           =   900
      End
      Begin VB.CheckBox chkAutoHide 
         Caption         =   "Auto-Hide Desktop when not in use for at least                       minutes"
         Height          =   255
         Left            =   435
         TabIndex        =   21
         Top             =   1530
         Width           =   5250
      End
      Begin VB.CheckBox chkNewTrans 
         Caption         =   "Background Is Transparent for Default Color Scheme"
         Height          =   210
         Left            =   435
         TabIndex        =   20
         Top             =   1140
         Width           =   4365
      End
      Begin VB.OptionButton optNewDef 
         Caption         =   "Use these colors >>"
         Height          =   225
         Index           =   0
         Left            =   420
         TabIndex        =   18
         Top             =   600
         Width           =   1995
      End
      Begin VB.OptionButton optNewDef 
         Caption         =   "Use System Default >>"
         Height          =   225
         Index           =   1
         Left            =   420
         TabIndex        =   19
         Top             =   885
         Value           =   -1  'True
         Width           =   2145
      End
      Begin VB.Line Line1 
         X1              =   450
         X2              =   5745
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label lblDefColor 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   3435
         TabIndex        =   37
         ToolTipText     =   "Not Configurable"
         Top             =   870
         Width           =   645
      End
      Begin VB.Label lblDefColor 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   4125
         TabIndex        =   36
         ToolTipText     =   "Not Configurable"
         Top             =   870
         Width           =   645
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   3435
         TabIndex        =   35
         ToolTipText     =   "Click to change"
         Top             =   615
         Width           =   645
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   3
         Left            =   4125
         TabIndex        =   34
         ToolTipText     =   "Click to change"
         Top             =   615
         Width           =   645
      End
      Begin VB.Label lblCapColor 
         Caption         =   "Background"
         Height          =   240
         Index           =   3
         Left            =   4860
         TabIndex        =   32
         ToolTipText     =   "Click to change"
         Top             =   690
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Default Color Scheme for new Desktop icons"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Top             =   330
         Width           =   3690
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Global Settings"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   3
         Left            =   285
         TabIndex        =   30
         Top             =   60
         Width           =   1185
      End
      Begin VB.Shape Shape1 
         Height          =   2550
         Index           =   1
         Left            =   105
         Top             =   150
         Width           =   6195
      End
      Begin VB.Label lblCapColor 
         Caption         =   "Text Color"
         Height          =   240
         Index           =   2
         Left            =   2565
         TabIndex        =   33
         ToolTipText     =   "Click to change"
         Top             =   690
         Width           =   810
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2820
      Index           =   0
      Left            =   75
      ScaleHeight     =   2760
      ScaleWidth      =   6390
      TabIndex        =   2
      Top             =   2715
      Width           =   6450
      Begin VB.PictureBox picIcon 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   525
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   40
         Top             =   2160
         Width           =   480
      End
      Begin VB.CommandButton cmdCaptionChg 
         Caption         =   "Reset"
         Height          =   375
         Index           =   1
         Left            =   4995
         TabIndex        =   5
         Top             =   570
         Width           =   810
      End
      Begin VB.CommandButton cmdFlash 
         Caption         =   "Flash"
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   11
         Top             =   2220
         Width           =   1695
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Make ALL Backgrounds Transparent"
         Height          =   375
         Index           =   2
         Left            =   2970
         TabIndex        =   12
         Top             =   2220
         Width           =   2850
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset to System Colors"
         Height          =   375
         Index           =   4
         Left            =   510
         TabIndex        =   8
         Top             =   1770
         Width           =   2400
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset ALL to System Colors"
         Height          =   375
         Index           =   0
         Left            =   2970
         TabIndex        =   10
         Top             =   1770
         Width           =   2850
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Apply Settings to ALL Icons"
         Height          =   375
         Index           =   1
         Left            =   2955
         TabIndex        =   9
         Top             =   1350
         Width           =   2865
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Apply to Selected"
         Height          =   375
         Index           =   3
         Left            =   510
         TabIndex        =   7
         Top             =   1350
         Width           =   2400
      End
      Begin VB.CheckBox chkTrans 
         Alignment       =   1  'Right Justify
         Caption         =   "Is Transparent?"
         Height          =   210
         Left            =   4320
         TabIndex        =   6
         Top             =   1065
         Width           =   1485
      End
      Begin VB.TextBox txtDisplay 
         Height          =   375
         Left            =   525
         MaxLength       =   260
         TabIndex        =   3
         Top             =   570
         Width           =   3840
      End
      Begin VB.CommandButton cmdCaptionChg 
         Caption         =   "Set"
         Height          =   375
         Index           =   0
         Left            =   4395
         TabIndex        =   4
         Top             =   570
         Width           =   585
      End
      Begin VB.Label lblCapColor 
         Caption         =   "Background Color"
         Height          =   240
         Index           =   1
         Left            =   2985
         TabIndex        =   27
         ToolTipText     =   "Click to change"
         Top             =   1065
         Width           =   1395
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   1380
         TabIndex        =   26
         ToolTipText     =   "Click to change"
         Top             =   1065
         Width           =   645
      End
      Begin VB.Label lblColor 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   2235
         TabIndex        =   25
         ToolTipText     =   "Click to change"
         Top             =   1065
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "Displayed Icon Caption (Changes are not Permanent)"
         Height          =   210
         Index           =   2
         Left            =   540
         TabIndex        =   24
         Top             =   360
         Width           =   4470
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Per-Icon Colors and Captions"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   0
         Left            =   285
         TabIndex        =   23
         Top             =   60
         Width           =   2235
      End
      Begin VB.Shape Shape1 
         Height          =   2550
         Index           =   0
         Left            =   105
         Top             =   150
         Width           =   6195
      End
      Begin VB.Label lblCapColor 
         Caption         =   "Text Color"
         Height          =   240
         Index           =   0
         Left            =   510
         TabIndex        =   28
         ToolTipText     =   "Click to change"
         Top             =   1065
         Width           =   930
      End
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "Terminate"
      Height          =   375
      Index           =   0
      Left            =   5520
      TabIndex        =   16
      Top             =   5595
      Width           =   1005
   End
   Begin VB.CommandButton cmdActivate 
      Caption         =   "Twea&k"
      Height          =   375
      Index           =   0
      Left            =   45
      TabIndex        =   13
      Top             =   5595
      Width           =   1005
   End
   Begin VB.CheckBox chkView 
      Caption         =   "Show Global Settings"
      Height          =   255
      Left            =   2655
      TabIndex        =   1
      Top             =   2460
      Width           =   1815
   End
   Begin VB.Label lblActiveDT 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Cannot be activated while in Active Desktop Mode"
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   2145
      TabIndex        =   39
      Top             =   5565
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Caption         =   "Found xxx Icons"
      Height          =   210
      Left            =   4650
      TabIndex        =   38
      Top             =   2475
      Width           =   1755
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Desktop"
      Index           =   1
      Begin VB.Menu mnuDesktop 
         Caption         =   "Twea&k Icons"
         Index           =   0
      End
      Begin VB.Menu mnuDesktop 
         Caption         =   "Do &Not Tweak Icons"
         Index           =   1
      End
      Begin VB.Menu mnuDesktop 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuDesktop 
         Caption         =   "&Hide Desktop"
         Index           =   3
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuDesktop 
         Caption         =   "&Display Desktop"
         Index           =   4
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuDesktop 
         Caption         =   "&Refresh Desktop"
         Index           =   5
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuDesktop 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuDesktop 
         Caption         =   "Save Icon &Positions"
         Index           =   7
      End
      Begin VB.Menu mnuDesktop 
         Caption         =   "&Restore Icon Postions"
         Index           =   8
      End
      Begin VB.Menu mnuDesktop 
         Caption         =   "&Un-Restore Icon Positions"
         Index           =   9
      End
      Begin VB.Menu mnuDesktop 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuDesktop 
         Caption         =   "Reset all &Icons to System Defaults"
         Index           =   11
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Global"
      Index           =   2
      Begin VB.Menu mnuGlobal 
         Caption         =   "&AutoHide Desktop"
         Index           =   0
      End
      Begin VB.Menu mnuGlobal 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuGlobal 
         Caption         =   "Save icons when application terminates"
         Index           =   2
      End
      Begin VB.Menu mnuGlobal 
         Caption         =   "Restore icons when applications starts"
         Index           =   3
      End
      Begin VB.Menu mnuGlobal 
         Caption         =   "Restore icoins when desktop changes resolution"
         Index           =   4
      End
      Begin VB.Menu mnuGlobal 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuGlobal 
         Caption         =   "Hide balloon tip when icons are hidden"
         Index           =   6
      End
      Begin VB.Menu mnuGlobal 
         Caption         =   "Do not show in the system tray"
         Index           =   7
      End
      Begin VB.Menu mnuGlobal 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuGlobal 
         Caption         =   "More..."
         Index           =   9
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuPop 
         Caption         =   "&Open"
         Index           =   1
      End
      Begin VB.Menu mnuPop 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuPop 
         Caption         =   "&Show Desktop"
         Index           =   3
      End
      Begin VB.Menu mnuPop 
         Caption         =   "&Hide Desktop"
         Index           =   4
      End
      Begin VB.Menu mnuPop 
         Caption         =   "Auto-Hide &Desktop"
         Index           =   5
      End
      Begin VB.Menu mnuPop 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuPop 
         Caption         =   "Sa&ve Icon Positions"
         Index           =   7
      End
      Begin VB.Menu mnuPop 
         Caption         =   "&Restore Icon Positions"
         Index           =   8
      End
      Begin VB.Menu mnuPop 
         Caption         =   "&Un-Restore Icon Positions"
         Index           =   9
      End
      Begin VB.Menu mnuPop 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuPop 
         Caption         =   "Twea&king"
         Index           =   11
      End
      Begin VB.Menu mnuPop 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuPop 
         Caption         =   "&Terminate"
         Index           =   13
      End
   End
End
Attribute VB_Name = "frmDTop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' HYBRID DTopTweaker DLL VERSION REQUIRED: v2.1 (injects & subclasses)
' HYBRID DTopTwkR DLL VERSION REQUIRED: v1.0    (removes injection)
' This version of the application: v2.1
' The HowToSetUp.RTF file has been updated too & tells you how to compile needed DLLs

' Last Updated: 18-20 May 05
'   - InstallDLL would fail if optional version was passed & DLL never existed in target folder
'   - Added prevInstance check and option to show prevInstance
'   - Added a few more options & tweaked GUI
'   - Overhauled RebuildListView routine due to NT's lack of saving icon positions
'       until it closes or other specific actions occur. See Limitations below.
' Updated: 17 May 05 PM.
'   - Provided more robust InstallDLL function
'   - Fixed desktop crash potential when sending items to/from the Recycle Bin
' Again later than night
'   - located reliable command to toggle autoarrange in both 9x and NT (Vista?)
'   - fixed losing saved icon positions when app started
' Next update? Depends on bug reports.

' PREFACE:
' ~~~~~~~~
' This project is designed to highlight the flexibility of using VB6 for global
'   hooking and DLL injection.  The DLL being used to subclass the desktop was
'   created using just VB. A few hacks were made to convert the DLL from an
'   Active-X (can't be used for global hooking/subclassing) to a standard DLL.
'   These hacks are combined into a nice single project by DreamVB. Using that
'   project and a custom .DEF file, the DLL is converted. This leads to a major
'   issue because the DLL will crash any process it is injected into if that
'   process isn't already running the VB6 runtime DLLs. Therefore, a workaround
'   is to create and compile a .TLB file which allows API calls, structures, and
'   other objects to be used that would otherwise crash the injected process.
'   Though this proves global subclassing/hooking is possible with VB, it also
'   proves it is a huge pain in the a$$ (PITA). If you decide to play with DLLs
'   created this way, you will soon find that out too. Debugging is nearly impossible.
'   So using an existing C++ or Delphi DLL can be used to subclass/hook, what are the
'   advantages of using VB? Not much except that you can change & modify the code at
'   will without having to know the other languages. Of course, this means you must
'   be very proficient with API usage.

' Because PSC does not allow uploading compiled files, the TLB and DLL will have
'   to be created by you. To assist I have included a step by step procedure on
'   how to create the needed DLL and TLB. Also with the download are two text files
'   that can be converted to binary using the attached TxtToBin project.
'   Those text files when converted will be the DLL and TLB should you want to
'   avoid compiling your own DLL and TLB files. Either way, the source for the
'   DLL exists as a separate project and the TLB source is provided as a .ODL file.
'   And if you chose to convert the text to binary, the step-by-step instructions
'   will help should you want to modify the DLL or create your own.

' Last but not least regarding the standard VB-created DLLs. The DLLs must exist in
' the system folder in order to subclass the desktop.

' WARNINGS:
' ~~~~~~~~~
' 1. The DLL has been tested quite a bit on Windows 2000 & 98. It has not been tested
'   on any other systems by me. Has been tested a bit on XP by other kind souls.
' 2. Because we are subclassing the Desktop, a logic error in the DLL could crash
'   the desktop. This isn't a big deal for NT systems but may be for Win 9x systems.
' 3. Even though the DLL is compiled, suggest using extreme caution editing this
'   project during runtime. I have experienced crashes that only occurred while
'   modifying code while the project was running but "paused".

' DESCRIPTION:
' ~~~~~~~~~~~~
' This project will allow you to customize each icon individually to include the
'   caption foreground and background colors along with the caption itself.
' 1. Other projects have been around for awhile that simply change the foreground and
'   background colors. But those colors applied to all icons because the code did not
'   subclass the desktop to intercept custom-drawn messages which are actually provided
'   by the desktop itself (similar to ownerdrawn menus).
' 2. This project allows a user to modify the following for each icon...
'   a. The text forecolor (any valid color)
'   b. The text background color (any valid color, including Transparent)
'   c. The text caption can be set to anything. This means you can rename the Recycle Bin too.
'   Theoretically, 100% of the drawing could be accomplished to provide some neat effects
'   like alphablending and translucent icons; but that is for future revisions possibly.
' 3. Other minor functions have been added and others may be included in future revisions
'   a. Automatically toggle the desktop visibility
'   b. Save and restore icon positions
'   c. Access this running project from the system tray or right clicking on desktop
' 4. This project uses an INI file and does not modify registry or system settings

' LIMITATIONS:
' ~~~~~~~~~~~~
' 1. Once the DLL is injected into the desktop, the following needs to be kept in mind:
'   a. The DLL cannot be removed through code. Ideas welcomed. In order to keep the DLL
'       in the process once it was injected, the DLL had to create another reference to
'       itself. But I cannot dereference the DLL from within the same DLL without
'       crashing the process. Other thoughts include creating an API window that could
'       be used to dereference the DLL or creating another standard DLL that would
'       be injected into the process for the sole purpose of removing the 1st DLL.
'       The 2nd DLL would be dereferenced naturally since it would not ref itself.
'   b. To remove the DLL manually:
'       (1) In NT based systems.
'           Close this project, Kill Explorer.Exe from task manager, & restart Explorer.Exe
'       (2) From 9x systems: Log off & log on
' ==========================================================================================
' ^^ ***** UPDATED/Fixed: The second DLL in this project is used to un-inject
' ==========================================================================================
' 2. Positions (NT) and colors/captions (9x & NT) are temporary. When you select to
'   sort icons via AutoArrange or by date/size/etc, and this app later relocates the
'   icons via its Restore option, the positions are not cached in NT. In this case,
'   the following actions will save the positions within NT: physically moving an icon,
'   restoring an icon from the Recycle Bin, and a couple of other less used functions.
'   Therefore, in this case, if icons are moved around from within this program and if
'   the desktop is Refreshed, the icons may revert back to their last NT-saved positions
'   (AutoArrange, by date/size/etc). When and if this happens, simply select this app's
'   menu option to Restore Icons. This "bug" in desktop icon management programs,
'   appears in both professional (i.e., Layout.dll) & amateur apps (like mine)
' 3. On XP, use of icon text drop shadows overrides any subclassing. Although captions
'   can still be changed, coloring text is not possible. Active Desktop in any version
'   likewise prevents effects & tests in the project warn of that setting.
' 4. I don't have XP, but am expecting the Grouping and Align to Grid options will
'    prevent successful restore/save/un-restore functions
' 5. The desktop listview does NOT allow you to tag a listview item to uniquely
'   identify it. This is not good because a user can change the caption of a desktop
'   icon or the icon can be rearranged via sorting and/or auto-arrange. Therefore,
'   this application attempts to use the icon name as the unique identifier. But this
'   is far from perfect. Even though I can determine if an individual item is being
'   edited, the desktop listview will Delete ALL items, and then re-insert anew,
'   when one of these actions occur....
'   a. User presses F5 on the desktop
'   b. User selects Refresh from the desktop's context menu
'   c. User chooses to toggle "Hide known file type extensions" option in Explorer
'   d. User adds/removes a system folder (My Desktop, etc) from the desktop
'   Whenever the desktop removes all items & then re-adds them again, if the file name
'   changed between those two actions, we will never know.
'   This is usually affected by toggling known extensions.

' 6. Because you can use the same caption for 2 or more of your icons, the project has
'   no real way of knowing which one is which in these cases. For example, there is
'   nothing preventing you from naming Network Neighborhood to My Computer which would
'   result in 2 icons captioned as My Computer. Additionally, if you had 3 files whose
'   captions were Test.Txt, Test.Zip, Test.Bak and you were hiding known file extensions,
'   then all three files will have the same caption of Test
' Thoughts for workarounds were to also track the X,Y coords of the icon position to
'   help uniquely identify it or to use the icon's index. These won't work well
'   since those values may be dynamic or changed without this project knowing about it.
' Also thought of using the lParam value of the sysListView's LVITEM structure which
' could be used. But that value is also dynamically changed by the desktop.

' All of the above comments basically say this: Don't use duplicate icon captions and
'   if you do, there is no guarantee the custom settings will be preserved. One
'   exception to this rule is the blank caption. Since it doesn't matter which
'   custom colors exist (i.e., no caption to display), all blank captions are
'   handled appropriately.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

' THE ABOVE SHOULD BE READ SO YOU UNDERSTAND THE RISKS AND LIMITATIONS.


' Paul Caton's subclassing techniques used to subclass this project. This project is
'  subclassed to receive Explorer restart notifications, system tray notifications; and
' other minor system messages (change of desktop colors for example); otherwise all
' of the subclassing could be replaced by the DLL using SetWindowText to a textbox
' to trigger this app to look at the mapped memory files. The sysTray notifications
' could be handled without subclassing too.

'*************************************************************************************************
'* fSelfSub - self-subclassing form template
'*
'* Note: it's a bad idea to sc_Terminate from a form's terminate event. The window will have been
'*       destroyed and so prevent the thunk from releasing its allocated memory. Instead, use the
'*       form's unload event to sc_Terminate when closing a form.
'*
'* Paul_Caton@hotmail.com
'* Copyright free, use and abuse as you see fit.
'*
'* v1.0 Re-write of the SelfSub/WinSubHook-2 submission to Planet Source Code............ 20060322
'* v1.1 VirtualAlloc memory to prevent Data Execution Prevention faults on Win64......... 20060324
'* v1.2 Thunk redesigned to handle unsubclassing and memory release...................... 20060325
'* v1.3 Data array scrapped in favour of property accessors.............................. 20060405
'* v1.4 Optional IDE protection added
'*      User-defined callback parameter added
'*      All user routines that pass in a hWnd get additional validation
'*      End removed from zError.......................................................... 20060411
'* v1.5 Added nOrdinal parameter to sc_Subclass
'*      Switched machine-code array from Currency to Long................................ 20060412
'* v1.6 Added an optional callback target object
'*      Added an IsBadCodePtr on the callback address in the thunk prior to callback..... 20060413
'*************************************************************************************************
'-Selfsub declarations----------------------------------------------------------------------------
Private Enum eMsgWhen                                                       'When to callback
  MSG_BEFORE = 1                                                            'Callback before the original WndProc
  MSG_AFTER = 2                                                             'Callback after the original WndProc
  MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                                'Callback before and after the original WndProc
End Enum

Private Const ALL_MESSAGES  As Long = -1                                    'All messages callback
Private Const MSG_ENTRIES   As Long = 32                                    'Number of msg table entries
Private Const WNDPROC_OFF   As Long = &H38                                  'Thunk offset to the WndProc execution address
Private Const GWL_WNDPROC   As Long = -4                                    'SetWindowsLong WndProc index
Private Const IDX_SHUTDOWN  As Long = 1                                     'Thunk data index of the shutdown flag
Private Const IDX_HWND      As Long = 2                                     'Thunk data index of the subclassed hWnd
Private Const IDX_WNDPROC   As Long = 9                                     'Thunk data index of the original WndProc
Private Const IDX_BTABLE    As Long = 11                                    'Thunk data index of the Before table
Private Const IDX_ATABLE    As Long = 12                                    'Thunk data index of the After table
Private Const IDX_PARM_USER As Long = 13                                    'Thunk data index of the User-defined callback parameter data index

Private z_ScMem             As Long                                         'Thunk base address
Private z_Sc(64)            As Long                                         'Thunk machine-code initialised here
Private z_Funk              As Collection                                   'hWnd/thunk-address collection

Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal length As Long)
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Above required for Paul Caton's routines

' Custom DLLs. These DLLs were created & compiled with only VB. See comments at top.
Private Declare Function HookDesktop Lib "DTopTweaker.dll" _
        (ByVal hWndParent As Long, _
        ByRef mapFileTx As String, _
        ByRef mapFileRx As String, _
        ByRef PrivateMsg As Long) As Long
Private Declare Function UnhookDesktop Lib "DTopTwkR.dll" () As Long
'``````````````````````````````````````````````````

' Required TYPES for retrieving desktop icon information
Private Type LVITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As Long
    cchTextMax As Long
    iImage As Long
    lParam As Long
End Type

Private Enum eAutoHideOptions
    ea_Disable = 0
    ea_Enable = 1
    ea_SetTimer = 2
    ea_ChkTimer = 3
    ea_Reset = 4
End Enum
Private Enum eActivateMode
    ea_Initialize = 0
    ea_Inactivate = 1
    ea_ReInitialize = 2
End Enum
Private Enum eOptions
    eo_NoBalloon = 0
    eo_SaveExit = 1
    eo_RestoreRun = 2
    eo_RestoreRes = 3
    eo_NoSysTray = 4
    eo_DblClkRestore = 5
End Enum

' local variables
Private Const m_maxChars As Long = 260
' ^^ a listview item has unlimited caption length; but will only display 260 max
'    we will set a limit of 260 characters
Private lvCaption(0 To m_maxChars) As Byte
' ^^ general purpose byte array to convert ANSI>Unicode captions
Private m_Attr(0 To 1) As Long ' values passed from custom DLL
'^^ 0=forecolor, 1=backcolor
Private m_Rebuilding As Boolean
' ^^ flag indicating not to process DLL requests while host listview is being built
Private m_CustomNewColors As Long
' ^^ new desktop icons use custom color scheme or system scheme, with/without transparency
Private m_dtAutoHide As Date
' ^^ target date/time to hide desktop (used with AutoHide function)
Private hListView As Long   ' handle to the desktop's listview
Private hShell As Long      ' handle to listview's parent object
Private hFileRx As Long       ' file handle for mapped memory
Private hMapRx As Long        ' mapped handle used for receipt of data only
Private hFileTx As Long       ' file handle for mapped memory
Private hMapTx As Long        ' mapped handle used for transmission of data only

Private wm_traynotify As Long   ' system tray notifications
Private wm_expcrash As Long    ' notification of explorer crash
Private wm_private As Long    ' GUID used for sending DLL messages
'^ used with SendMessage or PostMessage only, must be the wMsg parameter

Private cTray As clsSysTray ' system tray icon notifications
Private m_MasterRes As Long ' resolution when app started
Private m_CurrentRes As Long
Private m_CanRestore As Boolean     ' can restore? only if previous saved
Private m_CanUnRestore As Boolean   ' was restore activated?
Private m_AutoArrageSave As Boolean
Private m_AutoArrageRestore As Boolean

' //////////////// FORM AND FORM CONTROL ROUTINES \\\\\\\\\\\\\\\\\\\\
Private Sub Form_Load()

MsgBox "If you have run previous versions of this project, your custom settings may " & _
    "be lost. This is because of a slightly different way of storing/reading the " & _
    "background transparency setting from the INI file." & vbCrLf & vbCrLf & _
    "Remove this MsgBox from your Form_Load routine so you don't see it every time", vbExclamation + vbOKOnly, "FYI"
    
    
' get unique message for use in system tray notifications & to show prevInstance
wm_traynotify = RegisterWindowMessage("38A51467B3844d799B7AF3BDFD2CBD23")

If App.PrevInstance = True Then
    If MsgBox("This application is already running in your system tray, " & vbCrLf & _
        "would you like to view it now", vbYesNo + vbQuestion) = vbYes Then
        ' if it's already running, then it's monitoring its systray notifications.
        ' we will send top level windows our custom systray message. One of those
        ' top-level windows is our prevInstance and it will then show itself
        SendNotifyMessage HWND_BROADCAST, wm_traynotify, SW_SHOW, WM_LBUTTONDOWN
    End If
    Unload Me
    Exit Sub
End If
    
    
' RECOMMENDATIONS:
' ~~~~~~~~~~~~~~~~
' Should you want to compile and use this app in realtime, recommend the following:
' 1. Unrem the line ShowWindow Me.hwnd, SW_SHOWMINIMIZED below so project starts in systray
' 2. Unrem the line in InitializeInstance: sc_AddMsg Me.hwnd, wm_expcrash, MSG_AFTER
' 3. Create and add a RES file to this project
' 4. Upload the compiled DLLs as custom resources #101 and #102 (#102=DTopTwkR.dll)
' 5. Un-Rem the InstallDLL lines below
' 6. Change the resource identifiers in those lines if needed
' 7. Add modInstallDLL.bas to this project or copy & paste its code to an existing module
' 8. Test it (Ctrl+F5) and then compile it
' 9. Add it to your startup folder to have it run when you start Windows

    Const SW_SHOWMINIMIZED As Long = 2
    Dim I As Long

'    I = InstallDLL("DTopTweaker", 101)' or> I=InstallDLL("DTopTweaker",101,"2.1.0.0")
'    I = I Or InstallDLL("DTopTwkR", 102) ' or> I=InstallDLL("DTopTwkR",102,"1.0.0.0")
'    ' note. Adding optional version number in above functions is more efficient
'    If I <> 0 Then ' a dll didn't get installed
'        ' ^^ (is resfile attached, are res numbers above correct, does user have permission to write to system folder?)
'        MsgBox "Cannot load the required DLL(s) into the system folder", vbCritical + vbOKOnly, "Aborting"
'        On Error Resume Next
'        Unload Me
'        Exit Sub
'    End If
    
    For I = 1 To Picture1.UBound    ' line up picture boxes
        Picture1(I).Move Picture1(0).Left, Picture1(0).Top
        Picture1(I).Visible = False
    Next
    Picture1(0).ZOrder
    
    ' build the autoHide timeout combobox
    For I = 0 To 10: cboAutoHide.AddItem I: Next
    For I = 15 To 60 Step 5: cboAutoHide.AddItem I: Next
    
    GetDeskTopIconColors    ' cache current icon text/back colors
    If IsActiveDesktop = False Then ShowWindow hShell, 0
    DoEvents
    InitializeListView  ' get all icon names from desktop
    If lvDtop.ListItems.Count = 0 Then
        MsgBox "Found no icons on the desktop to manage. Closing application", vbExclamation + vbOKOnly
        ShowWindow hShell, SW_SHOW
        Unload Me
        Exit Sub
    End If
    
    For I = 0 To 2 Step 2
        lblColor(I).BackColor = lblDefColor(0).BackColor
        lblColor(I + 1).BackColor = lblDefColor(1).BackColor
    Next
    Set cTray = New clsSysTray
    With cTray
        .InitializeTray Me.Icon, "Desktop Icon Tweaker"
        If lstMisc.Selected(eo_NoSysTray) = False Then .BeginTrayNotifications Me.hwnd, Me.hwnd, wm_traynotify
    End With
    If cTray.isBalloonCapable = False Then
        lstMisc.List(eo_NoBalloon) = lstMisc.List(eo_NoBalloon) & " {not applicable}"
        lstMisc.Selected(eo_NoBalloon) = False
    End If
'    Call RemoveCloseButton(Me.hwnd)
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2, 6690, 6675
'*** Unrem below line to have app start out in the system tray
'    ShowWindow Me.hwnd, SW_SHOWMINIMIZED
    ' ^^ don't use Me.WindowState=vbMinimized here. May prevent app from coming to
    '       foreground when the app is first shown out of the system tray.

    On Error GoTo FailedLoad
    LoadSettings        ' load any cached icon captions & compare
    InitializeInstance ea_Initialize   ' now subclass our form and inject the DLL
    Call chkSort_Click      ' call routine to size visible column to listivew space
    tmrRebuild.Interval = 0
    tmrRebuild.Enabled = True
    If lstMisc.Selected(eo_RestoreRun) = True Then RestoreIcons True
    ShowWindow hShell, SW_SHOW
    
Exit Sub
   
FailedLoad:
    ShowWindow hShell, SW_SHOW
    MsgBox "Critical errors reported", vbExclamation + vbOKOnly, "Failed to Initialize"
    Err.Clear
    Unload Me
End Sub

Private Sub Form_Resize()
    ' in case some outside application minimizes us,
    ' otherwise is handled to animate towards system tray
    If Me.WindowState = vbMinimized Then ShowWindow Me.hwnd, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    ' note: should explorer crash when exiting, please let me know.
    ' A patch would be to delay unloading for a few seconds to ensure any messages
    ' still being processed by the injected DLL are completely flushed.
    
    If z_Funk Is Nothing Then Exit Sub
    ' ^^ haven't subclassed yet? Should only happen if a 2nd instance or missing DLLs
    
    ' terminate DLL subclasser and this form's subclassing
    Call CleanUp
    Set cTray = Nothing ' remove systray icon
    ' refresh the desktop & save settings
    Me.Visible = False
    sc_Terminate ' stop subclassing
    RefreshDeskTop False, False
    SaveSettings
    
End Sub

Private Sub cmdActivate_Click(Index As Integer)
    If Index = 0 Then
        If wm_private = 0 Then
            InitializeInstance ea_Initialize
        Else
            InitializeInstance ea_ReInitialize
        End If
    Else
        InitializeInstance ea_Inactivate
    End If
    If Index = 0 Then
        If lstMisc.Selected(eo_RestoreRun) = True Then
            StoreIconPositions
            RestoreIcons True
        End If
    End If
    RefreshDeskTop False, False

End Sub

Private Sub cmdCaptionChg_Click(Index As Integer)

    ' change/reset the display caption of a specific icon
    Dim x As Long
    Dim uItem As LVITEM
    Dim tMap As Long, tMapHndl As Long
    Dim lvMap As Long, lvMapHndl As Long
    
    If DeskTopMsgBox = True Then
        MsgBox "The desktop has a message box open at the moment." & vbCrLf & _
            "Close it first, then click the Set/Reset button again.", vbInformation + vbOKOnly
        Exit Sub
    End If
    
    x = Val(Mid$(lvDtop.SelectedItem.Key, 2)) ' sysListView icon index
    
    ' create 2 file maps (1 for the LVITEM structure, and one for the caption)
    lvMap = CreateMapping(Len(uItem), lvMapHndl, , hListView)
    If Not lvMap = 0 Then
        tMap = CreateMapping(1024, tMapHndl, , hListView)
    End If
    If Not tMap = 0 Then
    
        If Index = 1 Then txtDisplay.Text = lvDtop.SelectedItem.Text
        ' ^^ resetting to actual caption
        With uItem
            .cchTextMax = Len(txtDisplay.Text)
            .mask = LVIF_TEXT
            .pszText = tMap
            .iItem = x
        End With
        lvDtop.SelectedItem.Tag = txtDisplay.Text
        ' ^^ save adjusted caption to Tag property and set the new caption
        Call SetItemCaption(txtDisplay.Text, x, lvMapHndl, lvMap, tMapHndl, tMap, VarPtr(uItem), Len(uItem))
        
        FreeMapping tMapHndl, tMap
        FreeMapping lvMapHndl, lvMap
        
        RefreshDeskTop False, False, x
        Call lvDtop_ItemClick(lvDtop.SelectedItem)
    End If
    
End Sub

Private Sub cmdEnd_Click(Index As Integer)
    If Index = 0 Then   ' terminate
        Unload Me
    Else                ' hide
        Call cTray.MinimizeAnimated(Me.hwnd)
    End If
End Sub

Private Sub cmdFlash_Click(Index As Integer)

    ' routine flashes an icon's caption on screen.
    ' Meant to be an aid in finding the icon for the selected listview item

    If DeskTopMsgBox = True Then
        MsgBox "The desktop has a message box open at the moment." & vbCrLf & _
            "Close it first, then click the Set/Reset button again.", vbInformation + vbOKOnly
        Exit Sub
    End If
    
    Dim xItem As ListItem
    Dim I As Integer, x As Long
    
    ' this routine will be called 7 times with a 50 ms gap
    tmrFlash.Tag = Index + 1
    If Index = 0 Then               ' first called; set up the flash routine
        SetAutoHideMode ea_Disable  ' ensure desktop visible
        Me.Enabled = False      ' disable this project
        lvDtop.Sorted = False   ' remove sorting for now
        ' add a temp item to listview & copy the selected item data to it
        Set xItem = lvDtop.ListItems.Add(lvDtop.ListItems.Count + 1, "Flasher", "")
        For I = 1 To lvDtop.ColumnHeaders.Count - 1
            xItem.SubItems(I) = lvDtop.SelectedItem.SubItems(I)
        Next
        lvDtop.SelectedItem.SubItems(1) = 1 ' flag for user-defined fore/back colors
        If lvDtop.SelectedItem.Tag = "" Then
            ' when the item doesn't have a caption to display, we will
            ' reset it to its original caption for flashing purposes only
            lvDtop.SelectedItem.Tag = lvDtop.SelectedItem.Text
            txtDisplay.Text = lvDtop.SelectedItem.Tag
            Call cmdCaptionChg_Click(0)
        Else
            xItem.Tag = lvDtop.SelectedItem.Tag
        End If
    End If
    
    With lvDtop.SelectedItem
    
        ' get the sysListView icon index to flash
        x = Val(Mid$(.Key, 2))
        
        ' alternate coloring for flash effect
        If Index Mod 2 = 0 Then
            .SubItems(2) = lblDefColor(0).BackColor
            .SubItems(3) = lblDefColor(1).BackColor
        Else
            ' copy orignal settings
            .SubItems(2) = lblDefColor(1).BackColor
            .SubItems(3) = lblDefColor(0).BackColor
        End If
    End With
        
    RefreshDeskTop False, True, x        ' redraw
    If Index < 7 Then
        tmrFlash.Enabled = True          ' continue until 8th loop
    Else
        ' put things back the way they were
        Set xItem = lvDtop.ListItems("Flasher")
        For I = 1 To lvDtop.ColumnHeaders.Count - 1
            lvDtop.SelectedItem.SubItems(I) = xItem.SubItems(I)
        Next
        If lvDtop.SelectedItem.Tag <> xItem.Tag Then
            lvDtop.SelectedItem.Tag = xItem.Tag
            txtDisplay.Text = lvDtop.SelectedItem.Tag
            Call cmdCaptionChg_Click(0)
        End If
        lvDtop.ListItems.Remove "Flasher"
        lvDtop.Sorted = True
        lvDtop.SelectedItem.EnsureVisible
        Me.Enabled = True
        SetAutoHideMode chkAutoHide.Value
    End If

End Sub

Private Sub chkNewTrans_Click()
    ' for new desktop icons, option to have transparency with
    ' the system colors or custom colors
    m_CustomNewColors = chkNewTrans.Value * 2
    If optNewDef(0) = True Then m_CustomNewColors = m_CustomNewColors Or 1
    
End Sub

Private Sub cmdReset_Click(Index As Integer)
    
    ' changes color settings for a single icon, range of icons, or all icons

    If lvDtop.ListItems.Count = 0 Then Exit Sub
    
    If DeskTopMsgBox = True Then
        MsgBox "The desktop has a message box open at the moment." & vbCrLf & _
            "Close it first, then try again.", vbInformation + vbOKOnly
        Exit Sub
    End If
    
    Dim I As Long, cMode As Long
    Select Case Index
    Case 0 ' reset back to defaults
        For I = 1 To lvDtop.ListItems.Count
            lvDtop.ListItems(I).SubItems(1) = 0
        Next
        RefreshDeskTop False, False
        
    Case 1 ' change all to current setting
        If chkTrans.Value = 1 Then cMode = 3 Else cMode = 1
        For I = 1 To lvDtop.ListItems.Count
            lvDtop.ListItems(I).SubItems(1) = cMode
            lvDtop.ListItems(I).SubItems(2) = lblColor(0).BackColor
            lvDtop.ListItems(I).SubItems(3) = lblColor(1).BackColor
        Next
        RefreshDeskTop False, False
        
    Case 2 ' set all to transparent backgrounds
        For I = 1 To lvDtop.ListItems.Count
            cMode = Val(lvDtop.ListItems(I).SubItems(1))
            lvDtop.ListItems(I).SubItems(1) = cMode Or 2
        Next
        RefreshDeskTop False, False
        
    Case 3, 4 ' apply settings or reset settings per selected item
        If chkTrans.Value = 1 Then cMode = 3 Else cMode = 1
        For I = 1 To lvDtop.ListItems.Count
            If lvDtop.ListItems(I).Selected = True Then
                If Index = 3 Then ' apply else reset
                    lvDtop.ListItems(I).SubItems(2) = lblColor(0).BackColor
                    lvDtop.ListItems(I).SubItems(3) = lblColor(1).BackColor
                    lvDtop.ListItems(I).SubItems(1) = cMode
                Else
                    lvDtop.ListItems(I).SubItems(1) = 0 ' system colors flag
                End If
                RefreshDeskTop False, False, Val(Mid$(lvDtop.ListItems(I).Key, 2))
            End If
        Next
    End Select
    Call lvDtop_ItemClick(lvDtop.SelectedItem)
End Sub

Private Sub chkAutoHide_Click()
    ' call function to setup or remove autohide feature
    If chkAutoHide.Value = ea_Enable And cboAutoHide.ListIndex = 0 Then
        cboAutoHide.ListIndex = 1
    Else
        SetAutoHideMode chkAutoHide.Value
    End If
    mnuGlobal(0).Checked = (chkAutoHide.Value = ea_Enable)
End Sub

Private Sub chkSort_Click()

    If lvDtop.ListItems.Count = 0 Then Exit Sub
    
    Dim lvWidth As Long
    Dim cRect As RECT
    
    ' just to make it pretty, stretch the only visible column fully across listview
    GetClientRect lvDtop.hwnd, cRect
    lvWidth = ScaleX(cRect.Right, vbPixels, Me.ScaleMode)
    
    lvDtop.Visible = False ' prevent flicker as much as possible
    lvDtop.Sorted = False
    If chkSort.Value = 0 Then
        lvDtop.ColumnHeaders(1).Width = lvWidth
        lvDtop.ColumnHeaders(5).Width = 0
        lvDtop.SortKey = 0
    Else
        Dim I As Integer
        For I = 1 To lvDtop.ListItems.Count
            With lvDtop.ListItems(I)
                If .Tag = "" Then
                    ' add extra info for blank captions
                    .SubItems(4) = "[blank] " & .Text
                Else
                    .SubItems(4) = .Tag
                End If
            End With
        Next
        lvDtop.SortKey = 4 ' sort on display name, not actual name
        lvDtop.ColumnHeaders(5).Width = lvWidth ' show display column
        lvDtop.ColumnHeaders(1).Width = 0       ' hide actual caption
    End If
    lvDtop.Sorted = True
    lvDtop.SelectedItem.EnsureVisible
    lvDtop.Visible = True
    
    If chkView.Value = 1 Then chkView.Value = 0 'switch back to non-global settings
    
End Sub

Private Sub chkView_Click()
    ' toggle between Per-Icon & Global settings
    Picture1(0).Visible = (chkView.Value = 0)
    Picture1(1).Visible = (chkView.Value = 1)
    mnuGlobal(9).Enabled = (chkView.Value = 0)
End Sub

Private Sub lblCapColor_Click(Index As Integer)
    ' option to choose icon text and/or background colors
    With dlgColor
        .Color = lblColor(Index).BackColor
        .Flags = cdlCCRGBInit
    End With
    On Error GoTo ExitRoutine
    dlgColor.ShowColor
    lblColor(Index).BackColor = dlgColor.Color
    If Index = 1 Then chkTrans.Value = 0 ' if bkg color selected, uncheck transparency

ExitRoutine:
End Sub

Private Sub lblColor_Click(Index As Integer)
    ' same as lblColor_Click above
    Call lblCapColor_Click(Index)
End Sub

Private Sub lstMisc_ItemCheck(Item As Integer)
' miscelaneous options, sanity checks
    Select Case Item
    Case eo_NoBalloon
        If Not cTray Is Nothing Then
            ' don't allow to uncheck if running Win9x which isn't balloon capable
            If Not cTray.isBalloonCapable Then lstMisc.Selected(Item) = True
        Else
            lstMisc.Selected(Item) = True
        End If
    Case eo_NoSysTray
        If Not cTray Is Nothing Then
            ' toggle system tray if needed
            If lstMisc.Selected(Item) = False Then
                cTray.BeginTrayNotifications Me.hwnd, Me.hwnd, wm_traynotify
            Else
                cTray.RemoveTrayIcon
            End If
        End If
    End Select
    Call mnuMain_Click(-1) ' reset check marks on menus
End Sub

Private Sub lvDtop_ItemClick(ByVal Item As MSComctlLib.ListItem)

' column values:
' (1) Text Property is the icon caption reported by desktop
' (2) Customize mode: 0=none, 1=custom, 2=transparent bkg
' (3) Caption forecolor
' (4) Caption backcolor
' (5) User-adjusted icon caption
' (6) Current X,Y coords of icon
' (7) Last saved coords of icon

    ' show custom settings for the listview item that was clicked
    If Item Is Nothing Then Exit Sub
    If lvDtop.SelectedItem Is Nothing Then Exit Sub

    Dim tColor As Long, bColor As Long
    Dim lMsg As Long, lIndex As Long, ico32 As Long, hIcon As Long
    
    GetSetColors False, Val(Mid$(Item.Key, 2)), tColor, bColor, True
    lblColor(0).BackColor = tColor
    lblColor(1).BackColor = bColor
    txtDisplay.Text = Item.Tag
    cmdCaptionChg(1).Enabled = (StrComp(Item.Tag, Item.Text, vbBinaryCompare) <> 0)
    
    ' thanx to Herman Liu for brainstorming an easy way to get the icon for the selected item
    picIcon.Cls
    lMsg = MakeDWord(WM_GETICON, 0) ' 0=large icon,1=small icon
    lIndex = Val(lvDtop.SelectedItem.SubItems(5)) ' icon image index
    ico32 = SendMessage(hShell, wm_private, lMsg, ByVal lIndex)
    If ico32 <> 1 Then ' custom dll returns 1 if it failed to get an icon
        hIcon = CopyIcon(ico32) ' make copy first; experienced some crashes otherwise
        DrawIconEx picIcon.hdc, 0, 0, hIcon, 32, 32, 0, 0, &H3
        DestroyIcon hIcon
    End If
    picIcon.Refresh
    
End Sub

Private Sub mnuDesktop_Click(Index As Integer)
    Select Case Index
        Case 0, 1: ' activate/inactivate
            Call cmdActivate_Click(Index)
        Case 3 ' hide desktop
            ShowWindow hListView, 0&
            tmrAutoHide.Enabled = False
        Case 4 ' show desktop
            ShowWindow hListView, SW_SHOW
            SetAutoHideMode chkAutoHide.Value
        Case 5: ' Refresh desktop
            If DeskTopMsgBox = False Then
                If m_Rebuilding = False Then
                    RebuildListView
                    RefreshDeskTop False, False
                End If
            End If
        Case 7: ' save positions
            Call StoreIconPositions(, , , True)
            m_CanUnRestore = False
        Case 8 ' restore icon positions
            StoreIconPositions
            Call RestoreIcons(True)
            m_CanUnRestore = True
        Case 9 ' un restore last restore (reset to previous X,y)
            Call RestoreIcons(False)
            m_CanUnRestore = False
        Case 11:
            Call cmdReset_Click(0) ' reset all icons to system colors
    End Select
    If Index > 6 And Index < 10 Then Call mnuMain_Click(-1) ' reset check marks on menus
End Sub

Private Sub mnuGlobal_Click(Index As Integer)
    Select Case Index
        Case 0 ' autohide function
            If mnuGlobal(Index).Checked = True Then
                chkAutoHide.Value = ea_Disable
            Else ' enable autohide
                If cboAutoHide.ListIndex = 0 Then
                    cboAutoHide.ListIndex = 1
                Else
                    chkAutoHide.Value = ea_Enable
                End If
            End If
        Case 2, 3, 4 ' toggle eo_SaveExit, eo_RestoreRun, eo_RestoreRes
            lstMisc.Selected(Index - 1) = (Not mnuGlobal(Index).Checked)
        Case 6 'toggle eo_NoBalloon
            lstMisc.Selected(eo_NoBalloon) = (Not mnuGlobal(Index).Checked)
        Case 7 ' toggle eo_noSysTray
            lstMisc.Selected(eo_NoSysTray) = (Not mnuGlobal(Index).Checked)
        Case 9 ' show the Global options portion of the form
            chkView.Value = 1
    End Select

End Sub

Private Sub mnuMain_Click(Index As Integer)
    If Index = -1 Then
        mnuDesktop(9).Enabled = m_CanUnRestore
        mnuDesktop(8).Enabled = m_CanRestore
        mnuGlobal(2).Checked = lstMisc.Selected(eo_SaveExit)
        mnuGlobal(3).Checked = lstMisc.Selected(eo_RestoreRun)
        mnuGlobal(4).Checked = lstMisc.Selected(eo_RestoreRes)
        mnuGlobal(6).Checked = lstMisc.Selected(eo_NoBalloon)
        mnuGlobal(7).Checked = lstMisc.Selected(eo_NoSysTray)
    End If
End Sub

Private Sub mnuPop_Click(Index As Integer)
Select Case Index
    Case 1 ' show me
        If cTray.RestoreAnimated(Me.hwnd) = False Then ForceForeGround Me.hwnd
    Case 3 ' show desktop
        SetAutoHideMode ea_Disable
    Case 4 ' hide desktop
        m_dtAutoHide = DateAdd("h", -1, Now())
        SetAutoHideMode ea_ChkTimer
    Case 5 ' toggle autoHide
        Call mnuGlobal_Click(0)
    Case 7, 8, 9 ' save, restore, unrestore icon positions
        Call mnuDesktop_Click(Index)
    Case 11 ' toggle tweaker active/inactive
        If cmdActivate(0).Enabled = True Then ' is inactive
            Call cmdActivate_Click(0)
        Else
           Call cmdActivate_Click(1)
        End If
    Case 13: PostMessage Me.hwnd, WM_CLOSE, 0, 0 ' exit from systray popup
End Select
End Sub

Private Sub optNewDef_Click(Index As Integer)
    ' option to use system colors or custom colors for any new items added to destkop
    Call chkNewTrans_Click
End Sub

Private Sub tmrActiveDesktop_Timer()
    ' delay timer to allow active desktop to set up or clear
    tmrActiveDesktop.Enabled = False
    If IsActiveDesktop = True Then
        ' make tweaker inactive (not compatible with Active Desktop)
        If cmdActivate(1).Enabled = True Then
            Call cmdActivate_Click(1)
            RefreshDeskTop False, False
        End If
    Else
        ' make tweaker active only if user previously said to auto-activate
        If cmdActivate(0).Tag = "AutoActivate" Then Call cmdActivate_Click(0)
    End If

End Sub

Private Sub tmrAutoHide_Timer()
    SetAutoHideMode ea_ChkTimer   ' hide the desktop icons
End Sub

Private Sub tmrFlash_Timer()
    tmrFlash.Enabled = False
    Call cmdFlash_Click(Val(tmrFlash.Tag)) ' set next phase of the flasher
End Sub

Private Sub tmrRebuild_Timer()
    ' icons were added or deleted from desktop, rebuild our listing
    tmrRebuild.Interval = 0
    If DeskTopMsgBox = True Then
        '^^ check for Delete Confirmation messages and other message boxes generated
        ' from the Desktop; if so, don't rebuild yet and wait another two seconds or so
        tmrRebuild.Interval = 2000
    Else
        RebuildListView
    End If
End Sub

' //////////////// PRIVATE ROUTINES \\\\\\\\\\\\\\\\\\\\

Private Sub CleanUp()
' called when unloading or injection was successful but memory mapping failed

    On Error Resume Next
    ' terminate any timers
    tmrAutoHide.Enabled = False
    tmrFlash.Enabled = False
    tmrActiveDesktop.Enabled = False
    tmrRebuild.Enabled = False
    
    If Not cTray Is Nothing Then
        If cTray.IsActive Then cTray.RemoveTrayIcon
    End If
    ' terminate subclassing of desktop
    SendMessage hShell, wm_private, WM_NOTIFY, ByVal WM_CLOSE
    UnhookDesktop ' uninject the subclassing DLL from the desktop process
    SetAutoHideMode ea_Disable   ' ensure desktop is visible
    ' clean up memory objects
    FreeMapping hFileRx, hMapRx, True
    FreeMapping hFileTx, hMapTx, True
    hFileTx = 0: hMapTx = 0
    hFileRx = 0: hMapRx = 0
    wm_private = 0  ' also indicates we are no longer injected into desktop's process

End Sub

Private Sub GetDeskTopIconColors()
    ' return text and backround colors used for icons by the desktop's listview
    ' also caches the desktop listview & parent handles
    
    ' the text color is created by O/S based on the desktop background color
    hShell = GetTargetWindow()
    hListView = FindWindowEx(hShell, 0&, "SysListView32", vbNullString)
    lblDefColor(0).BackColor = SendMessage(hListView, LVM_GETTEXTCOLOR, 0&, ByVal 0&)
    ' the background color is a system color
    lblDefColor(1).BackColor = GetSysColor(COLOR_DESKTOP)

End Sub

Private Sub LoadSettings()

' column values:
' (1) Text Property is the icon caption reported by desktop
' (2) Customize mode: 0=none, 1=custom, 2=transparent bkg
' (3) Caption forecolor
' (4) Caption backcolor
' (5) User-adjusted icon caption
' (6) Current X,Y coords of icon
' (7) Last saved coords of icon

    Dim sAttr() As String
    Dim sSetting As String
    Dim tVal As Long
    Dim iniFile As String
    Dim I As Long
    Dim xItem As ListItem
    
    If lvDtop.ListItems.Count = 0 Then Exit Sub
    
    iniFile = SpecialFolderPath(CSIDL_PERSONAL) & GetUserName & ".desktop.twk"
    m_CurrentRes = MakeDWord((Screen.Width \ Screen.TwipsPerPixelX), (Screen.Height \ Screen.TwipsPerPixelY))
    
    ' get global cached settings
    ' 1. The setting for icons added to desktop: either 0,1,2
    tVal = Val(ReadWriteINI(iniFile, "Get", "Defaults", "NewIcons", 0))
    If (tVal And 2) = 2 Then chkNewTrans = 1    ' uses transparent backcolor
    If (tVal And 1) = 1 Then                    ' custom settings vs sys default
        optNewDef(0) = True
        ' 2. Get the custom fore & back colors
        tVal = ReadWriteINI(iniFile, "Get", "Defaults", "NewIconForeColor", -1)
        If tVal < 0 Then 'check for bogus cached value
            optNewDef(1) = True
        Else
            lblColor(2).BackColor = tVal
            tVal = ReadWriteINI(iniFile, "Get", "Defaults", "NewIconBackColor", -1)
            If tVal < 0 Then ' check for bogus cached value
                optNewDef(1) = True
            Else
                lblColor(3).BackColor = tVal
            End If
        End If
    End If
    
    On Error Resume Next ' in case someone was manually messing with the INI
    ' 3. get resolution of last saved positions
    m_MasterRes = Val(ReadWriteINI(iniFile, "Get", "Positions", "Resolution", "0"))
    m_CanRestore = (m_MasterRes <> 0) ' temporary flag to indicate mismatch resolution
    m_CanUnRestore = False            ' always false to start with
    
    ' 4. now we will get each icon's cached setting
    ' The current desktop's icons were added to our listview by
    '   using the CRC value of its caption as the .Text property and
    '   added the actual icon caption as the .Tag property
    sSetting = ReadWriteINI(iniFile, "Get", "Settings", "Icon1", "")
    I = 1
    Do While sSetting <> ""
        sAttr = Split(sSetting, ";")
        '^^ current format of INI for an icon is:
        '   crcVal of LCase(Caption); customize mode; forecolor; backcolor; len of custom caption
        
        ' find the CRC value in our listview
        Set xItem = lvDtop.FindItem(sAttr(0), lvwText)
        Do Until xItem Is Nothing
            ' found, is it a dup? There is nothing preventing you from making 2 or more
            '   icons on the desktop have the same name. For example, Temp.txt and Temp.Doc
            '   are the same name on the desktop if you are hiding known extensions.
            If xItem.SubItems(1) = "" Then ' not already assigned
                xItem.Text = xItem.Tag     ' assign the real caption
                For tVal = 1 To 3          ' assign any custom settings
                    xItem.SubItems(tVal) = sAttr(tVal)
                Next
                ' now test for a custom caption....
                If sAttr(tVal) = "-1" Then
                    xItem.Tag = ""
                ElseIf Val(sAttr(tVal)) > 0 Then
                    xItem.Tag = ReadWriteINI(iniFile, "Get", "Settings", "IconCaption" & I, xItem.Tag, , sAttr(tVal))
                End If
                ' upload the last saved position if any
                If Not m_MasterRes = 0 Then xItem.SubItems(7) = ReadWriteINI(iniFile, "Get", "Positions", "Icon" & I, "")
                Exit Do
            Else ' duplicated icon name
                ' When this happens, we will use the first cached icon we matched on
                If xItem.Index = lvDtop.ListItems.Count Then
                    ' special case: no other matches will exist; treat as a new desktop item
                    Exit Do
                Else
                    ' find the next match, if possible
                    Set xItem = lvDtop.FindItem(sAttr(0), lvwText, xItem.Index + 1)
                End If
            End If
        Loop
        ' get next cached icon
        I = I + 1
        sSetting = ReadWriteINI(iniFile, "Get", "Settings", "Icon" & I, "")
    Loop
    ClearCRCtable
    
    ' 5. set any icons not cached in INI file to have default settings
    For I = 1 To lvDtop.ListItems.Count
        If lvDtop.ListItems(I).SubItems(1) = "" Then
            lvDtop.ListItems(I).Text = lvDtop.ListItems(I).Tag
            tVal = Val(Mid$(lvDtop.ListItems(I).Key, 2))
            ' get current position of new icon if no change in resolution since last save
            If m_MasterRes = m_CurrentRes Then StoreIconPositions tVal, , , True
            GetSetColors True, tVal, 0, 0
        End If
    Next
    
    ' 6. Get autoHide settings
    sSetting = ReadWriteINI(iniFile, "Get", "Defaults", "AutoHide", "0;0")
    I = InStr(sSetting, ";")
    cboAutoHide.ListIndex = Val(Mid$(sSetting, I + 1))
    chkAutoHide.Value = Val(Left$(sSetting, I - 1))
    ' 7. Get No balloon, no auto-res position, autosave position options
    sSetting = ReadWriteINI(iniFile, "Get", "Defaults", "Misc", "0")
    sAttr = Split(sSetting, ";")
    ReDim Preserve sAttr(0 To lstMisc.ListCount - 1)
    For I = 0 To lstMisc.ListCount - 1
        lstMisc.Selected(I) = CBool(sAttr(I))
    Next
        
    ' 8. Finish up
    m_CanRestore = True
    If m_MasterRes = 0 Then  ' we have no saved icon positions (1st run)
        StoreIconPositions , , , True ' all current postions act as saved positions
        m_MasterRes = m_CurrentRes
    Else
        m_AutoArrageSave = CBool(ReadWriteINI(iniFile, "Get", "Positions", "AutoArrange", "0"))
        ' see if last saved positions were in different resolution
        If m_MasterRes <> m_CurrentRes Then
            ResChangePositions m_CurrentRes ' resets m_curRes
            m_CurrentRes = m_MasterRes
        End If
    End If
    lvDtop.Refresh          ' ensure scrollbars are attached if it will have scrollbars
End Sub

Private Sub SaveSettings()
' lvDTop column values:
' (1) Text Property is the icon caption reported by desktop
' (2) Customize mode: 0=none, 1=custom, 2=transparent bkg
' (3) Caption forecolor
' (4) Caption backcolor
' (5) User-adjusted icon caption
' (6) Current X,Y coords of icon
' (7) Last saved coords of icon
    
    If lvDtop.ListItems.Count = 0 Then Exit Sub
    '^^ if something is wrong prevent overwriting existing INI
    
    Dim sAttr As String
    Dim tVal As Long
    Dim iniFile As String
    Dim I As Long, J As Long
    Dim capLen As Long
    
    iniFile = SpecialFolderPath(CSIDL_PERSONAL) & GetUserName & ".desktop.twk"
    ' clear the previous cached icons & positions from the INI file
    ReadWriteINI iniFile, "Write", "Settings", "", "", True
    ReadWriteINI iniFile, "Write", "Positions", "", "", True
        
    ' Write the global values first
    If lstMisc.Selected(eo_SaveExit) = True Then StoreIconPositions , , , True ' save current X,Y on exit
    ReadWriteINI iniFile, "Write", "Positions", "Resolution", m_MasterRes
    ReadWriteINI iniFile, "Write", "Positions", "AutoArrange", m_AutoArrageSave
    
    If ReadWriteINI(iniFile, "Write", "Defaults", "NewIcons", m_CustomNewColors) = "ok" Then
        If ReadWriteINI(iniFile, "Write", "Defaults", "NewIconForeColor", lblColor(2).BackColor) = "ok" Then
            If ReadWriteINI(iniFile, "Write", "Defaults", "NewIconBackColor", lblColor(3).BackColor) = "ok" Then
                If ReadWriteINI(iniFile, "Write", "Defaults", "AutoHide", chkAutoHide.Value & ";" & cboAutoHide.ListIndex) = "ok" Then
                    For I = 0 To lstMisc.ListCount - 1
                        sAttr = sAttr & lstMisc.Selected(I) & ";"
                    Next
                    If ReadWriteINI(iniFile, "Write", "Defaults", "Misc", sAttr) = "ok" Then
                        ' Now write the individual icon
                        For I = 1 To lvDtop.ListItems.Count
                            With lvDtop.ListItems(I)
                                ' get CRC value of actual caption (reduces size of INI)
                                tVal = CRC32(LCase$(.Text))
                                ReadWriteINI iniFile, "Write", "Positions", "Icon" & I, .SubItems(7)
                                sAttr = tVal & ";"
                                ' concatenate the customize mode, forecolor & backcolor
                                For J = 1 To 3
                                    sAttr = sAttr & Val(.SubItems(J)) & ";"
                                Next
                                ' test for a custom display caption
                                If StrComp(.Tag, .Text, vbBinaryCompare) = 0 Then
                                    capLen = 0  ' flag indicates same as real caption
                                ElseIf .Tag = "" Then
                                    capLen = -1 ' flag indicates blank display caption
                                Else
                                    capLen = Len(.Tag)  ' len of custom display caption
                                End If
                                sAttr = sAttr & capLen  ' concatenate the caption code
                            End With
                            ' write the icon data
                            If ReadWriteINI(iniFile, "Write", "Settings", "Icon" & I, sAttr) = "ok" Then
                                ' when a custom caption exists, save it too
                                If capLen > 0 Then ReadWriteINI iniFile, "Write", "Settings", "IconCaption" & I, lvDtop.ListItems(I).Tag
                            Else
                                Exit For
                            End If
                        Next
                    
                    End If
                End If
            End If
        End If
    End If
    ClearCRCtable

End Sub

Private Sub InitializeListView()
    
    ' routine gets all currently displayed icons on the desktop
    ' and loads them into our listview. Done before we subclass the desktop
    
    Dim x As Long
    Dim nrIcons As Long
    
    Dim mMap As Long, tMap As Long
    Dim mHandle As Long, tHandle As Long
    
    Dim uItem As LVITEM
    Dim m_Caption As String
    
    nrIcons = SendMessage(hListView, LVM_GETITEMCOUNT, 0, ByVal 0&)
    If nrIcons = 0 Then Exit Sub
    ' force desktop to repaint so we have a clean slate for grabbing icon info
    RefreshDeskTop True, True
    
    ' create 2 inter process avenues:
    ' One for the listview Item & the other for the item text (pointers)
    
    ' create the listview item first (mMap & mHandle)
    mMap = CreateMapping(Len(uItem), mHandle, , hListView)
    If mMap = 0 Then Exit Sub
    ' now create the caption/text map (tMap & tHandle)
    tMap = CreateMapping(1024, tHandle, , hListView)
    If tMap = 0 Then
        FreeMapping mHandle, mMap
        Exit Sub
    End If
    
    lvDtop.ListItems.Clear
    lvDtop.Sorted = False
    
    With uItem                     ' fill in listview template
        .cchTextMax = m_maxChars
        .mask = LVIF_TEXT Or LVIF_IMAGE
    End With
    For x = 0 To nrIcons - 1
        uItem.pszText = tMap    ' ensure this is not overwritten & is correct
        uItem.iItem = x
        m_Caption = GetItemCaption(x, mHandle, mMap, tHandle, tMap, VarPtr(uItem), Len(uItem))
        ' inifile save captions as CRC values to reduce file size & hamper tampering
        lvDtop.ListItems.Add x + 1, "i" & x, CRC32(LCase$(m_Caption))
        lvDtop.ListItems(x + 1).Tag = m_Caption ' real caption as is
        lvDtop.ListItems(x + 1).SubItems(5) = uItem.iImage ' icon image index
    Next
    
    ' release the inter process avenues
    FreeMapping mHandle, mMap
    FreeMapping tHandle, tMap
    
End Sub

Private Sub RebuildListView(Optional bInitial As Boolean = False)
    
    ' this function is called on two known instances and there may be more....
    ' 1. One or more icons are removed from the desktop
    ' 2. One or more icons are added to the desktop
    ' When pressing F5 on the desktop, O/S deletes all then re-inserts all
    Dim x As Long, lRtn As Long
    Dim nrIcons As Long
    
    Dim mMap As Long, tMap As Long
    Dim mHandle As Long, tHandle As Long
    
    Dim uItem As LVITEM, xItem As ListItem
    Dim m_Caption As String, bRefresh As Boolean
    Dim iWhere As Integer, iOffset As Integer
    Dim lBestMatch As Long, iAccuracy As Integer
    
    m_Rebuilding = True ' flag to prevent recurrence
    
    nrIcons = SendMessage(hListView, LVM_GETITEMCOUNT, 0, ByVal 0&)
    If nrIcons = 0 Then
        m_Rebuilding = False
        Exit Sub
    End If
    If bInitial Then iWhere = lvwText Else iWhere = lvwTag
   
    ' create 2 inter process avenues:
    ' One for the listview Item & the other for the item text (pointers)
    
    ' create the listview item first (mMap & mHandle)
    mMap = CreateMapping(Len(uItem), mHandle, , hListView)
    If Not mMap = 0 Then
        ' now create the caption/text map (tMap & tHandle)
        tMap = CreateMapping(1024, tHandle, , hListView)
    End If
    
    If Not tMap = 0 Then
        For x = 1 To lvDtop.ListItems.Count ' reset key value for each current list item
            lvDtop.ListItems(x).Key = ""
        Next
        With uItem              ' fill in listview template
            .cchTextMax = m_maxChars
            .mask = LVIF_TEXT Or LVIF_IMAGE
            .pszText = tMap
        End With
        For x = 0 To nrIcons - 1
            iOffset = 1     ' where to start searching our listview
            iAccuracy = 0   ' reliability of match when item is found
            uItem.iItem = x
            m_Caption = GetItemCaption(x, mHandle, mMap, tHandle, tMap, VarPtr(uItem), Len(uItem))
            
            Do
                ' find the desktop caption in our listing
                Set xItem = lvDtop.FindItem(m_Caption, iWhere, iOffset)
                If xItem Is Nothing Then
                    ' couldn't find it, but if not the initial rebuild, we check custom text to
                    If Not bInitial Then Set xItem = lvDtop.FindItem(m_Caption, lvwText, iOffset)
                End If
                If xItem Is Nothing Then Exit Do ' not found. period
                ' item was found somewhere let's see if it was duplicated
                If xItem.Key = "" Then
                    ' not duplicated, let's check for accuracy in case caption is used more than once
                    If Not bInitial Then
                        ' unless an icon is animated, it is impossible to change the caption
                        ' and icon at the same time, so we will see if the image index matches.
                        ' Known example is the Recycle Bin which also changes its icon
                        If xItem.SubItems(5) = CStr(uItem.iImage) Then iAccuracy = 100 Else iAccuracy = 50
                        ' ^^ if caption matches but icon doesn't we'll look for a better match
                    Else
                        ' about as perfect as we can get: caption & icon match
                        iAccuracy = 100 ' nothing to compare to if app hasn't started yet
                    End If
                Else
                    ' item in our listview already assigned, look again
                    If xItem.Index = lvDtop.ListItems.Count Then
                        Set xItem = Nothing ' at end of list, can't look further
                        Exit Do
                    End If
                End If
                If iAccuracy < 100 Then ' continue looking?
                    lBestMatch = xItem.Index    ' cache best match so far
                    If xItem.Index = lvDtop.ListItems.Count Then
                        Exit Do
                    Else
                        iOffset = xItem.Index + 1   ' start of new search
                    End If
                Else
                    Exit Do ' got a near perfect match, exit search loop
                End If
            Loop
            If xItem Is Nothing Then
                ' last search returned empty, but did we have a "best match" before that?
                If iAccuracy > 0 Then Set xItem = lvDtop.ListItems(lBestMatch)
            End If
            If xItem Is Nothing Then    ' new item as far as we know
                Set xItem = lvDtop.ListItems.Add(, "i" & x, m_Caption)
                xItem.Tag = m_Caption
                xItem.SubItems(5) = uItem.iImage ' update icon image index
                GetSetColors True, x, 0, 0 ' assign default color scheme
                bRefresh = True
            Else                        ' existing item
                xItem.Key = "i" & x
                If xItem.Tag <> m_Caption Then
                    ' custom caption to be displayed, update it now
                    Call SetItemCaption(xItem.Tag, x, mHandle, mMap, tHandle, tMap, VarPtr(uItem), Len(uItem))
                    bRefresh = True
                End If
            End If
        Next
        
        ' ok, now remove any items that were not updated; these have been deleted from desktop
        For x = lvDtop.ListItems.Count To 1 Step -1
            If lvDtop.ListItems(x).Key = "" Then
                lvDtop.ListItems.Remove x
            End If
        Next
        
    End If
    
    ' release the inter process avenues
    FreeMapping mHandle, mMap
    FreeMapping tHandle, tMap
    
    lblCount.Caption = "Managing " & nrIcons & " icons"
    SetAutoHideMode chkAutoHide.Value
    
    If m_MasterRes <> m_CurrentRes Then
        RestoreIcons lstMisc.Selected(eo_RestoreRes)
        m_CurrentRes = m_MasterRes
    End If
    m_Rebuilding = False
    If bRefresh Then RefreshDeskTop False, False
    Call lvDtop_ItemClick(lvDtop.SelectedItem)
    
End Sub

Private Sub RefreshDeskTop(ByVal bAbsolute As Boolean, ByVal bImmediately As Boolean, Optional itemNr As Long = -1)

If itemNr > -1 Then ' refreshing a single icon
    If bImmediately Then
        SendMessage hListView, LVM_REDRAWITEMS, itemNr, ByVal itemNr
    Else
        PostMessage hListView, LVM_REDRAWITEMS, itemNr, itemNr
    End If

Else
    If bAbsolute Then
        ' this option triggers a LVM_DELETEALLITEMS & forces a refresh
        Const IDM_SHVIEW_REFRESH As Long = &H7103
        If bImmediately Then
            SendMessage hShell, WM_COMMAND, IDM_SHVIEW_REFRESH, ByVal 0&
        Else
            PostMessage hShell, WM_COMMAND, IDM_SHVIEW_REFRESH, 0&
        End If
    Else
        ' this option simply redraws existing icons without deleting them
        Dim nrIcons As Long
        nrIcons = SendMessage(hListView, LVM_GETITEMCOUNT, 0, ByVal 0&) - 1
        If bImmediately Then
            SendMessage hListView, LVM_REDRAWITEMS, 0&, ByVal nrIcons
        Else
            PostMessage hListView, LVM_REDRAWITEMS, 0&, nrIcons
        End If
    End If
End If

End Sub

Private Function GetItemCaption(ByVal itemNr As Long, _
        ByVal objHandle As Long, ByVal objMap As Long, _
        ByVal txtHandle As Long, ByVal txtMap As Long, _
        ByVal pBuffer As Long, ByVal bufLen As Long) As String

    Dim lRtn As Long, tCaption As String
    
    ' place the LVITEM structure into mapped memory
    If isNT Then
        WriteProcessMemory objHandle, objMap, pBuffer, bufLen, 0
    Else
        CopyMemory ByVal objMap, ByVal pBuffer, bufLen
    End If
    ' request the caption and/or image index
    lRtn = SendMessage(hListView, LVM_GETITEMA, itemNr, ByVal objMap)
    If lRtn Then
        ' ok, now get the modified LVITEM & caption from memory (byte array)
        If isNT Then
            ReadProcessMemory txtHandle, txtMap, VarPtr(lvCaption(0)), m_maxChars, 0
            ReadProcessMemory objHandle, objMap, pBuffer, bufLen, 0
        Else
            CopyMemory ByVal VarPtr(lvCaption(0)), ByVal txtMap, m_maxChars
            CopyMemory ByVal pBuffer, ByVal objMap, bufLen
        End If
        ' convert byte array into a VB string
        tCaption = StrConv(lvCaption, vbUnicode) ' ANSI>Unicode
        GetItemCaption = Left$(tCaption, InStr(tCaption, Chr$(0)) - 1)
    End If
    

End Function

Private Function SetItemCaption(ByVal m_Caption As String, ByVal itemNr As Long, _
        ByVal objHandle As Long, ByVal objMap As Long, _
        ByVal txtHandle As Long, ByVal txtMap As Long, _
        ByVal pBuffer As Long, ByVal bufLen As Long) As Long

    Dim lRtn As Long, bCaption() As Byte
    
    ' convert string to byte array & place in mapped memory
    bCaption() = StrConv(m_Caption & Chr$(0), vbFromUnicode)
    If isNT Then
        WriteProcessMemory objHandle, objMap, pBuffer, bufLen, 0
        WriteProcessMemory txtHandle, txtMap, VarPtr(bCaption(0)), UBound(bCaption) + 1, 0
    Else
        CopyMemory ByVal objMap, ByVal pBuffer, bufLen
        CopyMemory ByVal txtMap, ByVal VarPtr(bCaption(0)), UBound(bCaption) + 1
    End If
    ' tell listview to update the caption
    lRtn = SendMessage(hListView, LVM_SETITEMTEXTA, itemNr, ByVal objMap)

End Function

Private Function GetSetColors(ByVal bSet As Boolean, ByVal Index As Long, _
                    ByRef fColor As Long, ByRef bColor As Long, _
                    Optional ByVal SetTransCheckBox As Boolean = False) As Long

' function either sets our listview's item colors or retrieves them
    If bSet Then
        ' only called when new items inserted on desktop
        With lvDtop.ListItems("i" & Index)
            If (m_CustomNewColors Or 1) = m_CustomNewColors Then ' custom vs default
                .SubItems(1) = m_CustomNewColors    ' custom code
                fColor = lblColor(2).BackColor      ' forecolor
                bColor = lblColor(3).BackColor      ' backcolor
            Else ' default
                If (m_CustomNewColors = 2) Then .SubItems(1) = 2 Else .SubItems(1) = 0
                fColor = lblDefColor(0).BackColor      ' forecolor
                bColor = lblDefColor(1).BackColor      ' backcolor
            End If
            .SubItems(2) = fColor
            .SubItems(3) = bColor
        End With
    Else
        Dim cMode As Long
        If Index < 0 Then   ' this is flag to return generic default colors
            cMode = m_CustomNewColors And Not 1
        Else                ' else we are returning colors for a specific item
            cMode = Val(lvDtop.ListItems("i" & Index).SubItems(1))
        End If
        Select Case cMode
            Case 0, 2: ' default fore & back colors w/wo transparency
                fColor = lblDefColor(0).BackColor
                bColor = lblDefColor(1).BackColor
            Case Else   ' custom fore & back colors
                fColor = Val(lvDtop.ListItems("i" & Index).SubItems(2))
                bColor = Val(lvDtop.ListItems("i" & Index).SubItems(3))
        End Select
        If Index < 0 Or SetTransCheckBox = True Then
            If bColor < 0 Then bColor = lblDefColor(1).BackColor
        Else
            If (cMode And 2) = 2 Then bColor = -1
        End If
        ' next line, if true, is called from the lvDtop_Click routine
        If SetTransCheckBox Then chkTrans.Value = Abs((cMode And 2) = 2)
        GetSetColors = cMode
                
    End If
    

End Function

Private Sub SetAutoHideMode(ByVal Mode As eAutoHideOptions)

Select Case Mode
    Case ea_Disable, ea_Reset
        ' disabler timer and ensure desktop is visible
        tmrAutoHide.Enabled = False
        If IsWindowVisible(hListView) = 0 Then ShowWindow hListView, SW_SHOW
        
    Case ea_Enable
        ' validate enable is good & then set timer if needed
        If cboAutoHide.ListIndex > 0 Then
            If cmdActivate(1).Enabled = True Then
                If GetForegroundWindow <> GetTargetWindow(True) Then SetAutoHideMode ea_SetTimer
            End If
        Else    ' bad values, turn autohide off
            chkAutoHide.Value = ea_Disable
        End If
        
    Case ea_SetTimer
        ' turn timer on
        If chkAutoHide.Value = ea_Enable Then
            m_dtAutoHide = DateAdd("n", Val(cboAutoHide.Text), Now())
            tmrAutoHide.Enabled = True
        End If
        
    Case ea_ChkTimer
        ' check to see if elapsed time passed & hide desktop if so
        If Now() >= m_dtAutoHide Then
            tmrAutoHide.Enabled = False
            If lstMisc.Selected(eo_NoBalloon) = False Then cTray.ShowBalloon "Clicking on the desktop will restore icons", "Hiding Icons", icInfo
            ShowWindow hListView, 0
        End If
    End Select
End Sub

Private Function InitializeInstance(ByVal Mode As eActivateMode) As Boolean

    Dim regMsgName As String
    Dim MapReceipt As String
    Dim mapTransmit As String
    Dim nrIcons As Long
    Dim lRtn As Long
    Dim bActiveDesktop As Boolean
    
    
    ' subclass our form so we can be notified of Explorer Window crashes,
    ' sysTray notifications, and notifications from our custom DLL
    If z_Funk Is Nothing Then
        sc_Subclass Me.hwnd
        ' get notification if explorer crashes so we can restart
        wm_expcrash = RegisterWindowMessage("TaskbarCreated")
'        sc_AddMsg Me.hwnd, wm_expcrash, MSG_AFTER
        sc_AddMsg Me.hwnd, wm_traynotify, MSG_BEFORE
        sc_AddMsg Me.hwnd, WM_SYSCOLORCHANGE, MSG_AFTER
        sc_AddMsg Me.hwnd, WM_SYSCOMMAND, MSG_BEFORE
        sc_AddMsg Me.hwnd, WM_SETTINGSCHANGE, MSG_AFTER
        sc_AddMsg Me.hwnd, WM_DISPLAYCHANGE, MSG_AFTER
        sc_AddMsg Me.hwnd, WM_NCHITTEST, MSG_BEFORE
    End If
    
    bActiveDesktop = IsActiveDesktop() ' check for active desktop
    lblActiveDT.Visible = bActiveDesktop
    ' below is true if we were tweaking and user switched to ActiveDesktop
    If Mode = ea_Inactivate And bActiveDesktop = True Then Mode = ea_ReInitialize

    If Not Mode = ea_Inactivate Then ' trying to initialize
        If bActiveDesktop Then   ' if wm_private=0 then we haven't hooked yet
            If wm_private = 0 Or (Me.Visible = True And Me.WindowState = vbDefault) Then
                If wm_private = 0 Then ForceForeGround Me.hwnd
                If MsgBox("Active Desktop is the current desktop mode." & vbCrLf & _
                    "Cannot tweak icons in that mode. Would you like to automatically " & vbCrLf & _
                    "start if you remove the Active Desktop mode?", vbYesNo + vbQuestion, "Active Desktop") = vbYes Then
                    cmdActivate(0).Tag = "AutoActivate"
                End If
            Else
                If cTray.isBalloonCapable Then ' inform user that we stopped tweaking
                    If lstMisc.Selected(eo_NoSysTray) = False Then
                        ' only show balloon if we are also in the system tray
                        cTray.ShowBalloon "Return to this program to re-Activate icon tweaking when Active Desktop is removed", "Cannot tweak icons in Active Desktop mode", icInfo
                    Else ' otherwise auto-activate for user
                        cmdActivate(0).Tag = "AutoActivate"
                    End If
                Else    ' user is Win9x, auto-activate for user
                    cmdActivate(0).Tag = "AutoActivate"
                End If
            End If
            Mode = ea_Inactivate
        End If
    End If
        
    If Mode = ea_Inactivate Then
        If Not wm_private = 0 Then ' else nothing to do
            SendMessage hShell, wm_private, WM_NOTIFY, ByVal WM_ENDSESSION
            If Me.Visible = True And bActiveDesktop = False Then
                MsgBox "Per-Icon Settings will be Disabled.", vbInformation + vbOKOnly, "Notice"
            End If
        End If
        
    Else
        cmdActivate(0).Tag = ""
        GetDeskTopIconColors    ' ensure current color scheme
        ' ensure icons are on the desktop -- non-icon desktop not fully tested yet
        nrIcons = SendMessage(hListView, LVM_GETITEMCOUNT, 0, ByVal 0&)
        For lRtn = 0 To Picture1.UBound
            Picture1(lRtn).Enabled = (nrIcons > 0)
        Next
        If nrIcons = 0 Then
            ForceForeGround Me.hwnd
            MsgBox "No Icons Found the Desktop", vbExclamation + vbOKOnly, "Cannot Customize"
            Mode = ea_Inactivate
        Else
        
            If wm_private = 0 Then
                ' now attempt to set the global hook to the desktop thread
                If HookDesktop(Me.hwnd, mapTransmit, MapReceipt, wm_private) = 0 Then
                    ForceForeGround Me.hwnd
                    MsgBox "Cannot communicate with the desktop.", vbCritical + vbOKOnly, "Error"
                    
                Else ' injected fine, we need to set up memory file to talk with DLL
                    sc_AddMsg Me.hwnd, wm_private, MSG_BEFORE
                    
                    ' create 2 general purpose 1kb files used for receipt & transmission
                    ' between this app and the desktop process (2-way communiation btwn DLL)
                    hMapRx = CreateMapping(1024, hFileRx, MapReceipt, 0, True)
                    If hMapRx <> 0 Then
                        hMapTx = CreateMapping(1024, hFileTx, mapTransmit, 0, True)
                    End If
                    If hMapRx = 0 Or hMapTx = 0 Then
                        Call CleanUp    ' shouldn't happen, but did :: abort subclassing
                        ForceForeGround Me.hwnd
                        MsgBox "Cannot tweak desktop at this time. ", vbExclamation + vbOKOnly, "Error"
                        Mode = ea_Inactivate
                    End If
                End If
            Else ' we've previously hooked, let's tell DLL to reactivate
                SendMessage hShell, wm_private, WM_NOTIFY, ByVal WM_ACTIVATE
            End If
        End If
    End If
    
    If Not Mode = ea_Inactivate Then
        ' ensure auto-hide timer set if appropriate
        SetAutoHideMode chkAutoHide.Value
        RebuildListView (Mode = ea_Initialize)
        If Not lvDtop.SelectedItem Is Nothing Then
            lvDtop.SelectedItem.EnsureVisible
            Call lvDtop_ItemClick(lvDtop.SelectedItem)
        End If
    End If
    
    ' finish up
    chkView.Enabled = Abs(Mode <> ea_Inactivate)
    If chkView.Enabled = False Then chkView.Value = 1
    mnuMain(1).Enabled = chkView.Enabled
    mnuMain(2).Enabled = chkView.Enabled
    cmdActivate(0).Enabled = Abs(Mode = ea_Inactivate)
    cmdActivate(1).Enabled = chkView.Enabled
    
    'return value
    InitializeInstance = Mode

End Function

Private Sub StoreIconPositions(Optional ByVal itemNr As Long = -1, _
        Optional ByVal mHandle As Long, Optional ByVal mMap As Long, _
        Optional ByVal UserDirected As Boolean)

' if ItemNr < 0 then retrieve positions for all icons
' elseif mMap = 0 then create map & retrieve single ItemNr
' else retrieve single itemnr (mMap won't be zero)

' column values:
' (6) Current X,Y coords of icon
' (7) Last saved coords of icon


Dim bFreeMap As Boolean
Dim tPt As POINTAPI, xItem As ListItem
Dim x As Long

If itemNr < 0 Then
    ' get all icon positions
    mMap = CreateMapping(Len(tPt), mHandle, , hListView)
    If mMap <> 0 Then
        For x = 1 To lvDtop.ListItems.Count
            itemNr = Val(Mid$(lvDtop.ListItems(x).Key, 2))
            If SendMessage(hListView, LVM_GETITEMPOSITION, itemNr, ByVal mMap) <> 0 Then
                If isNT Then
                    ReadProcessMemory mHandle, mMap, VarPtr(tPt), Len(tPt), 0
                Else
                    CopyMemory ByVal VarPtr(tPt), ByVal mMap, Len(tPt)
                End If
            Else ' don't know how this could happen...
                tPt.x = 0: tPt.y = 0
            End If
            ' save in currentXY column
            lvDtop.ListItems(x).SubItems(6) = tPt.x & ";" & tPt.y
            ' store in savedXY column if user chose to Save Icon Positions
            If UserDirected Then lvDtop.ListItems(x).SubItems(7) = lvDtop.ListItems(x).SubItems(6)
        Next
    End If
    FreeMapping mHandle, mMap
    ' now we will also save the AutoArrange flag
    x = GetWindowLong(hListView, GWL_STYLE)
    If UserDirected Then
        m_AutoArrageSave = ((x And LVS_AUTOARRANGE) = LVS_AUTOARRANGE)
    Else
        m_AutoArrageRestore = ((x And LVS_AUTOARRANGE) = LVS_AUTOARRANGE)
    End If
    
Else    ' get position for a single object
    If mMap = 0 Then
        bFreeMap = True ' need a map file else one was provided
        mMap = CreateMapping(Len(tPt), mHandle, , hListView)
        If mMap = 0 Then
            FreeMapping mHandle, mMap
            Exit Sub
        End If
    End If
    ' get the requested icon position
    If SendMessage(hListView, LVM_GETITEMPOSITION, itemNr, ByVal mMap) <> 0 Then
        If isNT Then
            ReadProcessMemory mHandle, mMap, VarPtr(tPt), Len(tPt), 0
        Else
            CopyMemory ByVal VarPtr(tPt), ByVal mMap, Len(tPt)
        End If
    End If
    If bFreeMap Then FreeMapping mHandle, mMap
    On Error Resume Next
    Set xItem = lvDtop.ListItems("i" & itemNr)
    If Not xItem Is Nothing Then xItem.SubItems(7) = tPt.x & ";" & tPt.y
End If
End Sub

Private Sub RestoreIcons(Optional ByVal UserDirected As Boolean)
' column values:
' (6) Current X,Y coords of icon
' (7) Last saved coords of icon
    
   
Dim mMap As Long, mHandle As Long
Dim itemNr As Long
Dim tPt As POINTAPI
Dim fromIndex As Long, I As Integer
Dim lStyle As Long, lToggleAutoArrange As Long

    If SendMessage(hListView, LVM_GETITEMCOUNT, 0, ByVal 0&) = 0 Then Exit Sub

    ' we won't position if called due to resolution change and user doesn't want auto changing
    mMap = CreateMapping(Len(tPt), mHandle, mMap, hListView)
    If mMap = 0 Then
        m_CanRestore = False
        m_CanUnRestore = False
        Exit Sub
    End If

    If UserDirected = True Then fromIndex = 7 Else fromIndex = 6
    If isNT Then ' determine which command to use to toggle AutoArrange
        lToggleAutoArrange = IDM_TOGGLEAUTOARRANGEnt
    Else
        lToggleAutoArrange = IDM_TOGGLEAUTOARRANGE9x
    End If
    
    lStyle = GetWindowLong(hListView, GWL_STYLE)
    If (lStyle Or LVS_AUTOARRANGE) = lStyle Then
        SendMessage hShell, WM_COMMAND, lToggleAutoArrange, ByVal 0&
    End If
    
    For itemNr = 1 To lvDtop.ListItems.Count
        With lvDtop.ListItems(itemNr)
            If Len(.SubItems(fromIndex)) > 0 Then
                I = InStr(.SubItems(fromIndex), ";")
                tPt.x = Val(Left$(.SubItems(fromIndex), I - 1))
                tPt.y = Val(Mid$(.SubItems(fromIndex), I + 1))
                If isNT Then
                    WriteProcessMemory mHandle, mMap, VarPtr(tPt), Len(tPt), 0
                Else
                    CopyMemory ByVal mMap, ByVal VarPtr(tPt), Len(tPt)
                End If
                SendMessage hListView, LVM_SETITEMPOSITION32, Val(Mid$(.Key, 2)), ByVal mMap
            End If
        End With
    Next
    
    If (UserDirected = True And m_AutoArrageSave = True) Or _
        (UserDirected = False And m_AutoArrageRestore = True) Then
            If (lStyle And LVS_AUTOARRANGE) = 0 Then
                SendMessage hShell, WM_COMMAND, lToggleAutoArrange, ByVal 0&
            End If
    End If

FreeMapping mHandle, mMap

m_CanRestore = True
m_CanUnRestore = True
Call mnuMain_Click(-1)

End Sub

Private Sub ResChangePositions(ByVal newRes As Long)
' column values:
' (6) Current X,Y coords of icon
' (7) Last saved coords of icon

    ' function modifies the two saved icon position X,Y coords to proper resolution
    
    ' Side Note: The WITH statement below is rem'd out because of the following
    '   anomoly. Feel free to unrem & then adjust lvDtop.ListItems(itemNr)
    '   statements within the With nest to test it out. FYI: This routine is called
    '   after our window processes WM_DISPLAYCHANGE but before the subclasser returns.
    '   The anomoly....
    '   Running Age of Empires II, the game resets resolution at least twice before
    '   game play beings, each time the subclassing reports the correct resolution as
    '   does Screen.Width & Screen.Height. However, after exiting the game, resolution
    '   is set once more (back to previous). Now, Screen.Width & Screen.Height are
    '   still correct exiting the subclassing. But when pausing/terminating the app
    '   thereafter, Screen.Width = Screen.Height. I was able to reproduce this anomoly
    '   on demand. By rem'ing out the WITH structure, the anomoly disappeared. Strange?
    
Dim tPt As POINTAPI, itemNr As Long, x As Long
Dim xRatio As Single, yRatio As Single
Dim I As Integer

If newRes <> m_MasterRes And newRes <> 0 Then
    tPt.x = LoWord(newRes)
    tPt.y = HiWord(newRes)
    xRatio = LoWord(m_MasterRes) / tPt.x
    yRatio = HiWord(m_MasterRes) / tPt.y
    ' change saved positions to current resolution
    For itemNr = 1 To lvDtop.ListItems.Count
        For x = 7 To 6 Step -1
'            With lvDtop.ListItems(itemNr)
                If Len(lvDtop.ListItems(itemNr).SubItems(x)) > 0 Then
                    I = InStr(lvDtop.ListItems(itemNr).SubItems(x), ";")
                    tPt.x = Val(Left$(lvDtop.ListItems(itemNr).SubItems(x), I - 1))
                    tPt.y = Val(Mid$(lvDtop.ListItems(itemNr).SubItems(x), I + 1))
                    tPt.x = tPt.x / xRatio
                    tPt.y = tPt.y / yRatio
                    lvDtop.ListItems(itemNr).SubItems(x) = tPt.x & ";" & tPt.y
                ElseIf x = 7 Then
                    ' this can only occur during LoadSettings & is a fake res change message
                    StoreIconPositions Val(Mid$(lvDtop(itemNr).ListItems.Key, 2)), , , True
                    x = x - 1 ' skip next loop cause above function filled in info
                End If
'            End With
        Next
    Next
    m_CurrentRes = 0 ' flag to force reposition on next rebuild which has been activated
End If
m_MasterRes = newRes

End Sub


' //////////////// PAUL CATON'S SUBCLASSING ROUTINES \\\\\\\\\\\\\\\\\\\\
'-SelfSub code------------------------------------------------------------------------------------
Private Function sc_Subclass(ByVal lng_hWnd As Long, _
                    Optional ByVal lParamUser As Long = 0, _
                    Optional ByVal nOrdinal As Long = 1, _
                    Optional ByVal oCallback As Object = Nothing, _
                    Optional ByVal bIdeSafety As Boolean = True) As Boolean 'Subclass the specified window handle
'*************************************************************************************************
'* lng_hWnd   - Handle of the window to subclass
'* lParamUser - Optional, user-defined callback parameter
'* nOrdinal   - Optional, ordinal index of the callback procedure. 1 = last private method, 2 = second last private method, etc.
'* oCallback  - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
'* bIdeSafety - Optional, enable/disable IDE safety measures. NB: you should really only disable IDE safety in a UserControl for design-time subclassing
'*************************************************************************************************
Const CODE_LEN      As Long = 260                                           'Thunk length in bytes
Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))            'Bytes to allocate per thunk, data + code + msg tables
Const PAGE_RWX      As Long = &H40&                                         'Allocate executable memory
Const MEM_COMMIT    As Long = &H1000&                                       'Commit allocated memory
Const MEM_RELEASE   As Long = &H8000&                                       'Release allocated memory flag
Const IDX_EBMODE    As Long = 3                                             'Thunk data index of the EbMode function address
Const IDX_CWP       As Long = 4                                             'Thunk data index of the CallWindowProc function address
Const IDX_SWL       As Long = 5                                             'Thunk data index of the SetWindowsLong function address
Const IDX_FREE      As Long = 6                                             'Thunk data index of the VirtualFree function address
Const IDX_BADPTR    As Long = 7                                             'Thunk data index of the IsBadCodePtr function address
Const IDX_OWNER     As Long = 8                                             'Thunk data index of the Owner object's vTable address
Const IDX_CALLBACK  As Long = 10                                            'Thunk data index of the callback method address
Const IDX_EBX       As Long = 16                                            'Thunk code patch index of the thunk data
Const SUB_NAME      As String = "sc_Subclass"                               'This routine's name
  Dim nAddr         As Long
  Dim nid           As Long
  Dim nMyID         As Long
  
  If IsWindow(lng_hWnd) = 0 Then                                            'Ensure the window handle is valid
    zError SUB_NAME, "Invalid window handle"
    Exit Function
  End If

  nMyID = GetCurrentProcessId                                               'Get this process's ID
  GetWindowThreadProcessId lng_hWnd, nid                                    'Get the process ID associated with the window handle
  If nid <> nMyID Then                                                      'Ensure that the window handle doesn't belong to another process
    zError SUB_NAME, "Window handle belongs to another process"
    Exit Function
  End If
  
  If oCallback Is Nothing Then                                              'If the user hasn't specified the callback owner
    Set oCallback = Me                                                      'Then it is me
  End If
  
  nAddr = zAddressOf(oCallback, nOrdinal)                                   'Get the address of the specified ordinal method
  If nAddr = 0 Then                                                         'Ensure that we've found the ordinal method
    zError SUB_NAME, "Callback method not found"
    Exit Function
  End If
    
  If z_Funk Is Nothing Then                                                 'If this is the first time through, do the one-time initialization
    Set z_Funk = New Collection                                             'Create the hWnd/thunk-address collection
    z_Sc(14) = &HD231C031: z_Sc(15) = &HBBE58960: z_Sc(17) = &H4339F631: z_Sc(18) = &H4A21750C: z_Sc(19) = &HE82C7B8B: z_Sc(20) = &H74&: z_Sc(21) = &H75147539: z_Sc(22) = &H21E80F: z_Sc(23) = &HD2310000: z_Sc(24) = &HE8307B8B: z_Sc(25) = &H60&: z_Sc(26) = &H10C261: z_Sc(27) = &H830C53FF: z_Sc(28) = &HD77401F8: z_Sc(29) = &H2874C085: z_Sc(30) = &H2E8&: z_Sc(31) = &HFFE9EB00: z_Sc(32) = &H75FF3075: z_Sc(33) = &H2875FF2C: z_Sc(34) = &HFF2475FF: z_Sc(35) = &H3FF2473: z_Sc(36) = &H891053FF: z_Sc(37) = &HBFF1C45: z_Sc(38) = &H73396775: z_Sc(39) = &H58627404
    z_Sc(40) = &H6A2473FF: z_Sc(41) = &H873FFFC: z_Sc(42) = &H891453FF: z_Sc(43) = &H7589285D: z_Sc(44) = &H3045C72C: z_Sc(45) = &H8000&: z_Sc(46) = &H8920458B: z_Sc(47) = &H4589145D: z_Sc(48) = &HC4836124: z_Sc(49) = &H1862FF04: z_Sc(50) = &H35E30F8B: z_Sc(51) = &HA78C985: z_Sc(52) = &H8B04C783: z_Sc(53) = &HAFF22845: z_Sc(54) = &H73FF2775: z_Sc(55) = &H1C53FF28: z_Sc(56) = &H438D1F75: z_Sc(57) = &H144D8D34: z_Sc(58) = &H1C458D50: z_Sc(59) = &HFF3075FF: z_Sc(60) = &H75FF2C75: z_Sc(61) = &H873FF28: z_Sc(62) = &HFF525150: z_Sc(63) = &H53FF2073: z_Sc(64) = &HC328&

    z_Sc(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")                    'Store CallWindowProc function address in the thunk data
    z_Sc(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")                     'Store the SetWindowLong function address in the thunk data
    z_Sc(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")                     'Store the VirtualFree function address in the thunk data
    z_Sc(IDX_BADPTR) = zFnAddr("kernel32", "IsBadCodePtr")                  'Store the IsBadCodePtr function address in the thunk data
  End If
  
  z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)                  'Allocate executable memory

  If z_ScMem <> 0 Then                                                      'Ensure the allocation succeeded
    On Error GoTo CatchDoubleSub                                            'Catch double subclassing
      z_Funk.Add z_ScMem, "h" & lng_hWnd                                    'Add the hWnd/thunk-address to the collection
    On Error GoTo 0
  
    If bIdeSafety Then                                                      'If the user wants IDE protection
      z_Sc(IDX_EBMODE) = zFnAddr("vba6", "EbMode")                          'Store the EbMode function address in the thunk data
    End If
    
    z_Sc(IDX_EBX) = z_ScMem                                                 'Patch the thunk data address
    z_Sc(IDX_HWND) = lng_hWnd                                               'Store the window handle in the thunk data
    z_Sc(IDX_BTABLE) = z_ScMem + CODE_LEN                                   'Store the address of the before table in the thunk data
    z_Sc(IDX_ATABLE) = z_ScMem + CODE_LEN + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
    z_Sc(IDX_OWNER) = ObjPtr(oCallback)                                     'Store the callback owner's object address in the thunk data
    z_Sc(IDX_CALLBACK) = nAddr                                              'Store the callback address in the thunk data
    z_Sc(IDX_PARM_USER) = lParamUser                                        'Store the lParamUser callback parameter in the thunk data
    
    nAddr = SetWindowLongA(lng_hWnd, GWL_WNDPROC, z_ScMem + WNDPROC_OFF)    'Set the new WndProc, return the address of the original WndProc
    If nAddr = 0 Then                                                       'Ensure the new WndProc was set correctly
      zError SUB_NAME, "SetWindowLong failed, error #" & Err.LastDllError
      GoTo ReleaseMemory
    End If
        
    z_Sc(IDX_WNDPROC) = nAddr                                               'Store the original WndProc address in the thunk data
    RtlMoveMemory z_ScMem, VarPtr(z_Sc(0)), CODE_LEN                        'Copy the thunk code/data to the allocated memory
    sc_Subclass = True                                                      'Indicate success
  Else
    zError SUB_NAME, "VirtualAlloc failed, error: " & Err.LastDllError
  End If
  
  Exit Function                                                             'Exit sc_Subclass

CatchDoubleSub:
  zError SUB_NAME, "Window handle is already subclassed"
  
ReleaseMemory:
  VirtualFree z_ScMem, 0, MEM_RELEASE                                       'sc_Subclass has failed after memory allocation, so release the memory
End Function

'Terminate all subclassing
Private Sub sc_Terminate()
  Dim I As Long

  If Not (z_Funk Is Nothing) Then                                           'Ensure that subclassing has been started
    With z_Funk
      For I = .Count To 1 Step -1                                           'Loop through the collection of window handles in reverse order
        z_ScMem = .Item(I)                                                  'Get the thunk address
        If IsBadCodePtr(z_ScMem) = 0 Then                                   'Ensure that the thunk hasn't already released its memory
          sc_UnSubclass zData(IDX_HWND)                                     'UnSubclass
        End If
      Next I                                                                'Next member of the collection
    End With
    Set z_Funk = Nothing                                                    'Destroy the hWnd/thunk-address collection
  End If
End Sub

'UnSubclass the specified window handle
Private Sub sc_UnSubclass(ByVal lng_hWnd As Long)
  If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
    zError "sc_UnSubclass", "Window handle isn't subclassed"
  Else
    If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                           'Ensure that the thunk hasn't already released its memory
      zData(IDX_SHUTDOWN) = -1                                              'Set the shutdown indicator
      zDelMsg ALL_MESSAGES, IDX_BTABLE                                      'Delete all before messages
      zDelMsg ALL_MESSAGES, IDX_ATABLE                                      'Delete all after messages
    End If
    z_Funk.Remove "h" & lng_hWnd                                            'Remove the specified window handle from the collection
  End If
End Sub

'Add the message value to the window handle's specified callback table
Private Sub sc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    If When And MSG_BEFORE Then                                             'If the message is to be added to the before original WndProc table...
      zAddMsg uMsg, IDX_BTABLE                                              'Add the message to the before table
    End If
    If When And MSG_AFTER Then                                              'If message is to be added to the after original WndProc table...
      zAddMsg uMsg, IDX_ATABLE                                              'Add the message to the after table
    End If
  End If
End Sub

'Delete the message value from the window handle's specified callback table
Private Sub sc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    If When And MSG_BEFORE Then                                             'If the message is to be deleted from the before original WndProc table...
      zDelMsg uMsg, IDX_BTABLE                                              'Delete the message from the before table
    End If
    If When And MSG_AFTER Then                                              'If the message is to be deleted from the after original WndProc table...
      zDelMsg uMsg, IDX_ATABLE                                              'Delete the message from the after table
    End If
  End If
End Sub

'Call the original WndProc
Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    sc_CallOrigWndProc = _
        CallWindowProcA(zData(IDX_WNDPROC), lng_hWnd, uMsg, wParam, lParam) 'Call the original WndProc of the passed window handle parameter
  End If
End Function

'Get the subclasser lParamUser callback parameter
Private Property Get sc_lParamUser(ByVal lng_hWnd As Long) As Long
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    sc_lParamUser = zData(IDX_PARM_USER)                                    'Get the lParamUser callback parameter
  End If
End Property

'Let the subclasser lParamUser callback parameter
Private Property Let sc_lParamUser(ByVal lng_hWnd As Long, ByVal NewValue As Long)
  If IsBadCodePtr(zMap_hWnd(lng_hWnd)) = 0 Then                             'Ensure that the thunk hasn't already released its memory
    zData(IDX_PARM_USER) = NewValue                                         'Set the lParamUser callback parameter
  End If
End Property

'-The following routines are exclusively for the sc_ subclass routines----------------------------

'Add the message to the specified table of the window handle
Private Sub zAddMsg(ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim nBase  As Long                                                        'Remember z_ScMem
  Dim I      As Long                                                        'Loop index

  nBase = z_ScMem                                                            'Remember z_ScMem so that we can restore its value on exit
  z_ScMem = zData(nTable)                                                    'Map zData() to the specified table

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
    nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
  Else
    nCount = zData(0)                                                       'Get the current table entry count
    If nCount >= MSG_ENTRIES Then                                           'Check for message table overflow
      zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values"
      GoTo Bail
    End If

    For I = 1 To nCount                                                     'Loop through the table entries
      If zData(I) = 0 Then                                                  'If the element is free...
        zData(I) = uMsg                                                     'Use this element
        GoTo Bail                                                           'Bail
      ElseIf zData(I) = uMsg Then                                           'If the message is already in the table...
        GoTo Bail                                                           'Bail
      End If
    Next I                                                                  'Next message table entry

    nCount = I                                                              'On drop through: i = nCount + 1, the new table entry count
    zData(nCount) = uMsg                                                    'Store the message in the appended table entry
  End If

  zData(0) = nCount                                                         'Store the new table entry count
Bail:
  z_ScMem = nBase                                                           'Restore the value of z_ScMem
End Sub

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim nBase  As Long                                                        'Remember z_ScMem
  Dim I      As Long                                                        'Loop index

  nBase = z_ScMem                                                           'Remember z_ScMem so that we can restore its value on exit
  z_ScMem = zData(nTable)                                                   'Map zData() to the specified table

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
    zData(0) = 0                                                            'Zero the table entry count
  Else
    nCount = zData(0)                                                       'Get the table entry count
    
    For I = 1 To nCount                                                     'Loop through the table entries
      If zData(I) = uMsg Then                                               'If the message is found...
        zData(I) = 0                                                        'Null the msg value -- also frees the element for re-use
        GoTo Bail                                                           'Bail
      End If
    Next I                                                                  'Next message table entry
    
    zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table"
  End If
  
Bail:
  z_ScMem = nBase                                                           'Restore the value of z_ScMem
End Sub

'Error handler
Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String)
  App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  MsgBox sMsg & ".", vbExclamation + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
  zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                   'Get the specified procedure address
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
End Function

'Map zData() to the thunk address for the specified window handle
Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long
  If z_Funk Is Nothing Then                                                 'Ensure that subclassing has been started
    zError "zMap_hWnd", "Subclassing hasn't been started"
  Else
    On Error GoTo Catch                                                     'Catch unsubclassed window handles
    z_ScMem = z_Funk("h" & lng_hWnd)                                        'Get the thunk address
    zMap_hWnd = z_ScMem
  End If
  
  Exit Function                                                             'Exit returning the thunk address

Catch:
  zError "zMap_hWnd", "Window handle isn't subclassed"
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim I     As Long                                                         'Loop index
  Dim J     As Long                                                         'Loop limit
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, I, bSub) Then                                 'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, I, bSub) Then                              'Probe for a Form method
      If Not zProbe(nAddr + &H7A4, I, bSub) Then                            'Probe for a UserControl method
        Exit Function                                                       'Bail...
      End If
    End If
  End If
  
  I = I + 4                                                                 'Bump to the next entry
  J = I + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  Do While I < J
    RtlMoveMemory VarPtr(nAddr), I, 4                                       'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      RtlMoveMemory VarPtr(zAddressOf), I - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      RtlMoveMemory VarPtr(zAddressOf), I - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    I = I + 4                                                             'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Function                                                       'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
End Function

Private Property Get zData(ByVal nIndex As Long) As Long
  RtlMoveMemory VarPtr(zData), z_ScMem + (nIndex * 4), 4
End Property

Private Property Let zData(ByVal nIndex As Long, ByVal nValue As Long)
  RtlMoveMemory z_ScMem + (nIndex * 4), VarPtr(nValue), 4
End Property

'-Subclass callback, usually ordinal #1, the last method in this source file----------------------
Private Sub zWndProc1(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc and (if set to do so) the after
'*              original WndProc callback.
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter
'*************************************************************************************************

' uMsg :: if wm_private then coming from DLL else coming from the system
' wParam :: if uMsg=wm_private then has special meaning
' lParam :: if uMsg=wm_private then has special meaning

' The layout of the shared mapped files are as follows
' bytes 0-3, caption fore color
' bytes 4-7, caption back color
' bytes 8-43, LVITEM structure
' bytes 44-99, reserved/buffer
' bytes 100+ , the caption

' lvDTop column values:
' (1) Text Property is the icon caption reported by desktop
' (2) Customize mode: 0=none, 1=custom, 2=transparent bkg
' (3) Caption forecolor
' (4) Caption backcolor
' (5) User-adjusted icon caption
' (6) Recent X,Y coords of icon
' (7) Last user-saved coords of icon

Select Case uMsg
    
    Case wm_private ' messages from out custom DLL
        
        Dim m_Caption As String ' icon caption (string vs byte)
        Dim xItem As LVITEM
        
        Select Case wParam
            Case NM_CUSTOMDRAW  ' lParam is the icon index
                If tmrRebuild.Interval > 0 Then ' desktop is in flux
                    tmrRebuild.Interval = 0
                    tmrRebuild.Interval = 1000
                    ' return default color scheme until done rebuilding
                    lReturn = GetSetColors(False, -1, m_Attr(0), m_Attr(1))
                ElseIf m_Rebuilding = False Then
                    ' Inserted/added icons are handled in the RebuildListView routine
                    If lParam < lvDtop.ListItems.Count Then
                        ' get caption (ANSI), convert & strip off nulls
                        CopyMemory ByVal VarPtr(lvCaption(0)), ByVal hMapRx + 100, m_maxChars + 1
                        m_Caption = StrConv(lvCaption, vbUnicode)
                        m_Caption = Left$(m_Caption, InStr(m_Caption, Chr$(0)) - 1)
                        
                        With lvDtop.ListItems("i" & lParam)
                            If .Tag = m_Caption Then ' if the display captions match
                                ' return the custom fore/back colors
                                lReturn = GetSetColors(False, lParam, m_Attr(0), m_Attr(1))
                                
                                If .SubItems(5) = "" Then ' only happens when user edits icon
                                    ' get the LVITEM structure
                                    CopyMemory ByVal VarPtr(xItem), ByVal hMapRx + 8, 36
                                    .SubItems(5) = xItem.iImage ' reassign the image index
                                    If lvDtop.SelectedItem.Key = .Key Then
                                        ' update GUI if modified item is the selected item
                                        Call lvDtop_ItemClick(lvDtop.SelectedItem)
                                        lvDtop.SelectedItem.EnsureVisible
                                    End If
                                End If
                            Else
                                ' item doesn't match what we have; can occur when user
                                ' choose a desktop sort from its menu or toggles "Hide Known Extensions"
                                tmrRebuild.Interval = 0     ' activate Rebuild
                                tmrRebuild.Interval = 1000
                            End If
                        End With
                    End If
                End If
                ' save the color scheme to our memory mapped file
                If Not lReturn = 0 Then CopyMemory ByVal hMapTx, ByVal VarPtr(m_Attr(0)), 8
                
            Case LVN_GETDISPINFO ' caption or icon being changed
                If m_Rebuilding = False And tmrRebuild.Interval = 0 Then
                    ' only modify if not currently rebuilding and not tagged to rebuild
                    If lParam < lvDtop.ListItems.Count Then ' else don't handle new icons
                        ' get caption (ANSI), convert & strip off nulls
                        CopyMemory ByVal VarPtr(lvCaption(0)), ByVal hMapRx + 100, m_maxChars + 1
                        m_Caption = StrConv(lvCaption, vbUnicode)
                        m_Caption = Left$(m_Caption, InStr(m_Caption, Chr$(0)) - 1)
                        ' adjust the caption toggling sort so list is resorted
                        lvDtop.Sorted = False
                        With lvDtop.ListItems("i" & lParam)
                            .Text = m_Caption   ' make caption & display name same
                            .Tag = m_Caption
                            .SubItems(5) = ""   ' reset so icon can be retrieved when redrawn
                        End With
                        lvDtop.Sorted = True    ' force re-sort
                        lReturn = 1
                    End If
                End If
                
           Case LVN_DELETEITEM, LVN_INSERTITEM ' deletion/addition
                ' let the RebuildListView routine handle these
                If m_Rebuilding = False Then
                    tmrRebuild.Interval = 0
                    tmrRebuild.Interval = 1000
                End If
                lReturn = 1
                
            Case WM_SETFOCUS    ' clicked on desktop, ensure it is visible
                SetAutoHideMode ea_Reset
                lReturn = 1
            Case WM_KILLFOCUS   ' desktop lost focus, set timer
                SetAutoHideMode ea_SetTimer
                lReturn = 1
            
            Case WM_SETTINGSCHANGE ' DLL wants to know values for its popup menu
                If cmdActivate(0).Enabled = False Then lReturn = 1 ' are we active & is AutoHide active
                If chkAutoHide.Value = ea_Enable And cboAutoHide.ListIndex > 0 Then lReturn = lReturn Or 2
                If m_CanRestore = True Then lReturn = lReturn Or 4
                If m_CanUnRestore = True Then lReturn = lReturn Or 8
                
            Case WM_MENUSELECT ' item selected from the desktop's context menu
                Call mnuPop_Click(lParam - WM_USER)
                lReturn = 1
                
            Case WM_LBUTTONDBLCLK
                If lstMisc.Selected(eo_DblClkRestore) = True Then RestoreIcons True
            
            Case Else
                lReturn = 0 ' something unexpected here?
        End Select
        bHandled = True
    
    Case WM_SETTINGSCHANGE ' something changed, see if it is the Active Desktop
        tmrActiveDesktop.Enabled = True
        
    Case WM_SYSCOLORCHANGE  ' user changed desktop color scheme
        GetDeskTopIconColors
        Call lvDtop_ItemClick(lvDtop.SelectedItem)
        
    Case WM_DISPLAYCHANGE
        ResChangePositions lParam ' adjust icon position X,Y to new resolution
        tmrRebuild.Interval = 0
        tmrRebuild.Interval = 1000
        
    Case wm_expcrash    ' Explorer crashed
        InitializeInstance ea_Initialize
        If Not cTray Is Nothing Then
            If cTray.IsActive Then cTray.Restore wm_traynotify
        End If
    
    Case WM_SYSCOMMAND  ' are we being minimized?
        If wParam = SC_MINIMIZE Then
            lReturn = 1
            bHandled = True
            cTray.MinimizeAnimated lng_hWnd
            
        ElseIf wParam = SC_CLOSE Then
            ' prevent the close button from closing app, make it minimize instead
            If DefWindowProc(hwnd, WM_NCHITTEST, 0, lParam) = HTCLOSE Then
                ' ^^ ask our window where cursor was when sc_close was received
                bHandled = True
                lReturn = 1
                cTray.MinimizeAnimated lng_hWnd
            End If
        End If
    
    Case wm_traynotify  ' messages from our system tray icon
        SetForegroundWindow hwnd ' if we don't do this, the menu may hang if not clicked on
        If lParam = WM_RBUTTONUP Then
            mnuPop(1).Enabled = (Me.Visible = False)
            mnuPop(11).Checked = (cmdActivate(0).Enabled = False)
            mnuPop(5).Checked = (chkAutoHide.Value = ea_Enable And cboAutoHide.ListIndex > 0)
            PopupMenu mnuPopUp
        ElseIf lParam = WM_LBUTTONDOWN Then
            Call mnuPop_Click(1) ' when left clicking on icon, show our form
        End If
        bHandled = True
        
    Case Else
    
End Select

End Sub


' win98
'AUTOARRANGE -2139066303 TURN ON
'AUTOARRANGE -2138542015 TURN OFF
'Print LoWord(-2139066303),LoWord(-2138542015)
'28737  28737

' NT
'AUTOARRANGE -2139066287 TURN ON
'AUTOARRANGE -2138541999 TURN OFF
'Print LoWord(-2139066287), LoWord(-2138541999)
'28753  28753
