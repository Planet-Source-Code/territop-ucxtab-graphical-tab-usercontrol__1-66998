VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ucXTab - Test Harness"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin prjucXTab.ucXTab ucXTab1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _extentx        =   11245
      _extenty        =   5741
      activetab       =   5
      usecontrolbox   =   -1
      showfocusrect   =   0
      tabstyle        =   1
      tabtheme        =   1
      activetabbackendcolor=   16514555
      activetabbackstartcolor=   16514555
      disabledtabbackcolor=   -2147483633
      disabledtabforecolor=   10526880
      inactivetabbackendcolor=   14204606
      inactivetabbackstartcolor=   16777215
      outerbordercolor=   10198161
      tabcount        =   6
      activetabfont   =   "frmMain.frx":324A
      inactivetabfont =   "frmMain.frx":3278
      backcolor       =   16514555
      forecolor       =   10198161
      picturemaskcolor=   16711935
      tabcaption(0)   =   "Tab 0"
      tabcaption(1)   =   "Tab 1"
      tabcaption(2)   =   "Tab 2"
      tabcaption(3)   =   "Tab 3"
      tabcaption(4)   =   "Tab 4"
      tabcaption(5)   =   "Tab 5"
      Begin VB.CheckBox chkTabEnable 
         BackColor       =   &H00FBFDFB&
         Caption         =   "Tab 5"
         Height          =   255
         Index           =   5
         Left            =   1800
         TabIndex        =   68
         Top             =   1680
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkTabEnable 
         BackColor       =   &H00FBFDFB&
         Caption         =   "Tab 4"
         Height          =   255
         Index           =   4
         Left            =   1800
         TabIndex        =   67
         Top             =   1320
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkTabEnable 
         BackColor       =   &H00FBFDFB&
         Caption         =   "Tab 3"
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   66
         Top             =   960
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkTabEnable 
         BackColor       =   &H00FBFDFB&
         Caption         =   "Tab 2"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   65
         Top             =   1680
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkTabEnable 
         BackColor       =   &H00FBFDFB&
         Caption         =   "Tab 1"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   64
         Top             =   1320
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkTabEnable 
         BackColor       =   &H00FBFDFB&
         Caption         =   "Tab 0"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   63
         Top             =   960
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CommandButton cmdEnableTabs 
         Caption         =   "Reset"
         Height          =   315
         Left            =   3480
         TabIndex        =   61
         Top             =   960
         Width           =   1095
      End
      Begin VB.ListBox lstChkItems 
         Height          =   510
         Left            =   43480
         Style           =   1  'Checkbox
         TabIndex        =   60
         Top             =   2520
         Width           =   2535
      End
      Begin VB.CommandButton cmdAddCtrls 
         Caption         =   "New Tab"
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   11680
         TabIndex        =   59
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddCtrls 
         Caption         =   "Current Tab"
         Height          =   315
         Index           =   0
         Left            =   10480
         TabIndex        =   57
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtTabCaption 
         Height          =   315
         Left            =   12040
         TabIndex        =   55
         Text            =   "NewTab"
         Top             =   1800
         Width           =   1815
      End
      Begin VB.OptionButton opDynamicType 
         BackColor       =   &H00FBFDFB&
         Caption         =   "Remove Tab"
         Height          =   255
         Index           =   1
         Left            =   10480
         TabIndex        =   54
         Top             =   1350
         Width           =   1455
      End
      Begin VB.OptionButton opDynamicType 
         BackColor       =   &H00FBFDFB&
         Caption         =   "Insert Tab"
         Height          =   255
         Index           =   0
         Left            =   10480
         TabIndex        =   53
         Top             =   960
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.ComboBox cmbActionIndex 
         Height          =   315
         Left            =   12040
         TabIndex        =   50
         Text            =   "Combo1"
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmdAction 
         Caption         =   "Insert"
         Height          =   315
         Left            =   13000
         TabIndex        =   49
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   300
         Left            =   35280
         TabIndex        =   47
         Top             =   2680
         Width           =   735
      End
      Begin VB.CheckBox chkUseImageList 
         BackColor       =   &H00FBFDFB&
         Caption         =   "Use ImageList Pictures"
         Height          =   375
         Left            =   23600
         TabIndex        =   46
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox txtYRadius 
         Height          =   315
         Left            =   41440
         TabIndex        =   44
         Text            =   "txtYRadius"
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtXRadius 
         Height          =   315
         Left            =   41440
         TabIndex        =   42
         Text            =   "txtXRadius"
         Top             =   2040
         Width           =   1455
      End
      Begin VB.ComboBox cmbTabPictureIndex 
         Height          =   315
         Left            =   21440
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CheckBox chkUseFocusedColor 
         BackColor       =   &H00FBFDFB&
         Caption         =   "Use Focused Color"
         Height          =   255
         Left            =   33480
         TabIndex        =   37
         Top             =   2400
         Width           =   1815
      End
      Begin VB.CheckBox chkUseMaskColor 
         BackColor       =   &H00FBFDFB&
         Caption         =   "Use Mask Color"
         Height          =   255
         Left            =   33480
         TabIndex        =   36
         Top             =   2760
         Width           =   1455
      End
      Begin VB.ListBox lstColorItems 
         Height          =   2400
         Left            =   30360
         TabIndex        =   35
         Top             =   600
         Width           =   2535
      End
      Begin VB.CommandButton cmdItemColor 
         Caption         =   "..."
         Height          =   275
         Left            =   35715
         TabIndex        =   32
         Top             =   870
         Width           =   275
      End
      Begin VB.ComboBox cmbPictureAlign 
         Height          =   315
         Left            =   24680
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox cmbPictureSize 
         Height          =   315
         Left            =   24680
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox cmbActiveTab 
         Height          =   315
         Left            =   41440
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1100
         Width           =   1455
      End
      Begin VB.TextBox txtActiveTabHeight 
         Height          =   315
         Left            =   44560
         TabIndex        =   16
         Text            =   "txtActiveTabHgt"
         Top             =   600
         Width           =   1455
      End
      Begin VB.ComboBox cmbTabCount 
         Height          =   315
         Left            =   41440
         TabIndex        =   15
         Text            =   "cmbTabCount"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtInActiveTabHeight 
         Height          =   315
         Left            =   44560
         TabIndex        =   14
         Text            =   "txtInActiveTabHgt"
         Top             =   1100
         Width           =   1455
      End
      Begin VB.ComboBox cmbTabStyle 
         Height          =   315
         Left            =   41440
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1560
         Width           =   1455
      End
      Begin VB.ComboBox cmbTabTheme 
         Height          =   315
         Left            =   44560
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtToolTipText 
         Height          =   315
         Left            =   44560
         TabIndex        =   11
         Text            =   "txtToolTipText"
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton cmdStartColor 
         Caption         =   "..."
         Height          =   275
         Left            =   35715
         TabIndex        =   6
         Top             =   1590
         Width           =   275
      End
      Begin VB.CommandButton cmdEndColor 
         Caption         =   "..."
         Height          =   275
         Left            =   35715
         TabIndex        =   5
         Top             =   1950
         Width           =   275
      End
      Begin VB.TextBox txtStartColor 
         Height          =   315
         Left            =   34560
         TabIndex        =   7
         Text            =   "Locate Color..."
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtEndColor 
         Height          =   315
         Left            =   34560
         TabIndex        =   8
         Text            =   "Locate Color..."
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton cmdTabPicture 
         Caption         =   "..."
         Height          =   275
         Index           =   6
         Left            =   22595
         TabIndex        =   28
         Top             =   630
         Width           =   275
      End
      Begin VB.TextBox txtTabPicture 
         Height          =   315
         Left            =   21440
         TabIndex        =   29
         Text            =   "Locate Picture..."
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtItemColor 
         Height          =   315
         Left            =   34560
         TabIndex        =   33
         Text            =   "Locate Color..."
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Enabled Tab State:"
         Height          =   255
         Left            =   3480
         TabIndex        =   69
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblTabEnable 
         BackStyle       =   0  'Transparent
         Caption         =   "Dynamically Enabled Tabs:"
         Height          =   255
         Left            =   360
         TabIndex        =   62
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Dynamically Add Controls To:"
         Height          =   255
         Left            =   10360
         TabIndex        =   58
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label lblTabCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "Tab Caption:"
         Height          =   255
         Left            =   10480
         TabIndex        =   56
         Top             =   1830
         Width           =   1335
      End
      Begin VB.Label lblDynamicTab 
         BackStyle       =   0  'Transparent
         Caption         =   "AfterTab Index:"
         Height          =   255
         Left            =   12040
         TabIndex        =   52
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblDynamicTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Dynamically Insert / Remove Tabs:"
         Height          =   255
         Left            =   10360
         TabIndex        =   51
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Version: 1.0.48"
         Height          =   255
         Left            =   54800
         TabIndex        =   48
         Top             =   2340
         Width           =   1335
      End
      Begin VB.Label lblYRadius 
         BackStyle       =   0  'Transparent
         Caption         =   "Y-Radius"
         Height          =   315
         Left            =   40360
         TabIndex        =   45
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label lblXRadius 
         BackStyle       =   0  'Transparent
         Caption         =   "X-Radius"
         Height          =   315
         Left            =   40360
         TabIndex        =   43
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label lblTabPictIndex 
         BackStyle       =   0  'Transparent
         Caption         =   "Tab Picture Index"
         Height          =   435
         Left            =   20360
         TabIndex        =   40
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblItemColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   315
         Left            =   33480
         TabIndex        =   34
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblPictureAlign 
         BackStyle       =   0  'Transparent
         Caption         =   "Picture Align"
         Height          =   435
         Left            =   23600
         TabIndex        =   30
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblTabPicture 
         BackStyle       =   0  'Transparent
         Caption         =   "Tab Picture "
         Height          =   435
         Left            =   20360
         TabIndex        =   27
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblPicSize 
         BackStyle       =   0  'Transparent
         Caption         =   "Picture Size"
         Height          =   435
         Left            =   23600
         TabIndex        =   25
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblActiveTab 
         BackStyle       =   0  'Transparent
         Caption         =   "ActiveTab"
         Height          =   315
         Left            =   40360
         TabIndex        =   24
         Top             =   1100
         Width           =   855
      End
      Begin VB.Label lblTabCount 
         BackStyle       =   0  'Transparent
         Caption         =   "Tab Count"
         Height          =   315
         Left            =   40360
         TabIndex        =   23
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblTabStyle 
         BackStyle       =   0  'Transparent
         Caption         =   "Tab Style"
         Height          =   315
         Left            =   40360
         TabIndex        =   22
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblActiveTabHgt 
         BackStyle       =   0  'Transparent
         Caption         =   "ActiveTab Height"
         Height          =   435
         Left            =   43480
         TabIndex        =   21
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblInActibeTabHgt 
         BackStyle       =   0  'Transparent
         Caption         =   "InActiveTab Height"
         Height          =   435
         Left            =   43480
         TabIndex        =   20
         Top             =   1095
         Width           =   975
      End
      Begin VB.Label lblTabTheme 
         BackStyle       =   0  'Transparent
         Caption         =   "Tab Theme"
         Height          =   315
         Left            =   43480
         TabIndex        =   19
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lblToolTipText 
         BackStyle       =   0  'Transparent
         Caption         =   "ToolTipText"
         Height          =   315
         Left            =   43480
         TabIndex        =   18
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label lblEndColor 
         BackStyle       =   0  'Transparent
         Caption         =   "End Color"
         Height          =   315
         Left            =   33480
         TabIndex        =   10
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblStartColor 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Color"
         Height          =   315
         Left            =   33480
         TabIndex        =   9
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lblLink 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "click here"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   54395
         MouseIcon       =   "frmMain.frx":32A6
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label lblAuthorMessage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "To provide constructive feedback on this control, please                 ...."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   50360
         TabIndex        =   3
         Top             =   2880
         Width           =   5415
      End
      Begin VB.Image imXImage 
         Height          =   1920
         Left            =   54320
         Picture         =   "frmMain.frx":35B0
         Stretch         =   -1  'True
         Top             =   720
         Width           =   2040
      End
      Begin VB.Label lblWelcome 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":67FA
         ForeColor       =   &H80000008&
         Height          =   2000
         Index           =   1
         Left            =   50360
         TabIndex        =   2
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label lblWelcome 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to ucXTab!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   50360
         TabIndex        =   1
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label lblGradColorSelect 
         BackColor       =   &H00FBFDFB&
         BackStyle       =   0  'Transparent
         Caption         =   "Gradient Color Selection"
         Height          =   375
         Left            =   33480
         TabIndex        =   38
         Top             =   1275
         Width           =   2535
      End
      Begin VB.Label lblItemColorSel 
         BackColor       =   &H00FBFDFB&
         BackStyle       =   0  'Transparent
         Caption         =   "Item Color Selection"
         Height          =   375
         Left            =   33480
         TabIndex        =   39
         Top             =   600
         Width           =   2535
      End
   End
   Begin MSComctlLib.ImageList ilXImageList 
      Left            =   6120
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":69BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9C17
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BFF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11C1B
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'+  File Description:
'       ucXTab - SelfSubclassed Tab Replacement Container
'
'   Product Name:
'       ucXTab.ctl
'
'   Compatability:
'       Windows: 98, ME, NT4, 2000, XP
'
'   Software Developed by:
'       Paul R. Territo, Ph.D
'
'   Adapted from the following online article(s):
'       (Neeraj Agrawal - Original XTab Control)
'           http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=56462&lngWId=1
'       (Paul Caton - SelfSubclassing Template)
'           http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1
'       (Evan Todder - ContainedControls Tab Routine)
'           Note: The link below is inactive, as these submissions were removed by the author
'           http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=57642&lngWId=1
'       (Randy Birch - OS Version Detection)
'           http://vbnet.mvps.org/Index.html?code/helpers/iswinversion.htm
'       (James Laferriere - EqualRect API Routine)
'           http://www.officecomputertraining.com/vbtutorial/tutpages/page45.asp
'       (Dieter Otter - GetCurrentThemeName)
'           http://www.vbarchiv.net/archiv/tipp_805.html
'       (Fred.cpp - APILine, APIFillRectByCoords, APIRectangle)
'           http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=61476&lngWId=1
'
'   Legal Copyright & Trademarks (Current Implementation):
'       Copyright © 2005, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'       Trademark ™ 2005, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'
'   Comments:
'       No claims or warranties are expressed or implied as to accuracy or fitness
'       for use of this software. Paul R. Territo, Ph.D shall not be liable
'       for any incidental or consequential damages suffered by any use of
'       this software.
'
'-  Modification(s) History:
'       05Sep05 - Initial Selfsubclasing build of the ucXTab Control
'       07Sep05 - Added OS Detection support and ability to change the Tab style for
'                 WinXP and Classic Windows style Tab Controls.
'       10Sep05 - Removed redundant code and remaining Collections code from the
'                 UserControl.
'               - Performed cleanup on existing routines and optimized several data
'                 handling routines.
'       17Oct05 - Added Galaxy Theme to the control to match the style created by Fred.cpp
'                 for the isButton.
'       25Oct05 - Replaced the SelfSubclasser routines with ones directy from
'                 Paul Caton's source and updated the calls to work with the
'                 current UserControl.
'       26Oct05 - Added common property setting routine to improve encapsulation and
'                 code reuse when calling ItemColors.
'               - Eliminated "Call Forwarding" (Caller -> pMySub -> MySub) used in the
'                 original XTab project which in the Self-Subclassed UserControl is not
'                 required and simply results in methods overhead.
'       28Oct05 - Added MouseWheel Support for Tab Scrolling along with Associated
'                 MouseWheel Events for MouseScrollUp/Down.
'               - Fixed FocusColor bug when changing Tabs via ActiveTab property..
'               - Fixed Hover and FocusColor in PropertyDialog style Tab drawing which
'                 caused an XOr of the Tab top Border.
'               - Further optimized the base code to eliminate and consolidate calls for
'                 several drawing routines and tab typing routines.
'               - Alpha Sorted the code and variables...
'       29Oct05 - Added Overloaded TranslateColorEx method to allow for color conversion
'                 when the color selected is not represented. The result is a color which
'                 is mapped to the current pallete for display with out an error on the
'                 the caller's end.
'               - Additional code optimization to eliminate redundent and call forwarded
'                 routines. In addtion, removed reoutines which were called but did not
'                 provide any functionality (i.e. case statements with empty cases).
'               - Optimzed the pHandleMouseDown and pHandleMouseUp handlers to remove all
'                 redunant calls (i.e. all calls were alike)
'               - Added RemoveTabImage sub to allow for individual or all tab image
'                 removal from the control.
'       22Oct06 - Converted All drawing routines to API methods to provide near realtime updates
'               - Removed pLine method, which wrapped the Line method and used APIs instead
'               - Removed SetDefaultColor which was a "Call Forwarding" to ResetColorsToDefault
'               - General cleanup and optimizations.
'               - Added Version Property
'               - Fixed minor BackColor and ForeColor bug which prevented persistance in the object
'       25Oct06 - Added pAlphaBlend method to provide color mixing along the tabs Focused Color or
'                 Hover Color when in XP Theme, and allows for smoother color transitions along the
'                 edges and a more rounded appearance.
'               - Added additional Highlights/Lowlights to the pDrawOverXOrdTabbed and pDrawOverXOrdProperty
'               - Fixed TabStripBackColor property bug which which prevented set backcolors from persisting
'                 in the object once set.
'       31Oct06 - Fixed FocusColor XOr Drawing bug which incorrectly painted the default XP cap color for the
'                 tabs when the conrol lost and regained focus.
'       04Nov06 - Fixed FocusRect size bug for XP Theme that painted the FocusRect over the Focus Cap
'                 color on the Focused Tab (Whew, too many Focuses in one statement ;-)
'       05Nov06 - Fixed ResetColorsToDefault bug in UserControl_ReadProperties method, which prevented
'                 custom colors to be retained from Design Time - Thanks Mirko Kressmann for catching this ;-)
'               - Added InsertTab for dynamic tab addition which can be placed anywhere in the tab order
'               - Added RemoveTab for dynamic tab removal from any place in the tab order
'       09Nov06 - Fixed Active and InActive Tab Cap Strip bug pointed out by Mirko Kressmann, when the
'                 FocusColor = ActiveTabStartBackcolor or InActiveTabStartBackcolor for both Tab and Prop page styles
'       10Nov06 - Fixed Active and InActive Tab Cap Strip bug or which painted the incorrect XOr Color
'                 when the HoverColor = ActiveTabStartBackcolor or InActiveTabStartBackcolor for both Tab and Prop page styles
'               - Added AddControl method to allow dynamic control addition to the tabs once
'                 they were created dynamically.
'       11Nov06 - Added ControlBox drawing code to paint a control closure box on each tab for all styles
'                 Added UseControlBox property to allow the developer to choose if for tab closure ControlBox is
'                 shown on each tab...
'       12Nov06 - Added ControlBoxRect to TabInfo Type to store values for hit testing of the control boxes
'               - Added WM_LMOUSEUP uMsg to subclass when the control box mouseup event occured. This allows
'                 the user to MouseDown using the WM_LMOUSEDOWN uMsg and paint the control, but not close the
'                 Tab.....if the mouse is still over the ControlBox on MouseUp (WM_LMOUSEUP) the Tab is removed
'               - Added additonal drawing routines to DrawControlBox method to allow for all styles and themes.
'               - Added Design Time Enum Locking for all Enums to prevent the Case Sensitive Bug from occuring
'                 when selecting variables in the IDE.
'               - Changed all Const which did not need to be public to private to provide better encapsulation.
'               - Explicitly set all Enums to an assigned values (i.e. &H0, &H1....&H6)
'               - Set GetThemeInfo to Public so the developer can probe the color name directly....
'               - Fixed Minor FocusRect alignment bug caused by adding ControlBox buttons to PropertyPage Styles
'               - Added ControlBoxEnter, ControlBoxHover, ControlBoxExit, ControlBoxMouseDown, ControlBoxMouseUp events
'               - Added TabRemove, TabInsert events
'       13Nov06 - Fixed InactiveTabs ControlBox Flicker for RoundTab and VisualStudio Themes, by limiting the
'                 paint events to the current tab.
'       27Nov06 - Added TabOffset to allow the developer to set the m_lMoveOffset values independent of
'                 the usercontrols width.
'       19Dec06 - Added Additonal Host-Subclassing methods to capture the QuereyClose event....this helps to
'                 prevent GPFs if the Host Object is closed before we can unsubclass our usercontrol.
'               - Fixed TabEnabled property and wire the functionality to the new controlbox drawings
'               - Added additional subclassing flags to prevent GFPs when the subclasser is shut down
'               - Added TabOffset property to allow the developer to set the Left shift value used in the
'                 HandleConatinedControls method. The original used a combination of a fixed offset (10000) units
'                 or a dynamic value which was set based on the design time property of the tab's width....
'                 the new method prevents the possible issue of setting the offset for tabs at design time
'                 and having them change at runtime or increased at a later design time....the result of such
'                 a change is tabs which do not shift the controls by enough to clear the screen ;-( By allowing
'                 the developer to set these values to a large (~50-70K) value, the chance for this occuring
'                 is minimized.
'               - Fixed painting bug in the ControlBox when tabs were disabled, but were painted as enabled
'               - Fixed all drawing routines to permit Enabled/Disabled painting for each tab
'               - Fixed drawing routines to Enabled/Disabled painting for whole usercontrol
'
'   Force Declarations
Option Explicit

'   Build Date & Time: 12/19/2006 10:00:37 PM
Const Major As Long = 1
Const Minor As Long = 0
Const Revision As Long = 130
Const DateTime As String = "12/19/2006 10:00:37 PM "

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Dim bLoading    As Boolean
Dim i           As Long
Dim lPrevTheme  As Long
Dim lPrevCount  As Long
Dim lIndex      As Long

'   Link URL address which searches for our control submission on PCS
Const sLink As String = "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&?lngWId=1&grpCategories=&txtMaxNumberOfEntriesPerPage=10&optSort=Alphabetical&chkThoroughSearch=&blnTopCode=False&blnNewestCode=False&blnAuthorSearch=False&lngAuthorId=&strAuthorName=&blnResetAllVariables=&blnEditCode=False&mblnIsSuperAdminAccessOn=False&intFirstRecordOnPage=1&intLastRecordOnPage=10&intMaxNumberOfEntriesPerPage=10&intLastRecordInRecordset=499&chkCodeTypeZip=&chkCodeDifficulty=&chkCodeTypeText=&chkCodeTypeArticle=&chkCode3rdPartyReview=&txtCriteria=ucXTab"

'Private Sub chkEnableAll_Click()
'    With Me
'        .ucXTab1.Enabled = (.chkEnableAll.Value = vbChecked)
'    End With
'End Sub

Private Sub chkTabEnable_Click(Index As Integer)
    With Me
        '   Dynamically Enable/Disable Tabs
        .ucXTab1.TabEnabled(Index) = (.chkTabEnable(Index).Value = vbChecked)
        '   Paint the controls like CheckBoxes which do not
        '   have a transparent backstyle
        Call pSetCtrlBackColors
        .ucXTab1.SetFocus
    End With
End Sub

Private Sub chkUseFocusedColor_Click()
    With Me
        If Not bLoading Then
            '   Set the focused color flag...for XP TabStyles Only
            .ucXTab1.UseFocusedColor = Abs(.chkUseFocusedColor.Value)
            .ucXTab1.SetFocus
        End If
    End With
End Sub

Private Sub chkUseImageList_Click()
    With Me
        If Not bLoading Then
            If Abs(.chkUseImageList.Value) Then
                Call .ucXTab1.CopyTabImagesFromImageList(.ilXImageList)
            Else
                Call .ucXTab1.RemoveTabImages(bRemoveAll:=True)
                .ucXTab1.SetFocus
            End If
        End If
    End With
End Sub

Private Sub chkUseMaskColor_Click()
    With Me
        If Not bLoading Then
            '   Set the MaskColor flag when using bitmaps
            .ucXTab1.UseMaskColor = Abs(.chkUseMaskColor.Value)
            .ucXTab1.SetFocus
        End If
    End With
End Sub

Private Sub cmbActiveTab_Click()
    On Error Resume Next
    With Me
        If Not bLoading Then
            '   Set the active tab...same as clicking the tab block
            .ucXTab1.ActiveTab = .cmbActiveTab.ListIndex
            .ucXTab1.SetFocus
        End If
    End With
End Sub

Private Sub cmbPictureAlign_Click()
    With Me
        If Not bLoading Then
            '   Set the picture alignment property
            .ucXTab1.PictureAlign = .cmbPictureAlign.ListIndex
            .ucXTab1.SetFocus
        End If
    End With
End Sub

Private Sub cmbPictureSize_Click()
    With Me
        If Not bLoading Then
            '   Set the picture size property
            .ucXTab1.PictureSize = .cmbPictureSize.ListIndex
            .ucXTab1.SetFocus
        End If
    End With
End Sub

Private Sub cmbTabCount_Click()
    With Me
        If Not bLoading Then
            '   Set the flag to prevent recursion
            bLoading = True
            '   Set the tab count properties
            .ucXTab1.TabCount = .cmbTabCount.List(.cmbTabCount.ListIndex)
            lPrevCount = .cmbTabCount.List(.cmbTabCount.ListIndex)
            '   Rebuild the droplists based on the tabcount property
            Call pReBuildListValues
            If .cmbTabCount.ListCount > 1 Then
                .cmbActiveTab.ListIndex = 1
            Else
                .cmbActiveTab.ListIndex = 0
            End If
            '   Now reflect the number of new tabs seleted
            '   Note: This is a special case since we don't want to
            '         have only one (1) tab, so we need to ajust the
            '         ListIndex to reflect this...
            .cmbTabCount.ListIndex = .ucXTab1.TabCount - 2
            '   Init the droplist index
            .cmbTabPictureIndex.ListIndex = 0
            '   Rebuild the tab names in case they were chnaged
            Call pBuildTabNames
            '   We are done, so set it back
            bLoading = False
        End If
    End With
End Sub

Private Sub cmbTabCount_KeyDown(KeyCode As Integer, Shift As Integer)
    With Me
        If Not bLoading Then
            If KeyCode = vbKeyReturn Then
                '   Allow the user to enter a value for TabCount
                If IsNumeric(cmbTabCount.Text) And (.cmbTabCount.Text > 1) Then
                    '   Set the count based on the changes made
                    .ucXTab1.TabCount = CLng(.cmbTabCount.Text)
                    lPrevCount = CLng(.cmbTabCount.Text)
                    '   Rebuild out droplists based on this info
                    Call pReBuildListValues
                    '   Set the activetab index
                    .cmbActiveTab.ListIndex = 1
                    '   Set the Tabcount index
                    .cmbTabCount.ListIndex = .cmbTabCount.ListCount - 1
                    '   Set the PictureIndex index
                    .cmbTabPictureIndex.ListIndex = 0
                Else
                    .cmbTabCount.Text = lPrevCount
                    MsgBox "The Value Selected is Invalid, Please Enter a Valid Numeric Value.", vbExclamation, "ucXTab"
                End If
            End If
            .ucXTab1.SetFocus
        End If
    End With
End Sub

Private Sub cmbTabStyle_Click()
    With Me
        If Not bLoading Then
            '   Set our tab style based on the selection
            '   (TabDialog or PropPageDialog Styles)
            .ucXTab1.TabStyle = .cmbTabStyle.ListIndex
            '   Paint the controls like CheckBoxes which do not
            '   have a transparent backstyle
            Call pSetCtrlBackColors
            .ucXTab1.SetFocus
        End If
    End With
End Sub

Private Sub cmbTabTheme_Click()
    With Me
        If Not bLoading Then
            '   Set our tab Theme based on the selection
            '   (Win9x, WinXP.....)
            .ucXTab1.TabTheme = .cmbTabTheme.ListIndex
            '   Paint the controls like CheckBoxes which do not
            '   have a transparent backstyle
            Call pSetCtrlBackColors
            .ucXTab1.SetFocus
        End If
    End With
End Sub

Private Sub cmdAction_Click()
    With Me
        '   See if we are inserting or removing
        If .opDynamicType(0).Value = True Then
            .cmbTabCount.AddItem .cmbActiveTab.ListCount
            .cmbActiveTab.AddItem .cmbActiveTab.ListCount
            .cmbActionIndex.AddItem .cmbActionIndex.ListCount
            lIndex = .cmbActionIndex.ListIndex + 1
            .cmdAddCtrls(1).Enabled = True
            '   Inserting, so call our routine
            Call .ucXTab1.InsertTab(.cmbActionIndex.ListIndex, .txtTabCaption.Text)
        Else
            If .ucXTab1.TabCount > 1 Then
                .cmbTabCount.AddItem .cmbActiveTab.ListCount - 1
                .cmbActiveTab.RemoveItem .cmbActiveTab.ListCount - 1
                .cmbActionIndex.RemoveItem .cmbActionIndex.ListCount - 1
                '   Removing, so call our routine
                Call .ucXTab1.RemoveTab(.cmbActionIndex.ListIndex)
            End If
        End If
    End With
End Sub

Private Sub cmdAddCtrls_Click(Index As Integer)
    Dim Ctrl As Control
    Dim lLeft As Long
    Dim lTop As Long
    Dim lhWnd As Long
    Static i As Long
    Static j As Long
    
    With Me
        If Index = 0 Then
            '   Keep the demo from running off the tab
            If i < 16 Then
                '   Dynamically add a button from its class name
                Set Ctrl = frmMain.Controls.Add("VB.CommandButton", "cmdTest" & i, .ucXTab1)
                '   Now adjust the size
                With Ctrl
                    .Caption = "cmdTest" & i
                    .Width = 1024
                    .Height = 315
                End With
                '   Compute the position on the current tab...
                If i < 2 Then
                    lLeft = 4000 + (Ctrl.Width * i)
                    lTop = 500
                Else
                    If (i Mod 2) = 0 Then
                        lLeft = 4000
                        lTop = 500 + (Ctrl.Height * i \ 2)
                    Else
                        lLeft = 4000 + Ctrl.Width
                        lTop = 500 + (Ctrl.Height * i \ 2) - Ctrl.Height / 2
                    End If
                End If
                '   Call our routine to add this as a member of the container
                '   Note: this effectively sets the parent property of the object
                '   via APIs so that the host object is the UserControl
                Call .ucXTab1.AddControl(Ctrl, , lLeft, lTop, , , 4)
                i = i + 1
            Else
                MsgBox "Do you really need more convincing that this works ;-)", vbExclamation + vbOKCancel, "ucXTab"
            End If
        ElseIf Index = 1 Then
            '   Add a New web page to the tab and navigate to PCS
            Set Ctrl = frmMain.Controls.Add("Shell.Explorer.2", "WebBrowser" & j, ucXTab1)
            '   Get the Handle to the WebBrowser, since it does not expose this
            'lhWnd = GetWebBrowserHandle(.ucXTab1.hwnd)
            '   Add the control to the ucXTab by setting its parent property via API
            '   Note: This can be done as and Object or Pointer to the object (hWnd)
            'Call .ucXTab1.AddControl(,lhWnd , 200, 500, .ucXTab1.Width - 400, .ucXTab1.Height - 600, lIndex)
            '   If one used only the hWnd of the object, then we need to process the
            '   the location.....the following is a work around until I figure out why
            '   the control Left value of is showing -40000 for example, when it should
            '   be the value we passed ;-(
            '   Note: this only occurs if the hWnd is pass and not object or the
            '   object and the hWnd...odd....
            '   Until I fix this, the following work around will do...
            'Ctrl.Move 200, 500, .ucXTab1.Width - 400, .ucXTab1.Height - 600
            '
            '   Call our routine to add this as a member of the container
            '   Note: this effectively sets the parent property of the object
            '   via APIs so that the host object is the UserControl
            Call .ucXTab1.AddControl(Ctrl, , 200, 500, .ucXTab1.Width - 400, .ucXTab1.Height - 600, lIndex)
            '   Navigate to PCS and find our page for ucXTab
            Call Ctrl.Navigate(sLink)
            j = j + 1
        End If
    End With
End Sub

Private Sub cmdEnableTabs_Click()
    Dim i As Long
    With Me.ucXTab1
        '   Enable every tab....
        For i = 0 To .TabCount - 1
            .TabEnabled(i) = True
            chkTabEnable(i).Value = vbChecked
        Next
    End With
End Sub

Private Sub cmdEndColor_Click()
    Dim psColor        As SelectedColor
    
    With Me
        '   Pick a color from the Color Dialog
        psColor = ShowColor(.hwnd, True)
        If psColor.bCanceled = False Then
            Select Case .lstColorItems.List(.lstColorItems.ListIndex)
                Case "ActiveTabBackColor"
                    '   Set the end color for the gradient for the ActiveTabBack
                    .txtEndColor.Text = pHexColorStr(psColor.oSelectedColor)
                    .ucXTab1.ActiveTabBackEndColor = CLng(psColor.oSelectedColor)
                Case "InActiveTabBackColor"
                    '   Set the end color for the gradient for the InActiveTabBack
                    .txtEndColor.Text = pHexColorStr(psColor.oSelectedColor)
                    .ucXTab1.InActiveTabBackEndColor = CLng(psColor.oSelectedColor)
            End Select
            '   Paint the controls like CheckBoxes which do not
            '   have a transparent backstyle
            Call pSetCtrlBackColors
        End If
        .ucXTab1.SetFocus
    End With

End Sub

Private Sub cmdItemColor_Click()
    Dim psColor        As SelectedColor
    
    With Me
        '   Pick a color from the Color Dialog
        psColor = ShowColor(.hwnd, True)
        If psColor.bCanceled = False Then
            Call pSetItemColor(CLng(psColor.oSelectedColor))
        End If
        .ucXTab1.SetFocus
    End With
End Sub

Private Sub cmdReset_Click()
    With Me
        .ucXTab1.ResetAllColors
        '   Paint the controls like CheckBoxes which do not
        '   have a transparent backstyle
        Call pSetCtrlBackColors
        '   Now set the current colors in the dialog back
        Call lstColorItems_Click
        .ucXTab1.SetFocus
    End With
End Sub

Private Sub cmdStartColor_Click()
    Dim psColor        As SelectedColor
    
    With Me
        '   Pick a color from the Color Dialog
        psColor = ShowColor(.hwnd, True)
        If psColor.bCanceled = False Then
            Select Case .lstColorItems.List(.lstColorItems.ListIndex)
                Case "ActiveTabBackColor"
                    '   Set the Start color for the gradient for the ActiveTabBack
                    .txtStartColor.Text = pHexColorStr(psColor.oSelectedColor)
                    .ucXTab1.ActiveTabBackStartColor = CLng(psColor.oSelectedColor)
                Case "InActiveTabBackColor"
                    '   Set the Start color for the gradient for the InActiveTabBack
                    .txtStartColor.Text = pHexColorStr(psColor.oSelectedColor)
                    .ucXTab1.InActiveTabBackStartColor = CLng(psColor.oSelectedColor)
            End Select
            '   Paint the controls like CheckBoxes which do not
            '   have a transparent backstyle
            Call pSetCtrlBackColors
        End If
        .ucXTab1.SetFocus
    End With
End Sub

Private Sub cmdTabPicture_Click(Index As Integer)
    Dim psPicFile       As SelectedFile
    Dim sExt            As String
    
    With Me
        With FileDialog
            '   Set the filter list
            .sFilter = "Icons (*.ico)" & Chr$(0) & "*.ico" & Chr$(0) & "Bitmaps (*.bmp)" & Chr$(0) & "*.bmp"
            ' See Standard CommonDialog Flags for all options
            .flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY
            '   Set the Title of the dialog
            .sDlgTitle = "Please Select a Pitcure for Importing..."
            '   The init path
            .sInitDir = App.Path & "\Graphics\"
        End With
        '   How the open dialog
        psPicFile = ShowOpen(.hwnd, True)
        If psPicFile.bCanceled = False Then
            sExt = LCase(Right(psPicFile.sFiles(1), 3))
            Select Case sExt
                Case "ico", "bmp"
                    '   Set valid picture formats to the tab of choice...
                    Set .ucXTab1.TabPicture(.cmbTabPictureIndex.List(.cmbTabPictureIndex.ListIndex)) = LoadPicture(psPicFile.sFiles(1))
                    .txtTabPicture.Text = psPicFile.sFiles(1)
                Case Else
                    '   Ooops, you selected somthing other than these two...
                    MsgBox "                  The Picture Format Selected is Invalid" & vbCrLf & vbCrLf & "Please Select only Icon (*.ico) or Bitmap (*.bmp) file formats", vbExclamation, "ucXTab"
            End Select
        End If
        .ucXTab1.SetFocus
    End With
End Sub

Private Sub Form_Load()
    
    With Me
        
        '   Prevent recursion on controls...
        bLoading = True
        '   Set the version info label
        .lblVersion.Caption = "Version: " & .ucXTab1.Version(False)
        '   Init the tab captions
        With .ucXTab1
            .TabCaption(0) = "&Welcome"
            .TabCaption(1) = "&Settings"
            .TabCaption(2) = "&Colors"
            .TabCaption(3) = "&Pictures"
            .TabCaption(4) = "&Tools"
            .TabCaption(5) = "&Enable"
            .ActiveTab = 0
        End With
        '   Build the droplists
        For i = 0 To .ucXTab1.TabCount - 1
            If i <> 0 Then .cmbTabCount.AddItem i + 1
            .cmbActiveTab.AddItem i
            .cmbTabPictureIndex.AddItem i
        Next
        '   Set the previous number of Tabs
        lPrevCount = .ucXTab1.TabCount
        '   Set the init conditions
        .cmbTabCount.ListIndex = .cmbTabCount.ListCount - 1
        .cmbActiveTab.ListIndex = 0
        .cmbTabPictureIndex.ListIndex = 0
        '   Fill the TabStyle droplist
        With .cmbTabStyle
            .AddItem "xStyleTabbedDialog"
            .AddItem "xStylePropertyPages"
            .ListIndex = 0
        End With
        '   Fill the TabTheme droplist
        With .cmbTabTheme
            .AddItem "xThemeWin9X"
            .AddItem "xThemeWinXP"
            .AddItem "xThemeVisualStudio2003"
            .AddItem "xThemeRoundTabs"
            .AddItem "xThemeGalaxy"
            .ListIndex = 1
        End With
        '   Now get the values from the control and update the GUI
        .txtActiveTabHeight.Text = .ucXTab1.ActiveTabHeight
        .txtInActiveTabHeight.Text = .ucXTab1.InActiveTabHeight
        .txtXRadius.Text = .ucXTab1.XRadius
        .txtYRadius.Text = .ucXTab1.YRadius
        .txtToolTipText.Text = .ucXTab1.ToolTipText
        '   Set the Check List for the various options
        With lstChkItems
            .AddItem "UseControlBox"
            .AddItem "UseFocusRect"
            .AddItem "UseThemeDetection"
            .AddItem "UseMouseWheelScrolling"
            .Selected(0) = Abs(ucXTab1.UseControlBox)
            .Selected(1) = Abs(ucXTab1.ShowFocusRect)
            .Selected(2) = Abs(ucXTab1.UseThemeDetection)
            .Selected(3) = Abs(ucXTab1.UseMouseWheelScroll)
            .ListIndex = 0
        End With
        '   Fill the ColorItems ListBox
        With .lstColorItems
            .AddItem "ActiveTabBackColor"
            .AddItem "ActiveTabForeColor"
            .AddItem "BottomRightInnerBorderColor"
            .AddItem "DisabledTabBackColor"
            .AddItem "DisabledTabForeColor"
            .AddItem "Focused Color"
            .AddItem "ForeColor"
            .AddItem "HoverColor"
            .AddItem "InActiveTabBackColor"
            .AddItem "InActiveTabForeColor"
            .AddItem "Outer BorderColor"
            .AddItem "PictureMaskColor"
            .AddItem "TabStripBackColor"
            .AddItem "TopLeftInnerBorderColor"
            .ListIndex = 0
        End With
        '   Set the checkboxes to match the control
        .chkUseFocusedColor.Value = Abs(.ucXTab1.UseFocusedColor)
        .chkUseMaskColor.Value = Abs(.ucXTab1.UseMaskColor)
        '   Fill the PictureSize droplist
        With .cmbPictureSize
            .AddItem "xSizeSmall"
            .AddItem "xSizeLarge"
            .ListIndex = 0
        End With
        '   Fill the PictureAlign droplist
        With .cmbPictureAlign
            .AddItem "xAlignLeftEdge"
            .AddItem "xAlignRightEdge"
            .AddItem "xAlignLeftOfCaption"
            .AddItem "xAlignRightOfCaption"
            .ListIndex = 0
        End With
        With .cmbActionIndex
            For i = 0 To ucXTab1.TabCount - 1
                .AddItem i
            Next
            .ListIndex = 0
        End With
        '   Reset the flag
        bLoading = False
        '   Fire an event to set the rest of the controls
        Call lstColorItems_Click
        frmReSize.Show
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '   Unload the form
    Unload Me
End Sub

Public Function GetWebBrowserHandle(hWndHost As Long) As Long
    Dim lRet As Long
    Dim lResult As Long
    Dim hWndChild As Long
    Dim sClassString As String * 256
    '   Based on the following....
    'http://support.microsoft.com/kb/244310
    
    '   Enumerate the hard way to locate the WebBrowser controls hWnd
    hWndChild = GetWindow(hWndHost, GW_CHILD)
    While (lResult = 0) And (hWndChild <> 0)
        hWndChild = GetWindow(hWndChild, GW_HWNDNEXT)
        lRet = GetClassName(hWndChild, sClassString, 256)
        If hWndChild <> 0 Then
            lRet = GetClassName(hWndChild, sClassString, 256)
            If Left$(sClassString, InStr(sClassString, Chr$(0)) - 1) = "Shell Embedding" Then
                lResult = 1
            End If
        End If
    Wend
                
    GetWebBrowserHandle = hWndChild
End Function

Private Sub lblLink_Click()
    Dim OpenLink As String
    With Me
        '   Set the link color
        .lblLink.ForeColor = &HC000C0
        '   Launch the browser and follow the link
        OpenLink = ShellExecute(hwnd, "open", sLink, vbNull, vbNull, 1)
        .ucXTab1.SetFocus
    End With
End Sub

Private Sub lstChkItems_Click()
    With Me
        If Not bLoading Then
            Select Case .lstChkItems.ListIndex
                Case 0  'UseControlBox
                    .ucXTab1.UseControlBox = (.lstChkItems.Selected(.lstChkItems.ListIndex))
                Case 1  'UseFocusRect
                    .ucXTab1.ShowFocusRect = (.lstChkItems.Selected(.lstChkItems.ListIndex))
                Case 2  'UseThemeDetection
                    '   Are we using theme detection (i.e. Win9X or WinXP)?
                    If .lstChkItems.Selected(.lstChkItems.ListIndex) Then
                        .ucXTab1.UseThemeDetection = (.lstChkItems.Selected(.lstChkItems.ListIndex))
                        If .ucXTab1.IsWinXP Then
                            '   Store the value in case we need it for later
                            lPrevTheme = .cmbTabTheme.ListIndex
                            '   This is WinXP
                            .cmbTabTheme.ListIndex = 1
                        End If
                    Else
                        '   Roll back the changes
                        .cmbTabTheme.ListIndex = lPrevTheme
                        '   Call the event handler to update the GUI
                        Call cmbTabTheme_Click
                    End If
                    Call pSetCtrlBackColors
                Case 3  'UseMouseWheelScrolling
                    .ucXTab1.UseMouseWheelScroll = (.lstChkItems.Selected(.lstChkItems.ListIndex))
            End Select
            .ucXTab1.SetFocus
        End If
    End With
End Sub

Private Sub lstColorItems_Click()
    With Me
        If Not bLoading Then
            '   Fill the controls with the color values from the control
            Select Case .lstColorItems.List(.lstColorItems.ListIndex)
                Case "ActiveTabBackColor"
                    '   Dis/Enable the correct color selection box
                    Call pEnableColorCtrls(2)
                    .txtStartColor.Text = pHexColorStr(.ucXTab1.ActiveTabBackStartColor)
                    .txtEndColor.Text = pHexColorStr(.ucXTab1.ActiveTabBackEndColor)
                Case "ActiveTabForeColor"
                    '   Dis/Enable the correct color selection box
                    Call pEnableColorCtrls(1)
                    .txtItemColor.Text = pHexColorStr(.ucXTab1.ActiveTabForeColor)
                Case "BottomRightInnerBorderColor"
                    '   Dis/Enable the correct color selection box
                    Call pEnableColorCtrls(1)
                    .txtItemColor.Text = pHexColorStr(.ucXTab1.BottomRightInnerBorderColor)
                Case "DisabledTabBackColor"
                    '   Dis/Enable the correct color selection box
                    Call pEnableColorCtrls(1)
                    .txtItemColor.Text = pHexColorStr(.ucXTab1.DisabledTabBackColor)
                Case "DisabledTabForeColor"
                    '   Dis/Enable the correct color selection box
                    Call pEnableColorCtrls(1)
                    .txtItemColor.Text = pHexColorStr(.ucXTab1.DisabledTabForeColor)
                Case "Focused Color"
                    '   Dis/Enable the correct color selection box
                    Call pEnableColorCtrls(1)
                    .txtItemColor.Text = pHexColorStr(.ucXTab1.FocusedColor)
                Case "ForeColor"
                    '   Dis/Enable the correct color selection box
                    Call pEnableColorCtrls(1)
                    .txtItemColor.Text = pHexColorStr(.ucXTab1.ForeColor)
                Case "HoverColor"
                    '   Dis/Enable the correct color selection box
                    Call pEnableColorCtrls(1)
                    .txtItemColor.Text = pHexColorStr(.ucXTab1.HoverColor)
                Case "InActiveTabBackColor"
                    '   Dis/Enable the correct color selection box
                    Call pEnableColorCtrls(2)
                    .txtStartColor.Text = pHexColorStr(.ucXTab1.InActiveTabBackStartColor)
                    .txtEndColor.Text = pHexColorStr(.ucXTab1.InActiveTabBackEndColor)
                Case "InActiveTabForeColor"
                    '   Dis/Enable the correct color selection box
                    Call pEnableColorCtrls(1)
                    .txtItemColor.Text = pHexColorStr(.ucXTab1.InActiveTabForeColor)
                Case "Outer BorderColor"
                    '   Dis/Enable the correct color selection box
                    Call pEnableColorCtrls(1)
                    .txtItemColor.Text = pHexColorStr(.ucXTab1.OuterBorderColor)
                Case "PictureMaskColor"
                    '   Dis/Enable the correct color selection box
                    Call pEnableColorCtrls(1)
                    .txtItemColor.Text = pHexColorStr(.ucXTab1.PictureMaskColor)
                Case "TabStripBackColor"
                    '   Dis/Enable the correct color selection box
                    Call pEnableColorCtrls(1)
                    .txtItemColor.Text = pHexColorStr(.ucXTab1.TabStripBackColor)
                Case "TopLeftInnerBorderColor"
                    '   Dis/Enable the correct color selection box
                    Call pEnableColorCtrls(1)
                    .txtItemColor.Text = pHexColorStr(.ucXTab1.TopLeftInnerBorderColor)
            End Select
        End If
    End With
End Sub

Private Sub pBuildTabNames()
    Dim TabNum As Long
    
    With Me
        '   Reset the tab names in case they were changed by the
        '   TabCount property
        TabNum = .ucXTab1.TabCount
        .ucXTab1.TabCaption(0) = "Welcome"
        If TabNum >= 2 Then .ucXTab1.TabCaption(1) = "Settings"
        If TabNum >= 3 Then .ucXTab1.TabCaption(2) = "Colors"
        If TabNum >= 4 Then .ucXTab1.TabCaption(3) = "Pictures"
    End With
End Sub

Private Sub pEnableColorCtrls(lCtrlNum As Long, Optional bEnabled As Boolean = True)
    
    With Me
        '   Disable all the color controls
        .lblItemColor.Enabled = False
        .txtItemColor.Enabled = False
        .txtItemColor.Text = "Locate Color..."
        .cmdItemColor.Enabled = False
        .lblStartColor.Enabled = False
        .txtStartColor.Enabled = False
        .txtStartColor.Text = "Locate Color..."
        .cmdStartColor.Enabled = False
        .lblEndColor.Enabled = False
        .txtEndColor.Enabled = False
        .txtEndColor.Text = "Locate Color..."
        .cmdEndColor.Enabled = False
        '   Now only set the active ones...
        If lCtrlNum = 1 Then
            .lblItemColor.Enabled = bEnabled
            .txtItemColor.Enabled = bEnabled
            .cmdItemColor.Enabled = bEnabled
        Else
            .lblStartColor.Enabled = bEnabled
            .txtStartColor.Enabled = bEnabled
            .cmdStartColor.Enabled = bEnabled
            .lblEndColor.Enabled = bEnabled
            .txtEndColor.Enabled = bEnabled
            .cmdEndColor.Enabled = bEnabled
        End If
    End With
End Sub

Private Function pHexColorStr(lColor As Long) As String
    '   Get the Hex version of the color...
    pHexColorStr = UCase("&H" & Hex(lColor))
End Function

Private Sub pReBuildListValues()
    Dim i As Long
    With Me
        '   Clear the lists
        .cmbTabCount.Clear
        .cmbActiveTab.Clear
        .cmbTabPictureIndex.Clear
        '   Now rebuild them...
        For i = 0 To .ucXTab1.TabCount - 1
            If i <> 0 Then .cmbTabCount.AddItem i + 1
            .cmbActiveTab.AddItem i
            .cmbTabPictureIndex.AddItem i
        Next i
        .ucXTab1.SetFocus
    End With
End Sub

Private Sub pSelectText(TxtBox As TextBox)
    With TxtBox
        '   Select the text
        .SelStart = 0
        .SelLength = Len(TxtBox.Text)
    End With
End Sub

Private Sub pSetCtrlBackColors()
    Dim Ctl As Control
    On Error Resume Next
    For Each Ctl In Me.Controls
        If (TypeOf Ctl Is CheckBox) Or (TypeOf Ctl Is OptionButton) Then
            '   Checkbox/OptionsButton controls do not have a Transparent backstyle, so
            '   this routine sets the BackColor of the object to that of
            '   its host to give the illusion of Transparency...
            '
            '   This could be extended for any control which need to have its
            '   backcolor match the hosts...
            Ctl.BackColor = Me.ucXTab1.ActiveTabBackEndColor
        End If
    Next Ctl
End Sub

Private Sub pSetItemColor(lColor As Long)
    With Me
        Select Case .lstColorItems.List(.lstColorItems.ListIndex)
            Case "ActiveTabForeColor"
                '   Set the forecolor for the ActiveTab
                .txtItemColor.Text = pHexColorStr(lColor)
                .ucXTab1.ActiveTabForeColor = CLng(lColor)
            Case "BottomRightInnerBorderColor"
                '   Set the BorderColor for the BottomRightInner section of the ActiveTab
                .txtItemColor.Text = pHexColorStr(lColor)
                .ucXTab1.BottomRightInnerBorderColor = CLng(lColor)
            Case "DisabledTabBackColor"
                '   Set the BackColor for the DisabledTab
                .txtItemColor.Text = pHexColorStr(lColor)
                .ucXTab1.DisabledTabBackColor = CLng(lColor)
            Case "DisabledTabForeColor"
                '   Set the ForeColor for the DisabledTab
                .txtItemColor.Text = pHexColorStr(lColor)
                .ucXTab1.DisabledTabForeColor = CLng(lColor)
            Case "Focused Color"
                '   Set the FocusedColor for the ActiveTab
                '   (Used with WinXP Theme Only)
                .txtItemColor.Text = pHexColorStr(lColor)
                .ucXTab1.FocusedColor = CLng(lColor)
            Case "ForeColor"
                '   Set the ForeColor for the ActiveTab
                .txtItemColor.Text = pHexColorStr(lColor)
                .ucXTab1.ForeColor = CLng(lColor)
            Case "HoverColor"
                '   Set the HoverColor for the InActiveTabs
                .txtItemColor.Text = pHexColorStr(lColor)
                .ucXTab1.HoverColor = CLng(lColor)
            Case "InActiveTabForeColor"
                '   Set the ForeColor for the InActiveTabs
                .txtItemColor.Text = pHexColorStr(lColor)
                .ucXTab1.InActiveTabForeColor = CLng(lColor)
            Case "Outer BorderColor"
                '   Set the BorderColor for the Tabs
                .txtItemColor.Text = pHexColorStr(lColor)
                .ucXTab1.OuterBorderColor = CLng(lColor)
            Case "PictureMaskColor"
                '   Set the MaskColor for the Tab Bitmap Pictures
                .txtItemColor.Text = pHexColorStr(lColor)
                .ucXTab1.PictureMaskColor = CLng(lColor)
            Case "TabStripBackColor"
                '   Set the BackColor for the TabStrip
                .txtItemColor.Text = pHexColorStr(lColor)
                .ucXTab1.TabStripBackColor = CLng(lColor)
            Case "TopLeftInnerBorderColor"
                '   Set the BorderColor for the TopLeftInner edge of the Tabs
                .txtItemColor.Text = pHexColorStr(lColor)
                .ucXTab1.TopLeftInnerBorderColor = CLng(lColor)
        End Select
        '   Paint the controls like CheckBoxes which do not
        '   have a transparent backstyle
        Call pSetCtrlBackColors
        .ucXTab1.SetFocus
    End With

End Sub

Private Sub opDynamicType_Click(Index As Integer)
    With Me
        Select Case Index
            '   Toggle between Insert and Remove
            Case 0
                .lblDynamicTab.Caption = "AfterTab Index:"
                .cmdAction.Caption = "Insert"
                .lblTabCaption.Enabled = True
                .txtTabCaption.Enabled = True
                .cmdAddCtrls(1).Enabled = True
            Case 1
                .lblDynamicTab.Caption = "Remove Tab Index:"
                .cmdAction.Caption = "Remove"
                .lblTabCaption.Enabled = False
                .txtTabCaption.Enabled = False
                .cmdAddCtrls(1).Enabled = False
        End Select
    End With
End Sub

Private Sub txtActiveTabHeight_GotFocus()
    With Me
        '   Select the text for changing...
        Call pSelectText(.txtActiveTabHeight)
    End With
End Sub

Private Sub txtActiveTabHeight_KeyDown(KeyCode As Integer, Shift As Integer)
    With Me
        If (KeyCode = vbKeyReturn) Then
            '   Call the LostFocus Event Handler
            Call txtActiveTabHeight_LostFocus
        End If
    End With
End Sub

Private Sub txtActiveTabHeight_LostFocus()
    With Me
        If IsNumeric(.txtActiveTabHeight.Text) Then
            '   Set the ActiveTab height
            .ucXTab1.ActiveTabHeight = .txtActiveTabHeight.Text
        Else
            '   There was a problem, so select the text for the user
            Call pSelectText(txtActiveTabHeight)
            '   Warn them....
            MsgBox "The Value Selected is Invalid, Please Enter a Valid Numeric Value.", vbExclamation, "ucXTab"
        End If
    End With
End Sub

Private Sub txtEndColor_GotFocus()
    With Me
        '   Select the text for changing...
        Call pSelectText(.txtEndColor)
    End With
End Sub

Private Sub txtEndColor_KeyDown(KeyCode As Integer, Shift As Integer)
    With Me
        If (KeyCode = vbKeyReturn) Then
            '   Call the LostFocus Event Handler
            Call txtEndColor_LostFocus
        End If
    End With
End Sub

Private Sub txtEndColor_LostFocus()
    With Me
        If (.txtEndColor.Text <> "") And (IsNumeric(.txtEndColor.Text)) Then
            Select Case .lstColorItems.List(.lstColorItems.ListIndex)
                Case "ActiveTabBackColor"
                    '   Set the End color for the gradient for the ActiveTabBack
                    .txtEndColor.Text = pHexColorStr(.txtEndColor.Text)
                    .ucXTab1.ActiveTabBackEndColor = CLng(.txtEndColor.Text)
                Case "InActiveTabBackColor"
                    '   Set the End color for the gradient for the InActiveTabBack
                    .txtEndColor.Text = pHexColorStr(.txtEndColor.Text)
                    .ucXTab1.InActiveTabBackEndColor = CLng(.txtEndColor.Text)
            End Select
            '   Paint the controls like CheckBoxes which do not
            '   have a transparent backstyle
            Call pSetCtrlBackColors
        Else
            MsgBox "The Value Entered is Invalid!", vbExclamation + vbOKOnly, "ucXTab"
        End If
    End With
End Sub

Private Sub txtInActiveTabHeight_GotFocus()
    With Me
        '   Select the text for changing...
        Call pSelectText(.txtInActiveTabHeight)
    End With
End Sub

Private Sub txtInActiveTabHeight_KeyDown(KeyCode As Integer, Shift As Integer)
    With Me
        If (KeyCode = vbKeyReturn) Then
            '   Call the LostFocus Event Handler
            Call txtInActiveTabHeight_LostFocus
        End If
    End With
End Sub

Private Sub txtInActiveTabHeight_LostFocus()
    With Me
        If IsNumeric(.txtInActiveTabHeight.Text) Then
            '   Set the InActiveTab height
            .ucXTab1.InActiveTabHeight = .txtInActiveTabHeight.Text
        Else
            '   There was a problem, so select the text for the user
            Call pSelectText(txtInActiveTabHeight)
            '   Warn them....
            MsgBox "The Value Selected is Invalid, Please Enter a Valid Numeric Value.", vbExclamation, "ucXTab"
        End If
    End With
End Sub

Private Sub txtItemColor_GotFocus()
    With Me
        '   Select the text for changing...
        Call pSelectText(.txtItemColor)
    End With
End Sub

Private Sub txtItemColor_KeyDown(KeyCode As Integer, Shift As Integer)
    With Me
        If (KeyCode = vbKeyReturn) Then
            '   Call the LostFocus Event Handler
            Call txtItemColor_LostFocus
        End If
    End With
End Sub

Private Sub txtItemColor_LostFocus()
    With Me
        If (.txtItemColor.Text <> "") And (IsNumeric(.txtItemColor.Text)) Then
            Call pSetItemColor(CLng(.txtItemColor.Text))
        Else
            MsgBox "The Value Entered is Invalid!", vbExclamation + vbOKOnly, "ucXTab"
        End If
    End With
End Sub

Private Sub txtStartColor_GotFocus()
    With Me
        '   Select the text for changing...
        Call pSelectText(.txtStartColor)
    End With
End Sub

Private Sub txtStartColor_KeyDown(KeyCode As Integer, Shift As Integer)
    With Me
        If (KeyCode = vbKeyReturn) Then
            '   Call the LostFocus Event Handler
            Call txtStartColor_LostFocus
        End If
    End With
End Sub

Private Sub txtStartColor_LostFocus()
    With Me
        If (.txtStartColor.Text <> "") And (IsNumeric(.txtStartColor.Text)) Then
            Select Case .lstColorItems.List(.lstColorItems.ListIndex)
                Case "ActiveTabBackColor"
                    '   Set the Start color for the gradient for the ActiveTabBack
                    .txtStartColor.Text = pHexColorStr(.txtStartColor.Text)
                    .ucXTab1.ActiveTabBackStartColor = CLng(.txtStartColor.Text)
                Case "InActiveTabBackColor"
                    '   Set the Start color for the gradient for the InActiveTabBack
                    .txtStartColor.Text = pHexColorStr(.txtStartColor.Text)
                    .ucXTab1.InActiveTabBackStartColor = CLng(.txtStartColor.Text)
            End Select
            '   Paint the controls like CheckBoxes which do not
            '   have a transparent backstyle
            Call pSetCtrlBackColors
        Else
            MsgBox "The Value Entered is Invalid!", vbExclamation + vbOKOnly, "ucXTab"
        End If
    End With

End Sub

Private Sub txtToolTipText_GotFocus()
    With Me
        '   Select the text for changing...
        Call pSelectText(.txtToolTipText)
    End With
End Sub

Private Sub txtToolTipText_KeyDown(KeyCode As Integer, Shift As Integer)
    With Me
        If (KeyCode = vbKeyReturn) Then
            '   Call the LostFocus Event Handler
            Call txtToolTipText_LostFocus
        End If
    End With
End Sub

Private Sub txtToolTipText_LostFocus()
    With Me
        '   Set the ToolTipText
        .ucXTab1.ToolTipText = .txtToolTipText.Text
    End With
End Sub

Private Sub txtXRadius_GotFocus()
    With Me
        '   Select the text for changing...
        Call pSelectText(.txtXRadius)
    End With
End Sub

Private Sub txtXRadius_KeyDown(KeyCode As Integer, Shift As Integer)
    With Me
        If (KeyCode = vbKeyReturn) Then
            '   Call the LostFocus Event Handler
            Call txtXRadius_LostFocus
        End If
    End With
End Sub

Private Sub txtXRadius_LostFocus()
    With Me
        If IsNumeric(.txtXRadius.Text) Then
            '   Set the XRadius value
            .ucXTab1.XRadius = .txtXRadius.Text
        Else
            '   There was a problem, so select the text for the user
            Call pSelectText(txtXRadius)
            '   Warn them....
            MsgBox "The Value Selected is Invalid, Please Enter a Valid Numeric Value.", vbExclamation, "ucXTab"
        End If
    End With
End Sub

Private Sub txtYRadius_GotFocus()
    With Me
        '   Select the text for changing...
        Call pSelectText(.txtYRadius)
    End With
End Sub

Private Sub txtYRadius_KeyDown(KeyCode As Integer, Shift As Integer)
    With Me
        If (KeyCode = vbKeyReturn) Then
            '   Call the LostFocus Event Handler
            Call txtYRadius_LostFocus
        End If
    End With
End Sub

Private Sub txtYRadius_LostFocus()
    With Me
        If IsNumeric(.txtYRadius.Text) Then
            '   Set the YRadius value
            .ucXTab1.YRadius = .txtYRadius.Text
        Else
            '   There was a problem, so select the text for the user
            Call pSelectText(txtYRadius)
            '   Warn them....
            MsgBox "The Value Selected is Invalid, Please Enter a Valid Numeric Value.", vbExclamation, "ucXTab"
        End If
    End With
End Sub

Private Sub ucXTab1_AfterCompleteInit()
    '   We are done loading the control....
    Debug.Print "ucXTab Loading Complete!"
End Sub

Private Sub ucXTab1_BeforeTabSwitch(ByVal iNewActiveTab As Integer, bCancel As Boolean)
    '   We are changing Tabs...
    Debug.Print "New Tab Number: " & iNewActiveTab
End Sub

Private Sub ucXTab1_Click()
    Debug.Print "Mouse Click"
End Sub

Private Sub ucXTab1_ControlBoxEnter()
    Debug.Print "ControlBox Enter"
End Sub

Private Sub ucXTab1_ControlBoxExit()
    Debug.Print "ControlBox Exit"
End Sub

Private Sub ucXTab1_ControlBoxHover(x As Single, y As Single)
    Debug.Print "ControlBoxHover: X: " & x, "Y: " & y
End Sub

Private Sub ucXTab1_ControlBoxMouseDown(x As Single, y As Single)
    Debug.Print "ControlBoxMouseDown: X: " & x, "Y: " & y
End Sub

Private Sub ucXTab1_ControlBoxMouseUp(x As Single, y As Single)
    Debug.Print "ControlBoxMouseUp: X: " & x, "Y: " & y
End Sub

Private Sub ucXTab1_DblClick()
    Debug.Print "Mouse DblClick"
End Sub

Private Sub ucXTab1_GotFocus()
    Debug.Print "Got Focus"
End Sub

Private Sub ucXTab1_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug.Print "KeyDown: " & KeyCode
End Sub

Private Sub ucXTab1_KeyPress(KeyAscii As Integer)
    Debug.Print "KeyPress: " & KeyAscii
End Sub

Private Sub ucXTab1_LostFocus()
    Debug.Print "Lost Focus"
End Sub

Private Sub ucXTab1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print "MouseDown: X: " & x, "Y: " & y
End Sub

Private Sub ucXTab1_MouseEnter()
    Debug.Print "MouseEnter"
End Sub

Private Sub ucXTab1_MouseHover(ActiveTab As Long, x As Single, y As Single)
    '   Which tab are we on?
    Debug.Print "MouseHover @ ActiveTab: " & ActiveTab, "X: " & x, "Y: " & y
End Sub

Private Sub ucXTab1_MouseLeave()
    Debug.Print "MouseLeave"
End Sub

Private Sub ucXTab1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '   Show them where we are....
    Debug.Print "MouseMove: X: " & x, "Y: " & y
End Sub

Private Sub ucXTab1_MouseScrollDown()
    '   Tab Changed by MouseWheel Scrolling
    Debug.Print "MouseScrollDown"
End Sub

Private Sub ucXTab1_MouseScrollUp()
    '   Tab Changed by MouseWheel Scrolling
    Debug.Print "MouseScrollUp"
End Sub

Private Sub ucXTab1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print "MouseUp: X: " & x, "Y: " & y
End Sub

Private Sub ucXTab1_Status(ByVal sStatus As String)
    Debug.Print "Status: " & sStatus
End Sub

Private Sub ucXTab1_TabInsert(AfterTabIndex As Long)
    Debug.Print "TabInsert: " & AfterTabIndex
End Sub

Private Sub ucXTab1_TabRemove(TabIndex As Long)
    Debug.Print "TabRemove: " & TabIndex
End Sub

Private Sub ucXTab1_TabSwitch(ByVal lLastActiveTab As Integer)
    On Error Resume Next
    With Me
        If Not bLoading Then
            '   Set the droplist to reflect the changed tab
            .cmbActiveTab.ListIndex = .ucXTab1.ActiveTab
            .ucXTab1.SetFocus
        End If
    End With
End Sub


