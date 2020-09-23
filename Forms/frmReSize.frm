VERSION 5.00
Begin VB.Form frmReSize 
   Caption         =   "ReSize TestHarness"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin prjucXTab.ucXTab ucXTab1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4683
      TabCount        =   5
      TabCaption(0)   =   "Tab 0"
      TabCaption(1)   =   "Tab 1"
      TabCaption(2)   =   "Tab 2"
      TabCaption(3)   =   "Tab 3"
      TabCaption(4)   =   "Tab 4"
      ActiveTab       =   2
      ActiveTabBackEndColor=   16514555
      ActiveTabBackStartColor=   16514555
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BottomRightInnerBorderColor=   10070188
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   10526880
      Enabled         =   0   'False
      ForeColor       =   10526880
      InActiveTabBackEndColor=   14204606
      InActiveTabBackStartColor=   16777215
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   10198161
      TabStyle        =   1
      TabTheme        =   1
      TopLeftInnerBorderColor=   16777215
      TabOffset       =   50000
      Begin VB.PictureBox PictureBox1 
         Height          =   2055
         Left            =   1.00120e5
         ScaleHeight     =   1995
         ScaleWidth      =   3915
         TabIndex        =   4
         Top             =   420
         Width           =   3975
         Begin VB.Label Label1 
            Caption         =   $"frmReSize.frx":0000
            Height          =   975
            Left            =   720
            TabIndex        =   6
            Top             =   480
            Width           =   2655
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Index           =   2
         Left            =   2280
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Index           =   1
         Left            =   51200
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Index           =   0
         Left            =   1.00120e5
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Caption         =   "Dynamically adjust the width and height of the form and notice that the button controls placement is maintained!!!"
         Height          =   735
         Left            =   1.00600e5
         TabIndex        =   5
         Top             =   1080
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmReSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Sub Form_Resize()
'    With Me
'        .ucXTab1.Left = 120
'        .ucXTab1.Top = 120
'        If (.ScaleWidth - (.ucXTab1.Left * 2)) > 0 Then
'            .ucXTab1.Width = .ScaleWidth - (.ucXTab1.Left * 2)
'        End If
'        If (.ScaleHeight - (.ucXTab1.Top * 2)) > 0 Then
'            .ucXTab1.Height = .ScaleHeight - (.ucXTab1.Top * 2)
'        End If
'        .ucXTab1.Refresh
'    End With
'End Sub
Private Sub Form_Resize()

    With Me
        If (.ScaleWidth - (.ucXTab1.Left * 2)) > 0 Then
            .ucXTab1.Width = .ScaleWidth - (.ucXTab1.Left * 2)
        End If
        If (.ScaleHeight - (.ucXTab1.Top * 2)) > 0 Then
            .ucXTab1.Height = .ScaleHeight - .ucXTab1.Top - 160 '(.ucXTab1.Top * 2)
        End If
        .ucXTab1.Refresh
    End With
        
    With PictureBox1
        .Height = ucXTab1.Height - PictureBox1.Top - 120 '- 240
        .Width = ucXTab1.Width - 240
    End With
    
End Sub

