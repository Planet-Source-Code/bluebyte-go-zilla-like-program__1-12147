VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHelp 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4740
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "frmHelp.frx":0000
      Top             =   525
      Width           =   5685
   End
   Begin VB.Label cmdOk 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   4410
      MouseIcon       =   "frmHelp.frx":018F
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   5430
      Width           =   1275
   End
   Begin VB.Label lblHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "Blue Byte Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   525
      TabIndex        =   0
      Top             =   105
      Width           =   4950
   End
   Begin VB.Label lblHeader 
      BackStyle       =   0  'Transparent
      Caption         =   "Blue Byte Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   1
      Left            =   570
      TabIndex        =   1
      Top             =   120
      Width           =   4845
   End
   Begin VB.Image lblButton 
      Height          =   360
      Index           =   1
      Left            =   4410
      Picture         =   "frmHelp.frx":0499
      Top             =   5355
      Width           =   1335
   End
   Begin VB.Image imgMin 
      Height          =   225
      Left            =   5460
      MouseIcon       =   "frmHelp.frx":1DFD
      MousePointer    =   99  'Custom
      ToolTipText     =   "Minimize Window"
      Top             =   50
      Width           =   225
   End
   Begin VB.Image imgClose 
      Height          =   225
      Left            =   5670
      MouseIcon       =   "frmHelp.frx":2107
      MousePointer    =   99  'Custom
      ToolTipText     =   "Click here to Exit this Window"
      Top             =   50
      Width           =   225
   End
   Begin VB.Image imgBackDrop 
      Height          =   6000
      Left            =   0
      Picture         =   "frmHelp.frx":2411
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub imgBackDrop_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub imgClose_Click()
Unload Me
End Sub
