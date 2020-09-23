VERSION 5.00
Begin VB.Form FormAboutComplexMath 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Complex Mathematician"
   ClientHeight    =   2205
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "FormAboutComplexMath.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1521.93
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "FormAboutComplexMath.frx":030A
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4200
      TabIndex        =   0
      Top             =   1800
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   56.343
      X2              =   5281.227
      Y1              =   1159.566
      Y2              =   1159.566
   End
   Begin VB.Label lblDescription 
      Caption         =   "Written, Conceived and Copywritten by Ian J. Renwick, 2000.                e-mail: Soze99@.com. All comments welcome."
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   840
      TabIndex        =   2
      Top             =   1125
      Width           =   4545
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Complex Mathematician"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   1080
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Caption         =   "Version 2.001.000a"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Top             =   780
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   56.343
      X2              =   5267.141
      Y1              =   1159.566
      Y2              =   1159.566
   End
End
Attribute VB_Name = "FormAboutComplexMath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    'Me.Caption = "About " & App.Title
    'lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblTitle.Caption = App.Title
End Sub
