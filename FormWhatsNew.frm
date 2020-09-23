VERSION 5.00
Begin VB.Form FormWhatsNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "What's New In Version 2.0"
   ClientHeight    =   3045
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "FormWhatsNew.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2101.713
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "New To Version 2"
      Height          =   975
      Left            =   300
      TabIndex        =   4
      Top             =   1260
      Width           =   5175
      Begin VB.Label lblDescription 
         Caption         =   $"FormWhatsNew.frx":030A
         ForeColor       =   &H00000000&
         Height          =   630
         Left            =   180
         TabIndex        =   5
         Top             =   300
         Width           =   4785
      End
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "FormWhatsNew.frx":03F4
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
      Top             =   2625
      Width           =   1260
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Version 2.001.000a"
      Height          =   225
      Left            =   1080
      TabIndex        =   3
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label Label1 
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
      Left            =   1050
      TabIndex        =   2
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.582
      Y2              =   1687.582
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
End
Attribute VB_Name = "FormWhatsNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  Unload Me
End Sub

