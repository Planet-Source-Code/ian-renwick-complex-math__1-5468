VERSION 5.00
Begin VB.Form FormHelp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help"
   ClientHeight    =   4530
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "FormHelp.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3126.686
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4230
      TabIndex        =   2
      Top             =   4065
      Width           =   1260
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   270
      Picture         =   "FormHelp.frx":030A
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.Frame Frame1 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   330
      TabIndex        =   0
      Top             =   1260
      Width           =   5175
      Begin VB.TextBox Text1 
         Height          =   2055
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "FormHelp.frx":0614
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5337.57
      Y1              =   2650.436
      Y2              =   2650.436
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
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Width           =   3885
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Version 2.001.000a"
      Height          =   225
      Left            =   1110
      TabIndex        =   3
      Top             =   780
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   126.772
      X2              =   5337.57
      Y1              =   2650.436
      Y2              =   2650.436
   End
End
Attribute VB_Name = "FormHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub


