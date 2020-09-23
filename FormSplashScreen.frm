VERSION 5.00
Begin VB.Form FormSplashScreen 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7320
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FormSplashScreen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   125
      Left            =   6480
      Top             =   2820
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   255
      Left            =   6600
      TabIndex        =   3
      Top             =   60
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      X1              =   480
      X2              =   4980
      Y1              =   3900
      Y2              =   3900
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   480
      X2              =   4980
      Y1              =   3780
      Y2              =   3780
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   0
      X1              =   480
      X2              =   4980
      Y1              =   3660
      Y2              =   3660
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Loading"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4260
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Complex Mathematician"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   480
      TabIndex        =   2
      Top             =   2760
      Width           =   5535
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Version 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      TabIndex        =   1
      Top             =   3480
      Width           =   1080
   End
   Begin VB.Label lblCopyright 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      Caption         =   "Copyright: IJR - 2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5340
      TabIndex        =   0
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Image imgLogo 
      Height          =   2385
      Left            =   780
      Picture         =   "FormSplashScreen.frx":000C
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2235
   End
End
Attribute VB_Name = "FormSplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim counter

Private Sub Command1_Click()
    Unload Me
    Load FormComplexMath
    FormComplexMath.Show
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
Timer1.Enabled = True

'Start = Second(Now)
'1
'If Second(Now) - Start >= 4 Then
'    Call Command1_Click
'Else
'    DoEvents
'    GoTo 1
'End If
End Sub

Private Sub Timer1_Timer()
counter = counter + 1
Dot$ = ""
For Counts = 1 To counter
    Dot$ = Dot$ + "."
Next
Label1.Caption = "Loading" & Dot$
If counter = 12 Then
    Timer1.Enabled = False
    Call Command1_Click
    Unload Me
End If
End Sub
