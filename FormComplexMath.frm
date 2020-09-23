VERSION 5.00
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Begin VB.Form FormComplexMath 
   BackColor       =   &H80000004&
   Caption         =   "Complex Mathematician v2.0"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8595
   Icon            =   "FormComplexMath.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CommandCalculate 
      Caption         =   "CALCULATE !"
      Height          =   495
      Left            =   3240
      TabIndex        =   22
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton CommandSC1AC2 
      BackColor       =   &H8000000B&
      Caption         =   "    Swap 'Complex 1'            and 'Complex 2'"
      Height          =   495
      Left            =   5340
      TabIndex        =   21
      Top             =   4140
      Width           =   1815
   End
   Begin VB.CommandButton CommandRC2WR 
      BackColor       =   &H8000000B&
      Caption         =   "  Replace 'Complex 2'            with 'Result'"
      Height          =   495
      Left            =   3240
      TabIndex        =   20
      Top             =   4140
      Width           =   1815
   End
   Begin MSChartLib.MSChart Chart 
      Height          =   3015
      Left            =   4860
      OleObjectBlob   =   "FormComplexMath.frx":030A
      TabIndex        =   0
      Top             =   240
      Width           =   3675
   End
   Begin VB.CommandButton CommandRC1WR 
      BackColor       =   &H8000000B&
      Caption         =   "  Replace 'Complex 1'            with 'Result'"
      Height          =   495
      Left            =   1200
      TabIndex        =   19
      Top             =   4140
      Width           =   1815
   End
   Begin VB.Frame FrameRule 
      BackColor       =   &H80000004&
      Caption         =   "Rule"
      Height          =   615
      Left            =   60
      TabIndex        =   12
      Top             =   3300
      Width           =   8475
      Begin VB.Label LabelRule 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   13
         Top             =   240
         Width           =   8355
      End
   End
   Begin VB.Frame FrameResult 
      BackColor       =   &H80000004&
      Caption         =   "Result"
      Height          =   1035
      Left            =   60
      TabIndex        =   9
      Top             =   2220
      Width           =   4635
      Begin VB.TextBox TextResultImaginary 
         Height          =   315
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "0"
         Top             =   420
         Width           =   2175
      End
      Begin VB.TextBox TextResultReal 
         Height          =   315
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0"
         Top             =   420
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000004&
         Caption         =   "                      Real                                       Imaginary"
         Height          =   195
         Left            =   60
         TabIndex        =   18
         Top             =   780
         Width           =   4515
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF00FF&
         BorderWidth     =   2
         X1              =   60
         X2              =   4560
         Y1              =   300
         Y2              =   300
      End
   End
   Begin VB.Frame FrameComplex2 
      BackColor       =   &H80000004&
      Caption         =   "Complex 2"
      Height          =   1335
      Left            =   2460
      TabIndex        =   2
      Top             =   180
      Width           =   2235
      Begin VB.TextBox TextComplex2d 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1140
         TabIndex        =   6
         Text            =   "3"
         Top             =   720
         Width           =   1035
      End
      Begin VB.TextBox TextComplex2c 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   60
         TabIndex        =   5
         Text            =   "2"
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000004&
         Caption         =   "        Real             Imaginary"
         Height          =   195
         Left            =   60
         TabIndex        =   17
         Top             =   1080
         Width           =   2115
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   60
         X2              =   2160
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  c         +       i * d"
         Height          =   255
         Left            =   60
         TabIndex        =   8
         Top             =   240
         Width           =   2115
      End
   End
   Begin VB.Frame FrameComplex1 
      BackColor       =   &H80000004&
      Caption         =   "Complex 1"
      Height          =   1335
      Left            =   60
      TabIndex        =   1
      Top             =   180
      Width           =   2235
      Begin VB.TextBox TextComplex1b 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1140
         TabIndex        =   4
         Text            =   "1"
         Top             =   720
         Width           =   1035
      End
      Begin VB.TextBox TextComplex1a 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   60
         TabIndex        =   3
         Text            =   "1"
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000004&
         Caption         =   "        Real             Imaginary"
         Height          =   195
         Left            =   60
         TabIndex        =   16
         Top             =   1080
         Width           =   2115
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Index           =   0
         X1              =   60
         X2              =   2160
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  a         +       i * b"
         Height          =   255
         Left            =   60
         TabIndex        =   7
         Top             =   240
         Width           =   2115
      End
   End
   Begin VB.Frame FrameSelectAFunction 
      BackColor       =   &H80000004&
      Caption         =   "Select A Function"
      Height          =   615
      Left            =   60
      TabIndex        =   14
      Top             =   1560
      Width           =   3075
      Begin VB.ComboBox ComboFunction 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Text            =   "Select Function"
         Top             =   240
         Width           =   2835
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   8580
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line4 
      X1              =   60
      X2              =   8520
      Y1              =   4020
      Y2              =   4020
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   2
      X1              =   15
      X2              =   8580
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Menu MenuFile 
      Caption         =   "File"
      NegotiatePosition=   3  'Right
      Begin VB.Menu MenuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MenuFilePreset 
      Caption         =   "Preset Functions"
      Begin VB.Menu MenuFileE 
         Caption         =   "e ^ (pi * i) = -1"
      End
      Begin VB.Menu MenuFileI 
         Caption         =   "i ^ i = e ^ -(pi / 2)"
      End
      Begin VB.Menu MenuFileL1 
         Caption         =   "Log(-1) = pi * i"
      End
      Begin VB.Menu MenuFileL2 
         Caption         =   "Log(i) = i * pi / 2"
      End
   End
   Begin VB.Menu MenuHelp1 
      Caption         =   "Help"
      Begin VB.Menu MenuHelpHelp 
         Caption         =   "Help "
      End
      Begin VB.Menu MenuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuHelpWhatsNew 
         Caption         =   "What's New In Version 2.0"
      End
      Begin VB.Menu MenuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAbout1 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "FormComplexMath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim counter As Integer
Dim Response As String
Dim RealMemory As Double
Dim ImaginaryMemory As Double
Dim ResultReal As Double
Dim ResultImaginary As Double
Dim pi As Double

Private Sub CommandCalculate_Click()
CommandCalculate.Caption = "Done"
ResetChartAndBlueLine
Select Case ComboFunction
Case Is = "Magnitude of Complex 1"
    Call ComplexMagnitude(Val(TextComplex1a), Val(TextComplex1b), ResultReal)
    ResultImaginary = 0
Case Is = "Addition"
    Call ComplexAddition(Val(TextComplex1a), Val(TextComplex1b), Val(TextComplex2c), Val(TextComplex2d), ResultReal, ResultImaginary)
Case Is = "Subtraction"
    Call ComplexSubtraction(Val(TextComplex1a), Val(TextComplex1b), Val(TextComplex2c), Val(TextComplex2d), ResultReal, ResultImaginary)
Case Is = "Multiplication"
    Call ComplexMultiplication(Val(TextComplex1a), Val(TextComplex1b), Val(TextComplex2c), Val(TextComplex2d), ResultReal, ResultImaginary)
Case Is = "Division"
    If Val(TextComplex2c) = 0 And Val(TextComplex2d) = 0 Then
        Response = MsgBox("You Cannot Divide By <0,0>", 48, "Division Error")
        Exit Sub
    End If
    Call ComplexDivision(Val(TextComplex1a), Val(TextComplex1b), Val(TextComplex2c), Val(TextComplex2d), ResultReal, ResultImaginary)
Case Is = "e ^ (Complex 1)"
    Call ComplexExponentiation(Val(TextComplex1a), Val(TextComplex1b), ResultReal, ResultImaginary)
Case Is = "(Complex 1) ^ (Complex 2)"
    Call ComplexPower(Val(TextComplex1a), Val(TextComplex1b), Val(TextComplex2c), Val(TextComplex2d), ResultReal, ResultImaginary)
    If Val(TextComplex1a) = 0 And Val(TextComplex1b) = 0 And Val(TextComplex2d) = 0 And Val(TextComplex2c) > 0 Then
        ResultReal = 0
        ResultImaginary = 0
    End If
    If Val(TextComplex1a) = 0 And Val(TextComplex1b) = 0 And Val(TextComplex2d) = 0 And Val(TextComplex2c) = 0 Then
        ResultReal = 1
        ResultImaginary = 0
    End If
    If Val(TextComplex1a) = 0 And Val(TextComplex1b) = 0 And Val(TextComplex2d) = 0 And Val(TextComplex2c) < 0 Then
        Response = MsgBox("You Cannot Raise <0,0> To A Negative Power", 48, "Power Error")
        Exit Sub
    End If
Case Is = "Log (Complex 1)"
    If Val(TextComplex1a) = 0 And Val(TextComplex1b) = 0 Then
        Response = MsgBox("Log <0,0> Is Undefined", 48, "Logarithm Error")
        Exit Sub
    End If
    Call ComplexLog(Val(TextComplex1a), Val(TextComplex1b), ResultReal, ResultImaginary)
Case Is = "Log (Complex 1) [base (Complex 2)]"
    If Val(TextComplex1a) = 0 And Val(TextComplex1b) = 0 Then
        Response = MsgBox("Log <0,0> Is Undefined In Any Base", 48, "Logarithm Error")
        Exit Sub
    End If
    If Val(TextComplex2c) = 0 And Val(TextComplex2d) = 0 Then
        Response = MsgBox("Log (Base <0,0>) Is Undefined", 48, "Logarithm Error")
        Exit Sub
    End If
    If Val(TextComplex2c) = 1 And Val(TextComplex2d) = 0 Then
        Response = MsgBox("Log (Base <1,0>) Is Undefined", 48, "Logarithm Error")
        Exit Sub
    End If
    Call ComplexLogWithSpecialBase(Val(TextComplex1a), Val(TextComplex1b), Val(TextComplex2c), Val(TextComplex2d), ResultReal, ResultImaginary)
Case Is = "Sin (Complex 1)"
    Call ComplexSine(Val(TextComplex1a), Val(TextComplex1b), ResultReal, ResultImaginary)
Case Is = "Cos (Complex 1)"
    Call ComplexCosine(Val(TextComplex1a), Val(TextComplex1b), ResultReal, ResultImaginary)
Case Is = "Tan (Complex 1)"
    Call ComplexTangent(Val(TextComplex1a), Val(TextComplex1b), ResultReal, ResultImaginary)
End Select
UpdateDisplay
End Sub

Private Sub ComboFunction_Click()
CommandCalculate.Caption = "CALCULATE !"
ResetChartAndBlueLine
UpdateDisplay
End Sub

Private Sub CommandRC1WR_Click()
CommandCalculate.Caption = "CALCULATE !"
TextComplex1a = Val(TextResultReal)
TextComplex1b = Val(TextResultImaginary)
PutTextInChart
End Sub

Private Sub CommandRC2WR_Click()
CommandCalculate.Caption = "CALCULATE !"
TextComplex2c = Val(TextResultReal)
TextComplex2d = Val(TextResultImaginary)
PutTextInChart
End Sub

Private Sub CommandSC1AC2_Click()
CommandCalculate.Caption = "CALCULATE !"
RealMemory = Val(TextComplex1a)
ImaginaryMemory = Val(TextComplex1b)
TextComplex1a = Val(TextComplex2c)
TextComplex1b = Val(TextComplex2d)
TextComplex2c = RealMemory
TextComplex2d = ImaginaryMemory
PutTextInChart
End Sub

Public Sub Form_Load()

With ComboFunction
    .Clear
    .AddItem "Addition"
    .AddItem "Subtraction"
    .AddItem "Multiplication"
    .AddItem "Division"
    .AddItem "Magnitude of Complex 1"
    .AddItem "e ^ (Complex 1)"
    .AddItem "(Complex 1) ^ (Complex 2)"
    .AddItem "Log (Complex 1)"
    .AddItem "Log (Complex 1) [base (Complex 2)]"
    .AddItem "Sin (Complex 1)"
    .AddItem "Cos (Complex 1)"
    .AddItem "Tan (Complex 1)"
    .ListIndex = 0
End With
With Chart
    .RowCount = 2
    .ColumnCount = 6
    .Row = 1
End With

For counter = 1 To 6
    With Chart
        .Column = counter
        .Data = 0
    End With
Next

LabelRule = "(a+i*b)+(c+i*d) = (a+c)+i*(b+d)"
PutTextInChart

End Sub

Public Sub UpdateDisplay()
Select Case ComboFunction
Case Is = "Magnitude of Complex 1"
    LabelRule = "|a+i*b| = sqrt(a^2+b^2)"
    ClearBlueLine
    PutTextInChartWithoutComplex2
Case Is = "Addition"
    LabelRule = "(a+i*b)+(c+i*d) = (a+c)+i*(b+d)"
    PutTextInChart
Case Is = "Subtraction"
    LabelRule = "(a+i*b)-(c+i*d) = (a-c)+i*(b-d)"
    PutTextInChart
Case Is = "Multiplication"
    LabelRule = "(a+i*b)*(c+i*d) = a*c-b*d + i*(a*d+b*c)"
    PutTextInChart
Case Is = "Division"
    LabelRule = "(a+i*b)/(c+i*d) = [(a+i*b)*(c-i*d)]/(c^2+d^2)"
    PutTextInChart
Case Is = "e ^ (Complex 1)"
    LabelRule = "e ^ (a+i*b) = (e ^ a)[Cos(b)+i*Sin(b)]"
    ClearBlueLine
    PutTextInChartWithoutComplex2
Case Is = "(Complex 1) ^ (Complex 2)"
    LabelRule = "(a+i*b)^(c+i*d) = e ^ [(c+i*d)*log(a+i*b)]"
    PutTextInChart
Case Is = "Log (Complex 1)"
    LabelRule = "log(|a+i*b|) + i*arctan(b/a)"
    ClearBlueLine
    PutTextInChartWithoutComplex2
Case Is = "Log (Complex 1) [base (Complex 2)]"
    LabelRule = "Log (Complex 1) [base (Complex 2)] = [Log (Complex 1)]/[Log (Complex 2)]"
    PutTextInChart
Case Is = "Sin (Complex 1)"
    LabelRule = "Sin(a+i*b) = {e^[i*(a+i*b)]-e^[-i*(a+i*b)]}/(2*i)"
    ClearBlueLine
    PutTextInChartWithoutComplex2
Case Is = "Cos (Complex 1)"
    LabelRule = "Cos(a+i*b) = {e^[i*(a+i*b)]+e^[-i*(a+i*b)]}/i"
    ClearBlueLine
    PutTextInChartWithoutComplex2
Case Is = "Tan (Complex 1)"
    LabelRule = "Tan(a+i*b) = Sin(a+i*b)/Cos(a+i*b)"
    ClearBlueLine
    PutTextInChartWithoutComplex2
End Select
End Sub

Public Sub PutTextInChart()
TextResultReal = ResultReal
TextResultImaginary = ResultImaginary
With Chart
    .Row = 2
    .Column = 1
    .Data = TextComplex1a.Text
    .Column = 2
    .Data = TextComplex1b.Text
End With
If FrameComplex2.Enabled = True Then
    With Chart
        .Column = 3
        .Data = TextComplex2c.Text
        .Column = 4
        .Data = TextComplex2d.Text
    End With
End If
With Chart
    .Column = 5
    .Data = TextResultReal.Text
    .Column = 6
    .Data = TextResultImaginary.Text
End With
End Sub

Public Sub PutTextInChartWithoutComplex2()
TextResultReal = ResultReal
TextResultImaginary = ResultImaginary
With Chart
    .Row = 2
    .Column = 1
    .Data = TextComplex1a.Text
    .Column = 2
    .Data = TextComplex1b.Text
    .Column = 5
    .Data = TextResultReal.Text
    .Column = 6
    .Data = TextResultImaginary.Text
End With
End Sub

Private Sub MenuBar_Click()
Beep
End Sub

Private Sub MenuExit_Click()
End
End Sub

Public Sub ClearBlueLine()
    Line2.BorderColor = &H808080
    FrameComplex2.Enabled = False
    TextComplex2c.Enabled = False
    TextComplex2d.Enabled = False
    Label2.Enabled = False
    Label4.Enabled = False
    With Chart
        .Row = 2
        .Column = 3
        .Data = 0
        .Column = 4
        .Data = 0
    End With
End Sub

Public Sub ResetChartAndBlueLine()
    Line2.BorderColor = &HFF0000
    FrameComplex2.Enabled = True
    TextComplex2c.Enabled = True
    TextComplex2d.Enabled = True
    Label2.Enabled = True
    Label4.Enabled = True
    With Chart
        .Row = 2
        .Column = 3
        .Data = TextComplex2c.Text
        .Column = 4
        .Data = TextComplex2d.Text
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub MenuAbout1_Click()
FormAboutComplexMath.Show
End Sub

Private Sub MenuFileE_Click()
pi = 4 * Atn(1)
TextComplex1a = Exp(1)
TextComplex1b = 0
TextComplex2c = 0
TextComplex2d = pi
ComboFunction.ListIndex = 6
Call CommandCalculate_Click
End Sub

Private Sub MenuFileExit_Click()
End
End Sub

Private Sub MenuFileI_Click()
TextComplex1a = 0
TextComplex1b = 1
TextComplex2c = 0
TextComplex2d = 1
ComboFunction.ListIndex = 6
Call CommandCalculate_Click
End Sub

Private Sub MenuFileL1_Click()
TextComplex1a = -1
TextComplex1b = 0
TextComplex2c = 0
TextComplex2d = 0
ComboFunction.ListIndex = 7
Call CommandCalculate_Click
End Sub

Private Sub MenuFileL2_Click()
TextComplex1a = 0
TextComplex1b = 1
TextComplex2c = 0
TextComplex2d = 0
ComboFunction.ListIndex = 7
Call CommandCalculate_Click
End Sub

Private Sub MenuHelpHelp_Click()
FormHelp.Show
End Sub

Private Sub MenuHelpWhatsNew_Click()
FormWhatsNew.Show
End Sub

Private Sub TextComplex1a_Change()
CommandCalculate.Caption = "CALCULATE !"
'ResultReal = 0
'ResultImaginary = 0
PutTextInChart
End Sub

Private Sub TextComplex1b_Change()
CommandCalculate.Caption = "CALCULATE !"
'ResultReal = 0
'ResultImaginary = 0
PutTextInChart
End Sub

Private Sub TextComplex2c_Change()
CommandCalculate.Caption = "CALCULATE !"
'ResultReal = 0
'ResultImaginary = 0
PutTextInChart
End Sub

Private Sub TextComplex2d_Change()
CommandCalculate.Caption = "CALCULATE !"
'ResultReal = 0
'ResultImaginary = 0
PutTextInChart
End Sub
