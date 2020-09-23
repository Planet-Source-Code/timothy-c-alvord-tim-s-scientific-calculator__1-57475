VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Calculator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   ForeColor       =   &H00000000&
   Icon            =   "Calculator.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame DegRadGradFrame 
      Height          =   615
      Left            =   120
      TabIndex        =   36
      Top             =   1320
      Width           =   2415
      Begin VB.OptionButton Grad 
         Caption         =   "Grad"
         Height          =   255
         Left            =   1560
         TabIndex        =   39
         ToolTipText     =   "Gradients"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Rad 
         Caption         =   "Rad"
         Height          =   255
         Left            =   840
         TabIndex        =   38
         ToolTipText     =   "Radians"
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton Deg 
         Caption         =   "Deg"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         ToolTipText     =   "Degrees"
         Top             =   240
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame RadixFrame 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Radix"
      Top             =   600
      Width           =   4215
      Begin VB.OptionButton Roman 
         Caption         =   "Rom"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         ToolTipText     =   "Roman Numeral Input"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton Hex 
         Caption         =   "Hex"
         Height          =   375
         Left            =   960
         TabIndex        =   54
         ToolTipText     =   "Hex Input"
         Top             =   180
         Width           =   615
      End
      Begin VB.OptionButton Dec 
         Caption         =   "Dec"
         Height          =   375
         Left            =   1800
         TabIndex        =   53
         ToolTipText     =   "Decimal Input"
         Top             =   180
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Oct 
         Caption         =   "Oct"
         Height          =   375
         Left            =   2640
         TabIndex        =   52
         ToolTipText     =   "Octal Input"
         Top             =   180
         Width           =   615
      End
      Begin VB.OptionButton Bin 
         Caption         =   "Bin"
         Height          =   375
         Left            =   3480
         TabIndex        =   51
         ToolTipText     =   "Binary Input"
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.TextBox OutputWindow 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "0."
      ToolTipText     =   "Calculator Output Window"
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      Caption         =   "c Tim Alvord"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   360
      TabIndex        =   64
      Top             =   3120
      Width           =   990
   End
   Begin MSForms.CommandButton ButtonOzToGram 
      Height          =   450
      Left            =   840
      TabIndex        =   63
      ToolTipText     =   "Ounces <-> Grams"
      Top             =   3480
      Width           =   645
      ForeColor       =   12582912
      Caption         =   "Oz-Gm"
      Size            =   "1138;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonBS 
      Height          =   450
      Left            =   3720
      TabIndex        =   62
      ToolTipText     =   "Backspace Key"
      Top             =   3480
      Width           =   645
      ForeColor       =   255
      Caption         =   "BS"
      Size            =   "1138;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Image Yankees 
      BorderStyle     =   1  'Fixed Single
      Height          =   705
      Left            =   3480
      Picture         =   "Calculator.frx":0442
      Stretch         =   -1  'True
      ToolTipText     =   "Yankees Icon"
      Top             =   1295
      Width           =   855
   End
   Begin MSForms.CommandButton ButtonMileToKilo 
      Height          =   450
      Left            =   840
      TabIndex        =   61
      ToolTipText     =   "Mile <-> Kilometer"
      Top             =   4920
      Width           =   645
      ForeColor       =   12582912
      Caption         =   "Mi-Km"
      Size            =   "1138;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonV 
      Height          =   450
      Left            =   2280
      TabIndex        =   60
      ToolTipText     =   "Roam Numeral V"
      Top             =   2520
      Width           =   645
      Caption         =   "V"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonI 
      Height          =   450
      Left            =   3000
      TabIndex        =   59
      ToolTipText     =   "Roam Numeral I"
      Top             =   2520
      Width           =   645
      Caption         =   "I"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonL 
      Height          =   450
      Left            =   840
      TabIndex        =   58
      ToolTipText     =   "Roman Numeral L"
      Top             =   2520
      Width           =   645
      Caption         =   "L"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonX 
      Height          =   450
      Left            =   1560
      TabIndex        =   57
      ToolTipText     =   "Roman Numeral X"
      Top             =   2520
      Width           =   645
      Caption         =   "X"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonM 
      Height          =   450
      Left            =   120
      TabIndex        =   56
      ToolTipText     =   "Roman Numeral M"
      Top             =   2520
      Width           =   650
      Caption         =   "M"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonGallonToLitre 
      Height          =   450
      Left            =   840
      TabIndex        =   50
      ToolTipText     =   "Gallon <-> Litre"
      Top             =   4440
      Width           =   645
      ForeColor       =   12582912
      Caption         =   "Gal-Ltr"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonInchToCent 
      Height          =   450
      Left            =   840
      TabIndex        =   49
      ToolTipText     =   "Inch <-> Centimeter"
      Top             =   5400
      Width           =   645
      ForeColor       =   12582912
      Caption         =   "In-Cm"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Button1OverX 
      Height          =   450
      Left            =   3000
      TabIndex        =   48
      ToolTipText     =   "Invert Number"
      Top             =   3480
      Width           =   645
      ForeColor       =   12582912
      Caption         =   "1/X"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonSqrRoot 
      Height          =   450
      Left            =   2280
      TabIndex        =   47
      ToolTipText     =   "Square Root"
      Top             =   3480
      Width           =   645
      ForeColor       =   12582912
      Caption         =   "Sqr"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonFactorial 
      Height          =   450
      Left            =   1560
      TabIndex        =   46
      ToolTipText     =   "Factorial"
      Top             =   3480
      Width           =   645
      ForeColor       =   12582912
      Caption         =   "X!"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonPower 
      Height          =   450
      Left            =   3000
      TabIndex        =   45
      ToolTipText     =   "Power"
      Top             =   3960
      Width           =   645
      ForeColor       =   12582912
      Caption         =   "X^Y"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonCube 
      Height          =   450
      Left            =   2280
      TabIndex        =   44
      ToolTipText     =   "Cube"
      Top             =   3960
      Width           =   645
      ForeColor       =   12582912
      Caption         =   "X^3"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonSquare 
      Height          =   450
      Left            =   1560
      TabIndex        =   43
      ToolTipText     =   "Square"
      Top             =   3960
      Width           =   645
      ForeColor       =   12582912
      Caption         =   "X^2"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonSine 
      Height          =   450
      Left            =   1560
      TabIndex        =   42
      ToolTipText     =   "Sine"
      Top             =   3000
      Width           =   645
      ForeColor       =   12582912
      Caption         =   "Sin"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonCosine 
      Height          =   450
      Left            =   2280
      TabIndex        =   41
      ToolTipText     =   "Cosine"
      Top             =   3000
      Width           =   645
      ForeColor       =   12582912
      Caption         =   "Cos"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonTangent 
      Height          =   450
      Left            =   3000
      TabIndex        =   40
      ToolTipText     =   "Tangent"
      Top             =   3000
      Width           =   645
      ForeColor       =   12582912
      Caption         =   "Tan"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonLbToKilo 
      Height          =   450
      Left            =   840
      TabIndex        =   35
      ToolTipText     =   "Pound <-> Kilogram"
      Top             =   3960
      Width           =   645
      ForeColor       =   12582912
      Caption         =   "Lb-Kg"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonFToC 
      Height          =   450
      Left            =   840
      TabIndex        =   34
      ToolTipText     =   "Fahrenheit <-> Celsius"
      Top             =   5880
      Width           =   645
      ForeColor       =   12582912
      Caption         =   "F - C"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CheckBox InvCheckBox 
      Height          =   255
      Left            =   2640
      TabIndex        =   33
      ToolTipText     =   "Inverse - Causes some keys to perform the inverse of what's on the key"
      Top             =   1560
      Width           =   615
      VariousPropertyBits=   746588179
      BackColor       =   -2147483633
      ForeColor       =   0
      DisplayStyle    =   4
      Size            =   "1085;450"
      Value           =   "0"
      Caption         =   "Inv"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton Button7 
      Height          =   450
      Left            =   1560
      TabIndex        =   32
      ToolTipText     =   "7"
      Top             =   4440
      Width           =   645
      Caption         =   "7"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Button4 
      Height          =   450
      Left            =   1560
      TabIndex        =   31
      ToolTipText     =   "4"
      Top             =   4920
      Width           =   645
      Caption         =   "4"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Button1 
      Height          =   450
      Left            =   1560
      TabIndex        =   30
      ToolTipText     =   "1"
      Top             =   5400
      Width           =   645
      Caption         =   "1"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonPlusMinus 
      Height          =   450
      Left            =   1560
      TabIndex        =   29
      ToolTipText     =   "Plus/Minus"
      Top             =   5880
      Width           =   645
      ForeColor       =   12582912
      Caption         =   "+/-"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Button8 
      Height          =   450
      Left            =   2280
      TabIndex        =   28
      ToolTipText     =   "8"
      Top             =   4440
      Width           =   645
      Caption         =   "8"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Button5 
      Height          =   450
      Left            =   2280
      TabIndex        =   27
      ToolTipText     =   "5"
      Top             =   4920
      Width           =   645
      Caption         =   "5"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Button2 
      Height          =   450
      Left            =   2280
      TabIndex        =   26
      ToolTipText     =   "2"
      Top             =   5400
      Width           =   645
      Caption         =   "2"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Button0 
      Height          =   450
      Left            =   2280
      TabIndex        =   25
      ToolTipText     =   "0"
      Top             =   5880
      Width           =   645
      Caption         =   "0"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Button9 
      Height          =   450
      Left            =   3000
      TabIndex        =   24
      ToolTipText     =   "9"
      Top             =   4440
      Width           =   645
      Caption         =   "9"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Button6 
      Height          =   450
      Left            =   3000
      TabIndex        =   23
      ToolTipText     =   "6"
      Top             =   4920
      Width           =   645
      Caption         =   "6"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Button3 
      Height          =   450
      Left            =   3000
      TabIndex        =   22
      ToolTipText     =   "3"
      Top             =   5400
      Width           =   645
      Caption         =   "3"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonDecPoint 
      Height          =   450
      Left            =   3000
      TabIndex        =   21
      Top             =   5880
      Width           =   645
      Caption         =   "."
      Size            =   "1147;794"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton ButtonDivide 
      Height          =   450
      Left            =   3720
      TabIndex        =   20
      ToolTipText     =   "Divide"
      Top             =   3960
      Width           =   645
      Caption         =   "/"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonMultiply 
      Height          =   450
      Left            =   3720
      TabIndex        =   19
      ToolTipText     =   "Multiply"
      Top             =   4440
      Width           =   645
      Caption         =   "X"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonMinus 
      Height          =   450
      Left            =   3720
      TabIndex        =   18
      ToolTipText     =   "Minus"
      Top             =   4920
      Width           =   645
      Caption         =   "-"
      Size            =   "1147;794"
      FontEffects     =   1073741825
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton ButtonPlus 
      Height          =   450
      Left            =   3720
      TabIndex        =   17
      ToolTipText     =   "Plus"
      Top             =   5400
      Width           =   645
      Caption         =   "+"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonEquals 
      Height          =   450
      Left            =   3720
      TabIndex        =   16
      ToolTipText     =   "Equals"
      Top             =   5880
      Width           =   645
      Caption         =   "="
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonPI 
      Height          =   450
      Left            =   120
      TabIndex        =   15
      ToolTipText     =   "PI"
      Top             =   3480
      Width           =   645
      Caption         =   "PI"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonMMinus 
      Height          =   450
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Memory Minus"
      Top             =   5880
      Width           =   645
      Caption         =   "M-"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonMPlus 
      Height          =   450
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Memory Plus"
      Top             =   5400
      Width           =   645
      Caption         =   "M+"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonMS 
      Height          =   450
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Memory Save"
      Top             =   4920
      Width           =   645
      Caption         =   "MS"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonMR 
      Height          =   450
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Memory Recall"
      Top             =   4440
      Width           =   645
      Caption         =   "MR"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonMC 
      Height          =   450
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Memory Clear"
      Top             =   3960
      Width           =   645
      Caption         =   "MC"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonCE 
      Height          =   450
      Left            =   3720
      TabIndex        =   9
      ToolTipText     =   "Clear Entry"
      Top             =   3000
      Width           =   645
      ForeColor       =   255
      Caption         =   "CE"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonF 
      Height          =   450
      Left            =   3720
      TabIndex        =   8
      ToolTipText     =   "Hex F"
      Top             =   2040
      Width           =   645
      Caption         =   "F"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonE 
      Height          =   450
      Left            =   3000
      TabIndex        =   7
      ToolTipText     =   "Hex E"
      Top             =   2040
      Width           =   645
      Caption         =   "E"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonD 
      Height          =   450
      Left            =   2280
      TabIndex        =   6
      ToolTipText     =   "Hex D"
      Top             =   2040
      Width           =   645
      Caption         =   "D"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonC 
      Height          =   450
      Left            =   1560
      TabIndex        =   5
      ToolTipText     =   "Hex C"
      Top             =   2040
      Width           =   645
      Caption         =   "C"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonB 
      Height          =   450
      Left            =   840
      TabIndex        =   4
      ToolTipText     =   "Hex B"
      Top             =   2040
      Width           =   645
      Caption         =   "B"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonA 
      Height          =   450
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Hex A"
      Top             =   2040
      Width           =   650
      Caption         =   "A"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ButtonClear 
      Height          =   450
      Left            =   3720
      TabIndex        =   2
      ToolTipText     =   "Clear"
      Top             =   2520
      Width           =   645
      ForeColor       =   255
      Caption         =   "C"
      Size            =   "1147;794"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************
'   Scientific Calculator Program
'   Author: Timothy C. Alvord
'   E-Mail: tim8w@yahoo.com
'
'   Purpose:
'       This program is a Simple Scientific Calculator. It handles
'       Roman Numeral, Hex, Decimal, Octal and Binary numbers.
'       Degrees, Radians and Gradians. All the standard Trig functions.
'       Factorial, Sqr, Inverse, Square, Cube, X^Y.
'       Fun conversions like:
'           Ounce <-> Grams
'           Pounds <-> Kilograms
'           Gallon <-> Litre
'           Mile <-> Kilometer
'           Inch <-> Centimeter
'           Fahrenheight <-> Celsius
'***********************************************************************
Const PI = 3.14159265358979
' Function Mode
Const NONE = 0
Const MULTIPLY = 1
Const DIVIDE = 2
Const PLUS = 3
Const MINUS = 4
Const POWER = 5
' Base
Const ROMANNUM = 1
Const HEXNUM = 2
Const DECNUM = 3
Const OCTNUM = 4
Const BINNUM = 5
' Deg/Rad/Grad
Const DEGREES = 1
Const RADIANS = 2
Const GRADIANS = 3

Public xFirstNum As Double
Public xSecondNum As Double
Public xMemory As Double
Public bError As Boolean
Public bEntered As Boolean
Public bFirstNum As Boolean
Public iMathFunction As Integer
Public iCurrentRadix As Integer
Public iDegRadGrad As Integer

Private Sub Form_Load()
    iMathFunction = NONE
    iCurrentRadix = DECNUM
    iDegRadGrad = DEGREES
    xFirstNum = 0
    xSecondNum = 0
    bError = False
    bEntered = False
    
    Call Dec_Click
End Sub

Private Sub Form_KeyPress(KeyCode As Integer)
    
    MyChar = Chr(KeyCode)
    Select Case (MyChar)
    Case "0"
        Call Button0_Click
    Case "1"
        Call Button1_Click
    Case "2"
        If Dec.Value = True Or Oct.Value = True Or Hex.Value = True Then
            Call Button2_Click
        End If
    Case "3"
        If Dec.Value = True Or Oct.Value = True Or Hex.Value = True Then
            Call Button3_Click
        End If
    Case "4"
        If Dec.Value = True Or Oct.Value = True Or Hex.Value = True Then
            Call Button4_Click
        End If
    Case "5"
        If Dec.Value = True Or Oct.Value = True Or Hex.Value = True Then
            Call Button5_Click
        End If
    Case "6"
        If Dec.Value = True Or Oct.Value = True Or Hex.Value = True Then
            Call Button6_Click
        End If
    Case "7"
        If Dec.Value = True Or Oct.Value = True Or Hex.Value = True Then
            Call Button7_Click
        End If
    Case "8"
        If Dec.Value = True Or Oct.Value = True Or Hex.Value = True Then
            Call Button8_Click
        End If
    Case "9"
        If Dec.Value = True Or Hex.Value = True Then
            Call Button9_Click
        End If
    Case "A", "a"
        If Hex.Value = True Then
            Call ButtonA_Click
        End If
    Case "B", "b"
        If Hex.Value = True Then
            Call ButtonB_Click
        End If
    Case "C", "c"
        If iCurrentRadix = ROMANNUM Or iCurrentRadix = HEXNUM Then
            Call ButtonC_Click
        End If
    Case "D", "d"
        If iCurrentRadix = ROMANNUM Or iCurrentRadix = HEXNUM Then
            Call ButtonD_Click
        End If
    Case "E", "e"
        If Hex.Value = True Then
            Call ButtonE_Click
        End If
    Case "F", "f"
        If Hex.Value = True Then
            Call ButtonF_Click
        End If
    Case "M", "m"
        If iCurrentRadix = ROMANNUM Then
            Call ButtonM_Click
        End If
    Case "L", "l"
        If iCurrentRadix = ROMANNUM Then
            Call ButtonL_Click
        End If
    Case "X", "x"
        If iCurrentRadix = ROMANNUM Then
            Call ButtonX_Click
        End If
    Case "V", "v"
        If iCurrentRadix = ROMANNUM Then
            Call ButtonV_Click
        End If
    Case "I", "i"
        If iCurrentRadix = ROMANNUM Then
            Call ButtonI_Click
        End If
    Case "."
        If Dec.Value = True Then
            Call ButtonDecPoint_Click
        End If
    Case "/"
        Call ButtonDivide_Click
    Case "*"
        Call ButtonMultiply_Click
    Case "+"
        Call ButtonPlus_Click
    Case "-"
        Call ButtonMinus_Click
    Case "="
        Call ButtonEquals_Click
    Case Else
        Select Case (KeyCode)
        Case 8      '   Backspace Key
            Call ButtonBS_Click
        Case 13     '   Enter Key - Treat as Equal Key
            Call ButtonEquals_Click
        Case 27     '   Esc Key
            Call ButtonClear_Click
        End Select

    End Select
End Sub

Private Sub Roman_Click()
    If bError = False Then
        Call ConvertOutputWindowText(iCurrentRadix, ROMANNUM)
        iCurrentRadix = ROMANNUM
        Roman.Value = True
        Hex.Value = False
        Dec.Value = False
        Oct.Value = False
        Bin.Value = False
        
        ButtonM.Enabled = True
        ButtonC.Enabled = True
        ButtonD.Enabled = True
        ButtonL.Enabled = True
        ButtonX.Enabled = True
        ButtonV.Enabled = True
        ButtonI.Enabled = True
        
        ButtonSqrRoot.Enabled = True
        ButtonFactorial.Enabled = True
        ButtonSquare.Enabled = True
        ButtonCube.Enabled = True
        ButtonPower.Enabled = True

        ButtonA.Enabled = False
        ButtonB.Enabled = False
        ButtonE.Enabled = False
        ButtonF.Enabled = False
        Button9.Enabled = False
        Button8.Enabled = False
        Button7.Enabled = False
        Button6.Enabled = False
        Button5.Enabled = False
        Button4.Enabled = False
        Button3.Enabled = False
        Button2.Enabled = False
        Button1.Enabled = False
        Button0.Enabled = False
        ButtonTangent.Enabled = False
        ButtonCosine.Enabled = False
        ButtonSine.Enabled = False
        ButtonOzToGram.Enabled = False
        ButtonMileToKilo.Enabled = False
        ButtonGallonToLitre.Enabled = False
        ButtonInchToCent.Enabled = False
        ButtonLbToKilo.Enabled = False
        ButtonFToC.Enabled = False
        Button1OverX.Enabled = False
        ButtonPlusMinus.Enabled = False
    End If
End Sub

Private Sub Hex_Click()
    If bError = False Then
        Call ConvertOutputWindowText(iCurrentRadix, HEXNUM)
        iCurrentRadix = HEXNUM
        Hex.Value = True
        Roman.Value = False
        Dec.Value = False
        Oct.Value = False
        Bin.Value = False
        
        ButtonA.Enabled = True
        ButtonB.Enabled = True
        ButtonC.Enabled = True
        ButtonD.Enabled = True
        ButtonE.Enabled = True
        ButtonF.Enabled = True
        Button9.Enabled = True
        Button8.Enabled = True
        Button7.Enabled = True
        Button6.Enabled = True
        Button5.Enabled = True
        Button4.Enabled = True
        Button3.Enabled = True
        Button2.Enabled = True
        Button1.Enabled = True
        Button0.Enabled = True
        ButtonSqrRoot.Enabled = True
        ButtonFactorial.Enabled = True
        ButtonSquare.Enabled = True
        ButtonCube.Enabled = True
        ButtonPower.Enabled = True
        ButtonPlusMinus.Enabled = True

        ButtonM.Enabled = False
        ButtonL.Enabled = False
        ButtonX.Enabled = False
        ButtonV.Enabled = False
        ButtonI.Enabled = False
        ButtonTangent.Enabled = False
        ButtonCosine.Enabled = False
        ButtonSine.Enabled = False
        ButtonOzToGram.Enabled = False
        ButtonMileToKilo.Enabled = False
        ButtonGallonToLitre.Enabled = False
        ButtonInchToCent.Enabled = False
        ButtonLbToKilo.Enabled = False
        ButtonFToC.Enabled = False
        Button1OverX.Enabled = False
    End If
End Sub

Private Sub Dec_Click()
    If bError = False Then
        Call ConvertOutputWindowText(iCurrentRadix, DECNUM)
        iCurrentRadix = DECNUM
        
        Dec.Value = True
        Roman.Value = False
        Hex.Value = False
        Oct.Value = False
        Bin.Value = False
    
        ButtonA.Enabled = False
        ButtonB.Enabled = False
        ButtonC.Enabled = False
        ButtonD.Enabled = False
        ButtonE.Enabled = False
        ButtonF.Enabled = False
        ButtonM.Enabled = False
        ButtonL.Enabled = False
        ButtonX.Enabled = False
        ButtonV.Enabled = False
        ButtonI.Enabled = False
        
        Button9.Enabled = True
        Button8.Enabled = True
        Button7.Enabled = True
        Button6.Enabled = True
        Button5.Enabled = True
        Button4.Enabled = True
        Button3.Enabled = True
        Button2.Enabled = True
        Button1.Enabled = True
        Button0.Enabled = True
        ButtonSqrRoot.Enabled = True
        Button1OverX.Enabled = True
        ButtonSquare.Enabled = True
        ButtonCube.Enabled = True
        ButtonPower.Enabled = True
        ButtonTangent.Enabled = True
        ButtonCosine.Enabled = True
        ButtonSine.Enabled = True
        ButtonOzToGram.Enabled = True
        ButtonMileToKilo.Enabled = True
        ButtonGallonToLitre.Enabled = True
        ButtonInchToCent.Enabled = True
        ButtonLbToKilo.Enabled = True
        ButtonFToC.Enabled = True
        ButtonPlusMinus.Enabled = True
    End If
End Sub

Private Sub Oct_Click()
    If bError = False Then
        Call ConvertOutputWindowText(iCurrentRadix, OCTNUM)
        iCurrentRadix = OCTNUM
        
        Oct.Value = True
        Roman.Value = False
        Hex.Value = False
        Dec.Value = False
        Bin.Value = False
    
        ButtonTangent.Enabled = False
        ButtonCosine.Enabled = False
        ButtonSine.Enabled = False
        ButtonOzToGram.Enabled = False
        ButtonMileToKilo.Enabled = False
        ButtonGallonToLitre.Enabled = False
        ButtonInchToCent.Enabled = False
        ButtonLbToKilo.Enabled = False
        ButtonFToC.Enabled = False
        Button1OverX.Enabled = False
        ButtonA.Enabled = False
        ButtonB.Enabled = False
        ButtonC.Enabled = False
        ButtonD.Enabled = False
        ButtonE.Enabled = False
        ButtonF.Enabled = False
        ButtonM.Enabled = False
        ButtonL.Enabled = False
        ButtonX.Enabled = False
        ButtonV.Enabled = False
        ButtonI.Enabled = False
        Button9.Enabled = False
        
        ButtonSqrRoot.Enabled = True
        ButtonSquare.Enabled = True
        ButtonCube.Enabled = True
        ButtonPower.Enabled = True
        ButtonPlusMinus.Enabled = True
        Button8.Enabled = True
        Button7.Enabled = True
        Button6.Enabled = True
        Button5.Enabled = True
        Button4.Enabled = True
        Button3.Enabled = True
        Button2.Enabled = True
        Button1.Enabled = True
        Button0.Enabled = True
    End If
End Sub

Private Sub Bin_Click()
    If bError = False Then
        Call ConvertOutputWindowText(iCurrentRadix, BINNUM)
        iCurrentRadix = BINNUM
        
        Bin.Value = True
        Roman.Value = False
        Hex.Value = False
        Dec.Value = False
        Oct.Value = False
    
        ButtonSqrRoot.Enabled = True
        ButtonSquare.Enabled = True
        ButtonCube.Enabled = True
        ButtonPower.Enabled = True
        ButtonPlusMinus.Enabled = True
        Button1.Enabled = True
        Button0.Enabled = True
        
        ButtonA.Enabled = False
        ButtonB.Enabled = False
        ButtonC.Enabled = False
        ButtonD.Enabled = False
        ButtonE.Enabled = False
        ButtonF.Enabled = False
        ButtonM.Enabled = False
        ButtonL.Enabled = False
        ButtonX.Enabled = False
        ButtonV.Enabled = False
        ButtonI.Enabled = False
        Button9.Enabled = False
        Button8.Enabled = False
        Button7.Enabled = False
        Button6.Enabled = False
        Button5.Enabled = False
        Button4.Enabled = False
        Button3.Enabled = False
        Button2.Enabled = False
        ButtonTangent.Enabled = False
        ButtonCosine.Enabled = False
        ButtonSine.Enabled = False
        ButtonOzToGram.Enabled = False
        ButtonMileToKilo.Enabled = False
        ButtonGallonToLitre.Enabled = False
        ButtonInchToCent.Enabled = False
        ButtonLbToKilo.Enabled = False
        ButtonFToC.Enabled = False
        Button1OverX.Enabled = False
    End If
End Sub
Private Sub Deg_Click()
        Deg.Value = True
        Rad.Value = False
        Grad.Value = False
        iDegRadGrad = DEGREES
End Sub
Private Sub Rad_Click()
        Deg.Value = False
        Rad.Value = True
        Grad.Value = False
        iDegRadGrad = RADIANS
End Sub
Private Sub Grad_Click()
        Deg.Value = False
        Rad.Value = False
        Grad.Value = True
        iDegRadGrad = GRADIANS
End Sub
Private Sub ButtonPI_Click()
    If bError = False Then
        Select Case iCurrentRadix
            Case ROMANNUM
                OutputWindow.Text = GetDecRomanStr(PI)
            Case HEXNUM
                OutputWindow.Text = GetDecHexStr(PI)
            Case DECNUM
                OutputWindow.Text = PI
            Case OCTNUM
                OutputWindow.Text = GetDecOctStr(PI)
            Case BINNUM
                OutputWindow.Text = GetDecBinStr(PI)
        End Select
        bEntered = True
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonClear_Click()
    bEntered = False
    bFirstNum = False
    OutputWindow.Text = "0."
    OutputWindow.ForeColor = &H0&
    iMathFunction = 0
    bError = False
    Call ButtonMC_Click
    OutputWindow.SetFocus
End Sub

Private Sub ButtonCE_Click()
    If bError = False Then
        If bEntered = True Then
            If bFirstNum = False Then   '   Same as Clear
                bEntered = False
                bFirstNum = False
                OutputWindow.Text = "0."
                iMathFunction = 0
            Else    '   Allow User to Enter a new 2nd Number
                bEntered = False
                OutputWindow.Text = "0."
            End If
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonBS_Click()
    If bError = False Then
        If bEntered = True Then
            sStr = OutputWindow.Text
            If Len(sStr) > 1 Then
                OutputWindow.Text = Left(sStr, Len(sStr) - 1)
            Else
                OutputWindow.Text = "0"
            End If
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub Button0_Click()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "0"
        Else
            OutputWindow.Text = OutputWindow.Text + "0"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub Button1_Click()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "1"
        Else
            OutputWindow.Text = OutputWindow.Text + "1"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub Button2_Click()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "2"
        Else
            OutputWindow.Text = OutputWindow.Text + "2"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub Button3_Click()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "3"
        Else
            OutputWindow.Text = OutputWindow.Text + "3"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub Button4_Click()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "4"
        Else
            OutputWindow.Text = OutputWindow.Text + "4"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub Button5_Click()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "5"
        Else
            OutputWindow.Text = OutputWindow.Text + "5"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub Button6_Click()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "6"
        Else
            OutputWindow.Text = OutputWindow.Text + "6"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub Button7_Click()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "7"
        Else
            OutputWindow.Text = OutputWindow.Text + "7"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub Button8_Click()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "8"
        Else
            OutputWindow.Text = OutputWindow.Text + "8"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub Button9_Click()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "9"
        Else
            OutputWindow.Text = OutputWindow.Text + "9"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonA_Click()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "A"
        Else
            OutputWindow.Text = OutputWindow.Text + "A"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonB_Click()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "B"
        Else
            OutputWindow.Text = OutputWindow.Text + "B"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonC_Click()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "C"
        Else
            OutputWindow.Text = OutputWindow.Text + "C"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonD_Click()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "D"
        Else
            OutputWindow.Text = OutputWindow.Text + "D"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonE_Click()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "E"
        Else
            OutputWindow.Text = OutputWindow.Text + "E"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonF_Click()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "F"
        Else
            OutputWindow.Text = OutputWindow.Text + "F"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonM_Click()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "M"
        Else
            OutputWindow.Text = OutputWindow.Text + "M"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonL_Click()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "L"
        Else
            OutputWindow.Text = OutputWindow.Text + "L"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonX_Click()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "X"
        Else
            OutputWindow.Text = OutputWindow.Text + "X"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonV_Click()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "V"
        Else
            OutputWindow.Text = OutputWindow.Text + "V"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonI_Click()
    If bError = False Then
        If bEntered = False Then
            bEntered = True
            OutputWindow.Text = "I"
        Else
            OutputWindow.Text = OutputWindow.Text + "I"
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonDecPoint_Click()
    If bError = False Then
        If iCurrentRadix = DECNUM Then
            If bEntered = False Then    '   Start of New Number
                bEntered = True
                OutputWindow.Text = "."
            Else                        '   Append
                If InStr(OutputWindow.Text, ".") = 0 Then   '   Make Sure Only 1 Decimal Point
                    OutputWindow.Text = OutputWindow.Text + "."
                End If
            End If
        End If
    End If
    OutputWindow.SetFocus
End Sub
Private Sub ButtonPlusMinus_Click()
    If bError = False Then
        If iCurrentRadix = DECNUM Then
            If bEntered = True Then
                iValue = OutputWindow.Text * -1
                OutputWindow.Text = iValue
            End If
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonDivide_Click()
    If bError = False Then
        If bEntered = True Then
            If bFirstNum = True Then
                Select Case iCurrentRadix
                    Case ROMANNUM
                        xSecondNum = GetRomanDecNum(OutputWindow.Text)
                    Case HEXNUM
                        xSecondNum = Val("&H" + OutputWindow.Text)
                    Case DECNUM
                        xSecondNum = OutputWindow.Text
                    Case OCTNUM
                        xSecondNum = Val("&O" + OutputWindow.Text)
                    Case BINNUM
                        xSecondNum = GetBinDecNum(OutputWindow.Text)
                End Select
                If xSecondNum = 0 Then
                    OutputWindow.Text = "ERROR - Divide by Zero"
                    OutputWindow.ForeColor = &HFF&
                    bError = True
                Else
                    xTotal = xFirstNum / xSecondNum
                    Select Case iCurrentRadix
                        Case ROMANNUM
                            OutputWindow.Text = GetDecRomanStr(xTotal)
                        Case HEXNUM
                            OutputWindow.Text = GetDecHexStr(xTotal)
                        Case DECNUM
                            OutputWindow.Text = xTotal
                        Case OCTNUM
                            OutputWindow.Text = GetDecOctStr(xTotal)
                        Case BINNUM
                            OutputWindow.Text = GetDecBinStr(xTotal)
                    End Select
                    
                    xFirstNum = xTotal
                    bFirstNum = True
                    bEntered = False
                End If
                iMathFunction = DIVIDE
            Else
                Select Case iCurrentRadix
                    Case ROMANNUM
                        xFirstNum = GetRomanDecNum(OutputWindow.Text)
                    Case HEXNUM
                        xFirstNum = Val("&H" + OutputWindow.Text)
                    Case DECNUM
                        xFirstNum = OutputWindow.Text
                    Case OCTNUM
                        xFirstNum = Val("&O" + OutputWindow.Text)
                    Case BINNUM
                        xFirstNum = GetBinDecNum(OutputWindow.Text)
                End Select
                bFirstNum = True
                bEntered = False
                iMathFunction = DIVIDE
            End If
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonMultiply_Click()
    If bError = False Then
        If bEntered = True Then
            If bFirstNum = True Then
                Select Case iCurrentRadix
                    Case ROMANNUM
                        xSecondNum = GetRomanDecNum(OutputWindow.Text)
                    Case HEXNUM
                        xSecondNum = Val("&H" + OutputWindow.Text)
                    Case DECNUM
                        xSecondNum = OutputWindow.Text
                    Case OCTNUM
                        xSecondNum = Val("&O" + OutputWindow.Text)
                    Case BINNUM
                        xSecondNum = GetBinDecNum(OutputWindow.Text)
                End Select

                xTotal = xFirstNum * xSecondNum
                Select Case iCurrentRadix
                   Case ROMANNUM
                       OutputWindow.Text = GetDecRomanStr(xTotal)
                   Case HEXNUM
                       OutputWindow.Text = GetDecHexStr(xTotal)
                   Case DECNUM
                       OutputWindow.Text = xTotal
                   Case OCTNUM
                       OutputWindow.Text = GetDecOctStr(xTotal)
                   Case BINNUM
                       OutputWindow.Text = GetDecBinStr(xTotal)
                End Select
                
                xFirstNum = xTotal
                bFirstNum = True
                bEntered = False
                iMathFunction = MULTIPLY
           Else
                Select Case iCurrentRadix
                    Case ROMANNUM
                        xFirstNum = GetRomanDecNum(OutputWindow.Text)
                    Case HEXNUM
                        xFirstNum = Val("&H" + OutputWindow.Text)
                    Case DECNUM
                        xFirstNum = OutputWindow.Text
                    Case OCTNUM
                        xFirstNum = Val("&O" + OutputWindow.Text)
                    Case BINNUM
                        xFirstNum = GetBinDecNum(OutputWindow.Text)
                End Select
                bFirstNum = True
                bEntered = False
                iMathFunction = MULTIPLY
            End If
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonPlus_Click()
    If bError = False Then
        If bEntered = True Then
            If bFirstNum = True Then
                Select Case iCurrentRadix
                    Case ROMANNUM
                        xSecondNum = GetRomanDecNum(OutputWindow.Text)
                    Case HEXNUM
                        xSecondNum = Val("&H" + OutputWindow.Text)
                    Case DECNUM
                        xSecondNum = OutputWindow.Text
                    Case OCTNUM
                        xSecondNum = Val("&O" + OutputWindow.Text)
                    Case BINNUM
                        xSecondNum = GetBinDecNum(OutputWindow.Text)
                End Select
                xTotal = xFirstNum + xSecondNum
                Select Case iCurrentRadix
                    Case ROMANNUM
                        OutputWindow.Text = GetDecRomanStr(xTotal)
                    Case HEXNUM
                        OutputWindow.Text = GetDecHexStr(xTotal)
                    Case DECNUM
                        OutputWindow.Text = xTotal
                    Case OCTNUM
                        OutputWindow.Text = GetDecOctStr(xTotal)
                    Case BINNUM
                        OutputWindow.Text = GetDecBinStr(xTotal)
                End Select
                
                xFirstNum = xTotal
                bFirstNum = True
                bEntered = False
                iMathFunction = PLUS
            Else
                Select Case iCurrentRadix
                    Case ROMANNUM
                        xFirstNum = GetRomanDecNum(OutputWindow.Text)
                    Case HEXNUM
                        xFirstNum = Val("&H" + OutputWindow.Text)
                    Case DECNUM
                        xFirstNum = OutputWindow.Text
                    Case OCTNUM
                        xFirstNum = Val("&O" + OutputWindow.Text)
                    Case BINNUM
                        xFirstNum = GetBinDecNum(OutputWindow.Text)
                End Select
                bFirstNum = True
                bEntered = False
                iMathFunction = PLUS
            End If
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonMinus_Click()
    If bError = False Then
        If bEntered = True Then
            If bFirstNum = True Then
                Select Case iCurrentRadix
                    Case ROMANNUM
                        xSecondNum = GetRomanDecNum(OutputWindow.Text)
                    Case HEXNUM
                        xSecondNum = Val("&H" + OutputWindow.Text)
                    Case DECNUM
                        xSecondNum = OutputWindow.Text
                    Case OCTNUM
                        xSecondNum = Val("&O" + OutputWindow.Text)
                    Case BINNUM
                        xSecondNum = GetBinDecNum(OutputWindow.Text)
                End Select
                xTotal = xFirstNum - xSecondNum
                Select Case iCurrentRadix
                    Case ROMANNUM
                        OutputWindow.Text = GetDecRomanStr(xTotal)
                    Case HEXNUM
                        OutputWindow.Text = GetDecHexStr(xTotal)
                    Case DECNUM
                        OutputWindow.Text = xTotal
                    Case OCTNUM
                        OutputWindow.Text = GetDecOctStr(xTotal)
                    Case BINNUM
                        OutputWindow.Text = GetDecBinStr(xTotal)
                End Select
                xFirstNum = xTotal
                bFirstNum = True
                bEntered = False
                iMathFunction = MINUS
            Else
                Select Case iCurrentRadix
                    Case ROMANNUM
                        xFirstNum = GetRomanDecNum(OutputWindow.Text)
                    Case HEXNUM
                        xFirstNum = Val("&H" + OutputWindow.Text)
                    Case DECNUM
                        xFirstNum = OutputWindow.Text
                    Case OCTNUM
                        xFirstNum = Val("&O" + OutputWindow.Text)
                    Case BINNUM
                        xFirstNum = GetBinDecNum(OutputWindow.Text)
                End Select
                bFirstNum = True
                bEntered = False
                iMathFunction = MINUS
            End If
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonPower_Click()
    If bError = False Then
        If bEntered = True Then
            Select Case iCurrentRadix
                Case ROMANNUM
                    xFirstNum = GetRomanDecNum(OutputWindow.Text)
                Case HEXNUM
                    xFirstNum = Val("&H" + OutputWindow.Text)
                Case DECNUM
                    xFirstNum = OutputWindow.Text
                Case OCTNUM
                    xFirstNum = Val("&O" + OutputWindow.Text)
                Case BINNUM
                    xFirstNum = GetBinDecNum(OutputWindow.Text)
            End Select
            bFirstNum = True
            bEntered = False
            iMathFunction = POWER
        End If
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonEquals_Click()
    If bError = False Then
        If bFirstNum = True Then
            Select Case iCurrentRadix
                Case ROMANNUM
                    xSecondNum = GetRomanDecNum(OutputWindow.Text)
                Case HEXNUM
                    xSecondNum = Val("&H" + OutputWindow.Text)
                Case DECNUM
                    xSecondNum = OutputWindow.Text
                Case OCTNUM
                    xSecondNum = Val("&O" + OutputWindow.Text)
                Case BINNUM
                    xSecondNum = GetBinDecNum(OutputWindow.Text)
            End Select
            
            Select Case iMathFunction
            Case DIVIDE
                If xSecondNum = 0 Then
                    OutputWindow.Text = "ERROR - Divide by Zero"
                    OutputWindow.ForeColor = &HFF&
                    bError = True
                Else
                    xTotal = xFirstNum / xSecondNum
                    Select Case iCurrentRadix
                        Case ROMANNUM
                            OutputWindow.Text = GetDecRomanStr(xTotal)
                        Case HEXNUM
                            OutputWindow.Text = GetDecHexStr(xTotal)
                        Case DECNUM
                            OutputWindow.Text = xTotal
                        Case OCTNUM
                            OutputWindow.Text = GetDecOctStr(xTotal)
                        Case BINNUM
                            OutputWindow.Text = GetDecBinStr(xTotal)
                    End Select
                    
                    xFirstNum = xTotal
                    bFirstNum = True
                    bEntered = True
                End If
                iMathFunction = 0
            Case MULTIPLY
                xTotal = xFirstNum * xSecondNum
                Select Case iCurrentRadix
                    Case ROMANNUM
                        OutputWindow.Text = GetDecRomanStr(xTotal)
                    Case HEXNUM
                        OutputWindow.Text = GetDecHexStr(xTotal)
                    Case DECNUM
                        OutputWindow.Text = xTotal
                    Case OCTNUM
                        OutputWindow.Text = GetDecOctStr(xTotal)
                    Case BINNUM
                        OutputWindow.Text = GetDecBinStr(xTotal)
                End Select
                
                xFirstNum = xTotal
                bFirstNum = True
                bEntered = True
                iMathFunction = 0
            Case PLUS
                xTotal = xFirstNum + xSecondNum
                Select Case iCurrentRadix
                    Case ROMANNUM
                        OutputWindow.Text = GetDecRomanStr(xTotal)
                    Case HEXNUM
                        OutputWindow.Text = GetDecHexStr(xTotal)
                    Case DECNUM
                        OutputWindow.Text = xTotal
                    Case OCTNUM
                        OutputWindow.Text = GetDecOctStr(xTotal)
                    Case BINNUM
                        OutputWindow.Text = GetDecBinStr(xTotal)
                End Select
                
                xFirstNum = xTotal
                bFirstNum = True
                bEntered = True
                iMathFunction = 0
            Case MINUS
                xTotal = xFirstNum - xSecondNum
                Select Case iCurrentRadix
                    Case ROMANNUM
                        OutputWindow.Text = GetDecRomanStr(xTotal)
                    Case HEXNUM
                        OutputWindow.Text = GetDecHexStr(xTotal)
                    Case DECNUM
                        OutputWindow.Text = xTotal
                    Case OCTNUM
                        OutputWindow.Text = GetDecOctStr(xTotal)
                    Case BINNUM
                        OutputWindow.Text = GetDecBinStr(xTotal)
                End Select
                xFirstNum = xTotal
                bFirstNum = True
                bEntered = True
                iMathFunction = 0
            Case POWER
                xTotal = 1
                If xSecondNum <> 0 Then
                    For iCounter = 1 To xSecondNum
                        xTotal = xTotal * xFirstNum
                    Next
                End If
                Select Case iCurrentRadix
                    Case ROMANNUM
                        OutputWindow.Text = GetDecRomanStr(xTotal)
                    Case HEXNUM
                        OutputWindow.Text = GetDecHexStr(xTotal)
                    Case DECNUM
                        OutputWindow.Text = xTotal
                    Case OCTNUM
                        OutputWindow.Text = GetDecOctStr(xTotal)
                    Case BINNUM
                        OutputWindow.Text = GetDecBinStr(xTotal)
                End Select
                xFirstNum = xTotal
                bFirstNum = True
                bEntered = True
                iMathFunction = 0
            End Select
        End If
    End If
    OutputWindow.SetFocus
End Sub
Private Sub ButtonMC_Click()
    xMemory = 0
    OutputWindow.SetFocus
End Sub

Private Sub ButtonMR_Click()
    If xMemory <> 0 Then
        Select Case iCurrentRadix
            Case ROMANNUM
                OutputWindow.Text = GetDecRomanStr(xMemory)
            Case HEXNUM
                OutputWindow.Text = GetDecHexStr(xMemory)
            Case DECNUM
                OutputWindow.Text = xMemory
            Case OCTNUM
                OutputWindow.Text = GetDecOctStr(xMemory)
            Case BINNUM
                OutputWindow.Text = GetDecBinStr(xMemory)
        End Select
        bEntered = True
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonMS_Click()
    Select Case iCurrentRadix
        Case ROMANNUM
            xMemory = GetRomanDecNum(OutputWindow.Text)
        Case HEXNUM
            xMemory = Val("&H" + OutputWindow.Text)
        Case DECNUM
            xMemory = OutputWindow.Text
        Case OCTNUM
            xMemory = Val("&O" + OutputWindow.Text)
        Case BINNUM
            xMemory = GetBinDecNum(OutputWindow.Text)
    End Select
    OutputWindow.SetFocus
End Sub

Private Sub ButtonMPlus_Click()
    Select Case iCurrentRadix
        Case ROMANNUM
            xMemory = xMemory + GetRomanDecNum(OutputWindow.Text)
        Case HEXNUM
            xMemory = xMemory + Val("&H" + OutputWindow.Text)
        Case DECNUM
            xMemory = xMemory + OutputWindow.Text
        Case OCTNUM
            xMemory = xMemory + Val("&O" + OutputWindow.Text)
        Case BINNUM
            xMemory = xMemory + GetBinDecNum(OutputWindow.Text)
    End Select
    OutputWindow.SetFocus
End Sub

Private Sub ButtonMMinus_Click()
    Select Case iCurrentRadix
        Case ROMANNUM
            xMemory = xMemory - GetRomanDecNum(OutputWindow.Text)
        Case HEXNUM
            xMemory = xMemory - Val("&H" + OutputWindow.Text)
        Case DECNUM
            xMemory = xMemory - OutputWindow.Text
        Case OCTNUM
            xMemory = xMemory - Val("&O" + OutputWindow.Text)
        Case BINNUM
            xMemory = xMemory - GetBinDecNum(OutputWindow.Text)
    End Select
    OutputWindow.SetFocus
End Sub

Private Sub InvCheckBox_Click()
    If InvCheckBox.Value = True Then
        ButtonOzToGram.Caption = "Gm-Oz"
        ButtonMileToKilo.Caption = "Km-Mi"
        ButtonGallonToLitre.Caption = "Ltr-Gal"
        ButtonInchToCent.Caption = "Cm-In"
        ButtonFToC.Caption = "C - F"
        ButtonLbToKilo.Caption = "Kg-Lb"
        ButtonSine.Caption = "Asn"
        ButtonCosine.Caption = "Acs"
        ButtonTangent.Caption = "Atn"
    Else
        ButtonOzToGram.Caption = "Oz-Gm"
        ButtonMileToKilo.Caption = "Mi-Km"
        ButtonGallonToLitre.Caption = "Gal-Ltr"
        ButtonInchToCent.Caption = "In-Cm"
        ButtonFToC.Caption = "F - C"
        ButtonLbToKilo.Caption = "Lb-Kg"
        ButtonSine.Caption = "Sin"
        ButtonCosine.Caption = "Cos"
        ButtonTangent.Caption = "Tan"
    End If
    OutputWindow.SetFocus
End Sub
Private Sub ButtonOzToGram_Click()
    If InvCheckBox.Value = True Then
        xGram = OutputWindow.Text
        xOunce = xGram / 28.34952313
        OutputWindow.Text = xOunce
        InvCheckBox.Value = False
        Call InvCheckBox_Click
    Else
        xOunce = OutputWindow.Text
        xGram = xOunce * 28.34952313
        OutputWindow.Text = xGram
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonMileToKilo_Click()
    If InvCheckBox.Value = True Then
        xKilometer = OutputWindow.Text
        xMile = xKilometer / 1.609344
        OutputWindow.Text = xMile
        InvCheckBox.Value = False
        Call InvCheckBox_Click
    Else
        xMile = OutputWindow.Text
        xKilometer = xMile * 1.609344
        OutputWindow.Text = xKilometer
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonGallonToLitre_Click()
    If InvCheckBox.Value = True Then
        xLitre = OutputWindow.Text
        xGallon = xLitre / 3.785412
        OutputWindow.Text = xGallon
        InvCheckBox.Value = False
        Call InvCheckBox_Click
    Else
        xGallon = OutputWindow.Text
        xLitre = xGallon * 3.785412
        OutputWindow.Text = xLitre
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonInchToCent_Click()
    If InvCheckBox.Value = True Then
        xCent = OutputWindow.Text
        xInch = xCent / 2.54
        OutputWindow.Text = xInch
        InvCheckBox.Value = False
        Call InvCheckBox_Click
    Else
        xInch = OutputWindow.Text
        xCent = xInch * 2.54
        OutputWindow.Text = xCent
    End If
    OutputWindow.SetFocus
End Sub
Private Sub ButtonFToC_Click()
    If InvCheckBox.Value = True Then
        xCelsius = OutputWindow.Text
        xFahrenheit = 32 + (9 * xCelsius / 5)
        OutputWindow.Text = xFahrenheit
        InvCheckBox.Value = False
        Call InvCheckBox_Click
    Else
        xFahrenheit = OutputWindow.Text
        xCelsius = ((xFahrenheit - 32) * 5) / 9
        OutputWindow.Text = xCelsius
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonLbToKilo_Click()
    If InvCheckBox.Value = True Then
        xKilograms = OutputWindow.Text
        xPounds = xKilograms * 2.204623
        OutputWindow.Text = xPounds
        InvCheckBox.Value = False
        Call InvCheckBox_Click
    Else
        xPounds = OutputWindow.Text
        xKilograms = xPounds / 2.204623
        OutputWindow.Text = xKilograms
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonSine_Click()
    If InvCheckBox.Value = True Then
        xAsin = ArcSin(OutputWindow.Text)
        If bError Then Exit Sub
        Select Case iDegRadGrad
            Case DEGREES
                OutputWindow.Text = xAsin * 180 / PI
            Case RADIANS
                OutputWindow.Text = xAsin
            Case GRADIANS
                xDeg = xAsin * 180 / PI
                OutputWindow.Text = xDeg * 10 / 9
        End Select
        InvCheckBox.Value = False
        Call InvCheckBox_Click
    Else
        Select Case iDegRadGrad
            Case DEGREES
                xsin = Sin(OutputWindow.Text * PI / 180)
            Case RADIANS
                xsin = Sin(OutputWindow.Text)
            Case GRADIANS
                xDeg = (OutputWindow.Text * 9) / 10
                xsin = Sin(xDeg * PI / 180)
        End Select
        OutputWindow.Text = xsin
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonCosine_Click()
    If InvCheckBox.Value = True Then
        xAcos = ArcCos(OutputWindow.Text)
        If bError Then Exit Sub
        Select Case iDegRadGrad
            Case DEGREES
                OutputWindow.Text = xAcos * 180 / PI
            Case RADIANS
                OutputWindow.Text = xAcos
            Case GRADIANS
                xDeg = xAcos * 10 / 9
                OutputWindow.Text = xDeg * 180 / PI
        End Select
        InvCheckBox.Value = False
        Call InvCheckBox_Click
    Else
        Select Case iDegRadGrad
            Case DEGREES
                xcos = Cos(OutputWindow.Text * PI / 180)
            Case RADIANS
                xcos = Cos(OutputWindow.Text)
            Case GRADIANS
                xDeg = (OutputWindow.Text * 9) / 10
                xcos = Cos(xDeg * PI / 180)
        End Select
        OutputWindow.Text = xcos
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonTangent_Click()
    If InvCheckBox.Value = True Then
        xAtan = Atn(OutputWindow.Text)
        Select Case iDegRadGrad
            Case DEGREES
                OutputWindow.Text = xAtan * 180 / PI
            Case RADIANS
                OutputWindow.Text = xAtan
            Case GRADIANS
                xDeg = xAtan * 180 / PI
                OutputWindow.Text = xDeg * 10 / 9
        End Select
        InvCheckBox.Value = False
        Call InvCheckBox_Click
    Else
        Select Case iDegRadGrad
            Case DEGREES
                xtan = Tan(OutputWindow.Text * PI / 180)
            Case RADIANS
                xtan = Tan(OutputWindow.Text)
            Case GRADIANS
                xDeg = (OutputWindow.Text * 9) / 10
                xtan = Tan(xDeg * PI / 180)
        End Select
        OutputWindow.Text = xtan
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonSquare_Click()
    Select Case iCurrentRadix
        Case ROMANNUM
            xNum = GetRomanDecNum(OutputWindow.Text)
        Case HEXNUM
            xNum = Val("&H" + OutputWindow.Text)
        Case DECNUM
            xNum = OutputWindow.Text
        Case OCTNUM
            xNum = Val("&O" + OutputWindow.Text)
        Case BINNUM
            xNum = GetBinDecNum(OutputWindow.Text)
    End Select
    xNum = xNum * xNum
    Select Case iCurrentRadix
        Case ROMANNUM
            OutputWindow.Text = GetDecRomanStr(xNum)
        Case HEXNUM
            OutputWindow.Text = GetDecHexStr(xNum)
        Case DECNUM
            OutputWindow.Text = xNum
        Case OCTNUM
            OutputWindow.Text = GetDecOctStr(xNum)
        Case BINNUM
            OutputWindow.Text = GetDecBinStr(xNum)
    End Select
    OutputWindow.SetFocus
End Sub

Private Sub ButtonCube_Click()
    Select Case iCurrentRadix
        Case ROMANNUM
            xNum = GetRomanDecNum(OutputWindow.Text)
        Case HEXNUM
            xNum = Val("&H" + OutputWindow.Text)
        Case DECNUM
            xNum = OutputWindow.Text
        Case OCTNUM
            xNum = Val("&O" + OutputWindow.Text)
        Case BINNUM
            xNum = GetBinDecNum(OutputWindow.Text)
    End Select
    xNum = xNum * xNum * xNum
    Select Case iCurrentRadix
        Case ROMANNUM
            OutputWindow.Text = GetDecRomanStr(xNum)
        Case HEXNUM
            OutputWindow.Text = GetDecHexStr(xNum)
        Case DECNUM
            OutputWindow.Text = xNum
        Case OCTNUM
            OutputWindow.Text = GetDecOctStr(xNum)
        Case BINNUM
            OutputWindow.Text = GetDecBinStr(xNum)
    End Select
    OutputWindow.SetFocus
End Sub
Public Function ArcSin(x As Variant) As Variant
    Select Case x
        Case -1
            ArcSin = 6 * Atn(1)
        Case 0:
            ArcSin = 0
        Case 1:
            ArcSin = 2 * Atn(1)
        Case Else:
            ArcSin = Atn(x / Sqr(-x * x + 1))
    End Select
End Function
Public Function ArcCos(x As Variant) As Variant

    Select Case x
        Case -1
            ArcCos = 4 * Atn(1)
             
        Case 0:
            ArcCos = 2 * Atn(1)
             
        Case 1:
            ArcCos = 0
             
        Case Else:
            ArcCos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
    End Select
End Function

Private Sub Button1OverX_Click()
    xValue = OutputWindow.Text
    If xValue = 0 Then
        OutputWindow.Text = "ERROR - Divide by Zero"
        OutputWindow.ForeColor = &HFF&
        bError = True
    Else
        OutputWindow.Text = 1 / xValue
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonSqrRoot_Click()
    Select Case iCurrentRadix
        Case ROMANNUM
            xValue = GetRomanDecNum(OutputWindow.Text)
        Case HEXNUM
            xValue = Val("&H" + OutputWindow.Text)
        Case DECNUM
            xValue = OutputWindow.Text
        Case OCTNUM
            xValue = Val("&O" + OutputWindow.Text)
        Case BINNUM
            xValue = GetBinDecNum(OutputWindow.Text)
    End Select
    If xValue < 0 Then
        OutputWindow.Text = "ERROR - Square Root of Negative Number"
        OutputWindow.ForeColor = &HFF&
        bError = True
    Else
        Select Case iCurrentRadix
            Case ROMANNUM
                OutputWindow.Text = GetDecRomanStr(Sqr(xValue))
            Case HEXNUM
                OutputWindow.Text = GetDecHexStr(Sqr(xValue))
            Case DECNUM
                OutputWindow.Text = Sqr(xValue)
            Case OCTNUM
                OutputWindow.Text = GetDecOctStr(Sqr(xValue))
            Case BINNUM
                OutputWindow.Text = GetDecBinStr(Sqr(xValue))
        End Select
    End If
    OutputWindow.SetFocus
End Sub

Private Sub ButtonFactorial_Click()
    Dim xFactorial, xNum As Double
    
    Select Case iCurrentRadix
        Case ROMANNUM
            xNum = GetRomanDecNum(OutputWindow.Text)
        Case HEXNUM
            xNum = Val("&H" + OutputWindow.Text)
        Case DECNUM
            xNum = OutputWindow.Text
        Case OCTNUM
            xNum = Val("&O" + OutputWindow.Text)
        Case BINNUM
            xNum = GetBinDecNum(OutputWindow.Text)
    End Select
    xFactorial = xNum
    On Error GoTo Factorial_Error
    For iCounter = (xFactorial - 1) To 1 Step -1
        xFactorial = xFactorial * iCounter
    Next
    Select Case iCurrentRadix
        Case ROMANNUM
            OutputWindow.Text = GetDecRomanStr(xFactorial)
        Case HEXNUM
            OutputWindow.Text = GetDecHexStr(xFactorial)
        Case DECNUM
            OutputWindow.Text = xFactorial
        Case OCTNUM
            OutputWindow.Text = GetDecOctStr(xFactorial)
        Case BINNUM
            OutputWindow.Text = GetDecBinStr(xFactorial)
    End Select
    OutputWindow.SetFocus
    Exit Sub
    
Factorial_Error:
    OutputWindow.Text = "ERROR - " + Err.Description
    OutputWindow.ForeColor = &HFF&
    bError = True
    Err.Clear
    OutputWindow.SetFocus
End Sub

Public Function GetDecRomanStr(ByVal xDecimal As Double) As String
    Dim iThousands, iHundreds, iTens, iOnes As Integer
    Dim sReturnStr As String
    Dim sHunds(9) As String
    Dim sTens(9) As String
    Dim sOnes(9) As String
    
    sHunds(1) = "C"
    sHunds(2) = "CC"
    sHunds(3) = "CCC"
    sHunds(4) = "CD"
    sHunds(5) = "D"
    sHunds(6) = "DC"
    sHunds(7) = "DCC"
    sHunds(8) = "DCCC"
    sHunds(9) = "CM"
    sTens(1) = "X"
    sTens(2) = "XX"
    sTens(3) = "XXX"
    sTens(4) = "XL"
    sTens(5) = "L"
    sTens(6) = "LX"
    sTens(7) = "LXX"
    sTens(8) = "LXXX"
    sTens(9) = "XC"
    sOnes(1) = "I"
    sOnes(2) = "II"
    sOnes(3) = "III"
    sOnes(4) = "IV"
    sOnes(5) = "V"
    sOnes(6) = "VI"
    sOnes(7) = "VII"
    sOnes(8) = "VIII"
    sOnes(9) = "IX"
    
    iThousands = (xDecimal - (xDecimal Mod 1000)) / 1000
    xDecimal = xDecimal Mod 1000
    iHundreds = (xDecimal - (xDecimal Mod 100)) / 100
    xDecimal = xDecimal Mod 100
    iTens = (xDecimal - (xDecimal Mod 10)) / 10
    xDecimal = xDecimal Mod 10
    iOnes = xDecimal
    
    sReturnStr = ""
    For iCount = 1 To iThousands
        sReturnStr = sReturnStr + "M"
    Next
    If iHundreds > 0 Then
        sReturnStr = sReturnStr + sHunds(iHundreds)
    End If
    If iTens > 0 Then
        sReturnStr = sReturnStr + sTens(iTens)
    End If
    If iOnes > 0 Then
        sReturnStr = sReturnStr + sOnes(iOnes)
    End If
    GetDecRomanStr = sReturnStr
End Function

Public Function GetRomanDecNum(ByVal sRomanStr As String) As Double
    Dim xDecimal As Double
    Dim sStr As String
        
    sStr = Left(sRomanStr, 1)
    While sStr = "M"
        xDecimal = xDecimal + 1000
        sRomanStr = Right(sRomanStr, Len(sRomanStr) - 1)
        sStr = Left(sRomanStr, 1)
    Wend
    
    iHunds = 0
    If Left(sRomanStr, 2) = "CM" Then
        iHunds = 9
        sRomanStr = Right(sRomanStr, Len(sRomanStr) - 2)
    Else
        If Left(sRomanStr, 1) = "D" Then
            iHunds = 5
            sRomanStr = Right(sRomanStr, Len(sRomanStr) - 1)
        Else
            If Left(sRomanStr, 2) = "CD" Then
                iHunds = 4
                sRomanStr = Right(sRomanStr, Len(sRomanStr) - 2)
            End If
        End If
    End If
    If iHunds = 0 Or iHunds = 5 Then
        sStr = Left(sRomanStr, 1)
        While sStr = "C"
            iHunds = iHunds + 1
            sRomanStr = Right(sRomanStr, Len(sRomanStr) - 1)
            sStr = Left(sRomanStr, 1)
        Wend
    End If
    xDecimal = xDecimal + iHunds * 100
    
    iTens = 0
    If Left(sRomanStr, 2) = "XC" Then
        iTens = 9
        sRomanStr = Right(sRomanStr, Len(sRomanStr) - 2)
    Else
        If Left(sRomanStr, 1) = "L" Then
            iTens = 5
            sRomanStr = Right(sRomanStr, Len(sRomanStr) - 1)
        Else
            If Left(sRomanStr, 2) = "XL" Then
                iTens = 4
                sRomanStr = Right(sRomanStr, Len(sRomanStr) - 2)
            End If
        End If
    End If
    If iTens = 0 Or iTens = 5 Then
        sStr = Left(sRomanStr, 1)
        While sStr = "X"
            iTens = iTens + 1
            sRomanStr = Right(sRomanStr, Len(sRomanStr) - 1)
            sStr = Left(sRomanStr, 1)
        Wend
    End If
    xDecimal = xDecimal + iTens * 10
    
    iOnes = 0
    If Left(sRomanStr, 2) = "IX" Then
        iOnes = 9
        sRomanStr = Right(sRomanStr, Len(sRomanStr) - 2)
    Else
        If Left(sRomanStr, 1) = "V" Then
            iOnes = 5
            sRomanStr = Right(sRomanStr, Len(sRomanStr) - 1)
        Else
            If Left(sRomanStr, 2) = "IV" Then
                iOnes = 4
                sRomanStr = Right(sRomanStr, Len(sRomanStr) - 2)
            End If
        End If
    End If
    If iOnes = 0 Or iOnes = 5 Then
        sStr = Left(sRomanStr, 1)
        While sStr = "I"
            iOnes = iOnes + 1
            sRomanStr = Right(sRomanStr, Len(sRomanStr) - 1)
            sStr = Left(sRomanStr, 1)
        Wend
    End If
    xDecimal = xDecimal + iOnes

    GetRomanDecNum = xDecimal
End Function

Public Function GetDecHexStr(ByVal xDecimal As Double) As String
    Dim sReturnStr As String
    Dim lQuotient As Long
        
    iRemainder = xDecimal Mod 16
    lQuotient = xDecimal \ 16
    
    While lQuotient > 0
        sReturnStr = sReturnStr + GetHexDigit(iRemainder)
        xDecimal = lQuotient
        iRemainder = xDecimal Mod 16
        lQuotient = xDecimal \ 16
    Wend
    If iRemainder > 0 Then
        sReturnStr = sReturnStr + GetHexDigit(iRemainder)
    End If
    GetDecHexStr = StrReverse(sReturnStr)
    
End Function

Public Function GetHexDigit(ByVal iDigit As Integer) As String
    Dim sReturnStr As String
    
    Select Case iDigit
    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
        sReturnStr = CStr(iDigit)
    Case "10"
        sReturnStr = "A"
    Case "11"
        sReturnStr = "B"
    Case "12"
        sReturnStr = "C"
    Case "13"
        sReturnStr = "D"
    Case "14"
        sReturnStr = "E"
    Case "15"
        sReturnStr = "F"
    End Select
    GetHexDigit = sReturnStr

End Function

Public Function GetDecOctStr(ByVal xDecimal As Double) As String

    Dim sReturnStr As String
    Dim lQuotient As Long
        
    iRemainder = xDecimal Mod 8
    lQuotient = xDecimal \ 8
    
    Do While lQuotient > 0
        sReturnStr = sReturnStr + CStr(iRemainder)
        xDecimal = lQuotient
        iRemainder = xDecimal Mod 8
        lQuotient = xDecimal \ 8
    Loop
    If iRemainder > 0 Then
        sReturnStr = sReturnStr + CStr(iRemainder)
    End If
    GetDecOctStr = StrReverse(sReturnStr)
    
End Function

Public Function GetDecBinStr(ByVal xDecimal As Double) As String

    Dim sReturnStr As String
    Dim lQuotient As Long
        
    iRemainder = xDecimal Mod 2
    lQuotient = xDecimal \ 2
    
    Do While lQuotient > 0
        sReturnStr = sReturnStr + CStr(iRemainder)
        xDecimal = lQuotient
        iRemainder = xDecimal Mod 2
        lQuotient = xDecimal \ 2
    Loop
    If iRemainder > 0 Then
        sReturnStr = sReturnStr + CStr(iRemainder)
    End If
    GetDecBinStr = StrReverse(sReturnStr)
    
End Function
Public Function GetBinDecNum(ByVal sBinStr As String) As Double
    Dim iLength, Counter As Integer
    Dim xReturnVal As Double
    Dim sNewString As String
        
    xReturnVal = 0
    iLength = Len(sBinStr)
    sNewString = StrReverse(sBinStr)
    For Counter = 0 To iLength - 1
        xReturnVal = xReturnVal + Left(sNewString, 1) * 2 ^ Counter
        sNewString = Right(sNewString, Len(sNewString) - 1)
    Next
    
    GetBinDecNum = xReturnVal

End Function

Public Sub ConvertOutputWindowText(ByVal iOldRadix As Integer, ByVal iNewRadix As Integer)
    Select Case (iOldRadix)
        Case ROMANNUM
            Select Case (iNewRadix)
                Case HEXNUM
                    xDecimal = GetRomanDecNum(OutputWindow.Text)
                    OutputWindow.Text = GetDecHexStr(xDecimal)
                Case DECNUM
                    xDecimal = GetRomanDecNum(OutputWindow.Text)
                    OutputWindow.Text = xDecimal
                Case OCTNUM
                    xDecimal = GetRomanDecNum(OutputWindow.Text)
                    OutputWindow.Text = GetDecOctStr(xDecimal)
                Case BINNUM
                    xDecimal = GetRomanDecNum(OutputWindow.Text)
                    OutputWindow.Text = GetDecBinStr(xDecimal)
            End Select
        Case HEXNUM
            Select Case (iNewRadix)
                Case ROMANNUM
                    xDecimal = Val("&H" + OutputWindow.Text)
                    OutputWindow.Text = GetDecRomanStr(xDecimal)
                Case DECNUM
                    OutputWindow.Text = Val("&H" + OutputWindow.Text)
                Case OCTNUM
                    OutputWindow.Text = Val("&H" + OutputWindow.Text)
                    OutputWindow.Text = GetDecOctStr(OutputWindow.Text)
                Case BINNUM
                    OutputWindow.Text = Val("&H" + OutputWindow.Text)
                    OutputWindow.Text = GetDecBinStr(OutputWindow.Text)
            End Select
        Case DECNUM
            Select Case (iNewRadix)
                Case ROMANNUM
                    OutputWindow.Text = GetDecRomanStr(OutputWindow.Text)
                Case HEXNUM
                    OutputWindow.Text = GetDecHexStr(OutputWindow.Text)
                Case OCTNUM
                    OutputWindow.Text = GetDecOctStr(OutputWindow.Text)
                Case BINNUM
                    OutputWindow.Text = GetDecBinStr(OutputWindow.Text)
            End Select
        Case OCTNUM
            Select Case (iNewRadix)
                Case ROMANNUM
                    xDecimal = Val("&O" + OutputWindow.Text)
                    OutputWindow.Text = GetDecRomanStr(xDecimal)
                Case HEXNUM
                    OutputWindow.Text = Val("&O" + OutputWindow.Text)
                    OutputWindow.Text = GetDecHexStr(OutputWindow.Text)
                Case DECNUM
                    OutputWindow.Text = Val("&O" + OutputWindow.Text)
                Case BINNUM
                    OutputWindow.Text = Val("&O" + OutputWindow.Text)
                    OutputWindow.Text = GetDecBinStr(OutputWindow.Text)
            End Select
        Case BINNUM
            Select Case (iNewRadix)
                Case ROMANNUM
                    xDecimal = GetBinDecNum(OutputWindow.Text)
                    OutputWindow.Text = GetDecRomanStr(xDecimal)
                Case HEXNUM
                    OutputWindow.Text = GetBinDecNum(OutputWindow.Text)
                    OutputWindow.Text = GetDecHexStr(OutputWindow.Text)
                Case DECNUM
                    OutputWindow.Text = GetBinDecNum(OutputWindow.Text)
                Case OCTNUM
                    OutputWindow.Text = GetBinDecNum(OutputWindow.Text)
                    OutputWindow.Text = GetDecOctStr(OutputWindow.Text)
            End Select
    End Select
End Sub

