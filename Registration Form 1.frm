VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   10695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   ScaleHeight     =   12915
   ScaleWidth      =   23760
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox UpDown1 
      Height          =   495
      Left            =   3960
      ScaleHeight     =   435
      ScaleWidth      =   195
      TabIndex        =   39
      Top             =   5400
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Upload Image"
      Height          =   495
      Left            =   14400
      TabIndex        =   38
      Top             =   3960
      Width           =   1815
   End
   Begin VB.PictureBox CommonDialog1 
      Height          =   480
      Left            =   16200
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   41
      Top             =   3960
      Width           =   1200
   End
   Begin VB.PictureBox DTPicker1 
      Height          =   375
      Left            =   2280
      ScaleHeight     =   315
      ScaleWidth      =   2235
      TabIndex        =   37
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RESET"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   36
      Top             =   12480
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   35
      Top             =   12480
      Width           =   2535
   End
   Begin VB.TextBox Text12 
      Height          =   495
      Left            =   2280
      TabIndex        =   34
      Top             =   11760
      Width           =   3255
   End
   Begin VB.TextBox Text11 
      Height          =   495
      Left            =   2280
      TabIndex        =   32
      Top             =   10920
      Width           =   3015
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Sports"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   30
      Top             =   10320
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "NSS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   29
      Top             =   10320
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "NCC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   28
      Top             =   10320
      Width           =   855
   End
   Begin VB.TextBox Text10 
      Height          =   495
      Left            =   2280
      TabIndex        =   26
      Top             =   9480
      Width           =   2535
   End
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   2280
      TabIndex        =   24
      Top             =   8640
      Width           =   2535
   End
   Begin VB.TextBox Text8 
      Height          =   975
      Left            =   2280
      TabIndex        =   22
      Top             =   7440
      Width           =   2655
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   5160
      TabIndex        =   20
      Top             =   6840
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2280
      TabIndex        =   18
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   2280
      TabIndex        =   16
      Top             =   6120
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2280
      TabIndex        =   14
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2280
      TabIndex        =   11
      Top             =   3480
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Registration Form 1.frx":0000
      Left            =   2280
      List            =   "Registration Form 1.frx":0016
      TabIndex        =   9
      Top             =   4200
      Width           =   2775
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   7
      Top             =   2760
      Width           =   1935
   End
   Begin VB.OptionButton option1 
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label Label13 
      Caption         =   "Help?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      TabIndex        =   40
      Top             =   12480
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   2535
      Left            =   14400
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label18 
      Caption         =   "College :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   33
      ToolTipText     =   "Enter the College Name"
      Top             =   11880
      Width           =   975
   End
   Begin VB.Label Label17 
      Caption         =   "Hobbies :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   31
      ToolTipText     =   "Enter Your Hobbies"
      Top             =   11040
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "Extra Ciricular :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   27
      ToolTipText     =   "Select Any Extra Cericular Activities"
      Top             =   10320
      Width           =   1695
   End
   Begin VB.Label Label15 
      Caption         =   "School Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   25
      ToolTipText     =   "Enter Your HSE School Name"
      Top             =   9480
      Width           =   1695
   End
   Begin VB.Label Label14 
      Caption         =   "Mobile no :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   23
      ToolTipText     =   "phone Number"
      Top             =   8640
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "  Address :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   21
      ToolTipText     =   "Enter your Resedental Address"
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Marks :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   19
      ToolTipText     =   "Enter Your HSE Marks"
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "    Cutt Off :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   17
      ToolTipText     =   "Your 12th Cutt Off"
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "         Email :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   15
      ToolTipText     =   "Enter Your Email Address"
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "           Age :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   13
      ToolTipText     =   "Enter Your Age"
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "           DOB :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   12
      ToolTipText     =   "Enter Your Date Of Birth"
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Parents Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Enter your Parents Name"
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Department :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   8
      ToolTipText     =   "Select Your Department"
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Gender :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   5
      ToolTipText     =   "Your Gender"
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Regno :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   3
      ToolTipText     =   "Enter Regno"
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   1
      ToolTipText     =   "Enter your Name"
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Students Registration Form"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   0
      Top             =   360
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Text1.Text = nil Then
MsgBox "Please Enter Name*", vbInformation, "Info*"
ElseIf Text2.Text = nil Then
MsgBox "Please Enter Regno*", vbInformation, "Info*"
ElseIf option1.Value = False And Option2.Value = False Then
MsgBox "Please Select Gender*", vbInformation, "Info*"
ElseIf Text3.Text = nil Then
MsgBox "Please Enter Parents Name*", vbInformation, "Info*"
ElseIf Text4.Text = nil Then
MsgBox "Please Enter Your Age*", vbInformation, "Info*"
ElseIf Text5.Text = nil Then
MsgBox "Please Enter EmailID*", vbInformation, "Info*"
ElseIf Text6.Text = nil Then
MsgBox "Please Enter your HSE Marks*", vbInformation, "Info*"
ElseIf Text7.Text = nil Then
MsgBox "Please Enter Your cuttoff Mark*", vbInformation, "Info*"
ElseIf Text8.Text = nil Then
MsgBox "Please Enter Your Residental Address*", vbInformation, "Info*"
ElseIf Text9.Text = nil Then
MsgBox "Please Enter MobileNo*", vbInformation, "info*"
ElseIf Text10.Text = nil Then
MsgBox "Please Enter your school Name*", vbInformation, "Info*"
ElseIf Check1.Value = False And Check2.Value = False And Check3.Value = False Then
MsgBox "Please Choose ExtraCircular Activities", vbInformation, "Info*"
ElseIf Text11.Text = nil Then
MsgBox "please Enter Hobbies*", vbInformation, "Info*"
ElseIf Text12.Text = nil Then
MsgBox "Please Enter college Name*", vbInformation, "Info*"
Else
MsgBox "Details Saved.click ok to enter another submission"
End If
End Sub

Private Sub Command2_Click()
If MsgBox("Are You Want to Refill this Form", vbYesNo) = vbYes Then
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Text6.Text = " "
Text7.Text = " "
Text8.Text = " "
Text9.Text = " "
Text10.Text = " "
Text11.Text = " "
Text12.Text = " "
End If
End Sub

Private Sub Command3_Click()
CommonDialog1.ShowOpen
Image1.Picture = LoadPicture(CommonDialog1.FileName)
End Sub

Private Sub Image1_Click()
CommonDialog1.ShowOpen
Image1.Picture = LoadPicture(CommonDialog1.FileName)

End Sub

Private Sub Label13_Click()
MsgBox "You Need Any Help?"
End Sub

Private Sub UpDown1_DownClick()
If Text4.Text = nil Then
Text4.Text = 0
Text4.Text = Text4.Text - 1
Else
Text4.Text = Text4.Text - 1
End If
End Sub

Private Sub UpDown1_UpClick()
If Text4.Text = nil Then
Text4.Text = 0
Text4.Text = Text4.Text + 1
Else
Text4.Text = Text4.Text + 1
End If
End Sub
