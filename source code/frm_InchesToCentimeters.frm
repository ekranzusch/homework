VERSION 5.00
Begin VB.Form frm_InchesToCentimeters 
   Caption         =   "Inches to Centimeters"
   ClientHeight    =   4485
   ClientLeft      =   2325
   ClientTop       =   2640
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   4485
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_Exit 
      Caption         =   "E&xit"
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
      Left            =   2640
      TabIndex        =   9
      Top             =   3720
      Width           =   1215
   End
   Begin VB.HScrollBar hsb_Converter 
      Height          =   375
      Left            =   2040
      Max             =   100
      TabIndex        =   8
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Label lblComments 
      BackColor       =   &H0080FFFF&
      Caption         =   $"frm_InchesToCentimeters.frx":0000
      Height          =   1095
      Left            =   480
      TabIndex        =   10
      Top             =   4680
      Width           =   5415
   End
   Begin VB.Label lbl_CentimetersMaximum 
      Alignment       =   2  'Center
      Caption         =   "254"
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
      Left            =   5280
      TabIndex        =   7
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label lbl_InchesMaximum 
      Alignment       =   2  'Center
      Caption         =   "100"
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
      Left            =   5280
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lbl_CentimtersVariable 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      TabIndex        =   5
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label lbl_InchesVariable 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lbl_CentimetersMinimum 
      Caption         =   "0"
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
      Left            =   2040
      TabIndex        =   3
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label lbl_InchesMinimum 
      Caption         =   "0"
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
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lbl_Centimeters 
      Caption         =   "CENTIMETERS"
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
      Left            =   360
      TabIndex        =   1
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lbl_Inches 
      Caption         =   "INCHES"
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
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "frm_InchesToCentimeters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label7_Click()

End Sub

Private Sub HScroll1_Change()

End Sub

Private Sub cmd_Exit_Click()
' This code defines the button which ends the program
End

End Sub

Private Sub hsb_Converter_Change()

' This line of code gives lbl_InchesVariable a value
lbl_InchesVariable = hsb_Converter.Value

' This line of code gives lbl_CentimetersVariable a value
' which converts it from inch form to centimeter form.

lbl_CentimtersVariable = hsb_Converter.Value * 2.54

End Sub

Private Sub hsb_Converter_Scroll()
' This line of code gives lbl_InchesVariable a value
lbl_InchesVariable = hsb_Converter.Value

' This line of code gives lbl_CentimetersVariable a value
' which converts it from inch form to centimeter form.

lbl_CentimtersVariable = hsb_Converter.Value * 2.54
End Sub
