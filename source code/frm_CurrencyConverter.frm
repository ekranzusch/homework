VERSION 5.00
Begin VB.Form frm_CurrencyConverter 
   Caption         =   "Currency Converter"
   ClientHeight    =   7845
   ClientLeft      =   3990
   ClientTop       =   2055
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_Exit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   13
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmd_Calculate 
      Caption         =   "&Calculate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Top             =   6840
      Width           =   1215
   End
   Begin VB.TextBox txt_American 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lbl_ValueSpanish 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   11
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label lbl_ValueGerman 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lbl_ValueItalian 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   9
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label lbl_ValueFrench 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lbl_ValueEnglish 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lbl_Spanish 
      Alignment       =   2  'Center
      Caption         =   "Spanish [Peseta]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label lbl_German 
      Alignment       =   2  'Center
      Caption         =   "German [Mark]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lbl_Italian 
      Alignment       =   2  'Center
      Caption         =   "Italian [Lira]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lbl_French 
      Alignment       =   2  'Center
      Caption         =   "French [Franc]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lbl_English 
      Alignment       =   2  'Center
      Caption         =   "English [Pound]"
      BeginProperty Font 
         Name            =   "Arial"
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
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lbl_American 
      Alignment       =   2  'Center
      Caption         =   "American [Dollar]"
      BeginProperty Font 
         Name            =   "Arial"
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
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frm_CurrencyConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_Calculate_Click()
' This text calculates the English value
' relative to the American value times current exchange rate
lbl_ValueEnglish = Round(Val(txt_American.Text) * 1.4903, 2)

' This text calculates the French value
' relative to the American value times current exchange rate
lbl_ValueFrench = Round(Val(txt_American.Text) * 6.91, 2)

' This text calculates the German value
' relative to the American value times current exchange rate
lbl_ValueGerman = Round(Val(txt_American.Text) * 2.0601, 2)

' This text calculates the Italian value
' relative to the American value times current exchange rate
lbl_ValueItalian = Round(Val(txt_American.Text) * 2040, 2)

' This text calculates the Spanish value
' relative to the American value times current exchange rate
lbl_ValueSpanish = Round(Val(txt_American.Text) * 175.1, 2)

'the with statement selects txt_american after calculation
'the with statement also selects the characters for easier data entry
With txt_American
    .SelStart = 0
    .SelLength = Len(.Text)
    .SetFocus
End With

End Sub


Private Sub cmd_Exit_Click()
'This command ends the program
End
End Sub
