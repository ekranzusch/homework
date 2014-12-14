VERSION 5.00
Begin VB.Form frm_MaturityCalculator 
   Caption         =   "Maturity Calculator"
   ClientHeight    =   6120
   ClientLeft      =   3510
   ClientTop       =   2205
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   5910
   Begin VB.CommandButton cmd_CalculateMaturity 
      Caption         =   "&Calculate"
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
      Left            =   3000
      TabIndex        =   8
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox txt_MaturityValue 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   4920
      Width           =   2415
   End
   Begin VB.TextBox txt_NumberOfYears 
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
      Left            =   3000
      TabIndex        =   6
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox txt_InvestmentRate 
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
      Left            =   3000
      TabIndex        =   5
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox txt_Investment 
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
      Left            =   3000
      TabIndex        =   4
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lbl_MaturityValue 
      Caption         =   "Maturity Value [$]"
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
      TabIndex        =   3
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label lbl_NumberOfYears 
      Caption         =   "Number of Years"
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
      TabIndex        =   2
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label lbl_InvsetmentRate 
      Caption         =   "Investment Rate [%]"
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
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lbl_Investment 
      Caption         =   "Investment [$]"
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
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "frm_MaturityCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_CalculateMaturity_Click()

'this line of text calculates maturity value
txt_MaturityValue.Text = Round(Val(txt_Investment.Text) * (1 + Val(txt_InvestmentRate.Text) / 400) ^ (4 * Val(txt_NumberOfYears.Text)), 2)

'with statement selects txt_Investment after calculation,
'then highlights text for easier data entry

With txt_Investment
    .SelStart = 0
    .SelLength = Len(.Text)
    .SetFocus
End With

End Sub

