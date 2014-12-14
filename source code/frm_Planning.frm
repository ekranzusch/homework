VERSION 5.00
Begin VB.Form frm_Planning 
   Caption         =   "Budget Planning"
   ClientHeight    =   3492
   ClientLeft      =   1260
   ClientTop       =   1392
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   3492
   ScaleWidth      =   6240
   Begin VB.CommandButton cmd_Exit 
      Caption         =   "Exit"
      Height          =   372
      Left            =   3360
      TabIndex        =   4
      Top             =   2400
      Width           =   1092
   End
   Begin VB.TextBox txt_LastName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   3120
      TabIndex        =   3
      Top             =   1320
      Width           =   1572
   End
   Begin VB.TextBox txt_FirstName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   3120
      TabIndex        =   1
      Top             =   480
      Width           =   1572
   End
   Begin VB.Label lbl_LastName 
      Caption         =   "Enter Last Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   2052
   End
   Begin VB.Label lbl_FirstName 
      Caption         =   "Enter First Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2052
   End
End
Attribute VB_Name = "frm_Planning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Exit_Click()
Rem This piece of code will exit or end application
End
End Sub

Private Sub Form_Load()
Rem
End Sub


