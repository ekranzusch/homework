VERSION 5.00
Begin VB.Form frm_Converter 
   Caption         =   "Convert Miles to Kilometers"
   ClientHeight    =   4305
   ClientLeft      =   2835
   ClientTop       =   1680
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_Exit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmd_Convert 
      Caption         =   "&Convert"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txt_Kilometers 
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txt_Miles 
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lbl_Kilometers 
      Caption         =   "KILOMETERS"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lbl_Miles 
      Caption         =   "MILES"
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
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "frm_Converter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Convert_Click()
Rem this code changes miles value into kilometers
txt_Kilometers = Val(txt_Miles) * 1.61

End Sub

Private Sub cmd_Exit_Click()
Rem This code ends the program
End
End Sub
