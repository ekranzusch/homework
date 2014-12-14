VERSION 5.00
Begin VB.Form frm_MathOperations 
   Caption         =   "Math Operations"
   ClientHeight    =   5355
   ClientLeft      =   3195
   ClientTop       =   2430
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   6690
   Begin VB.CommandButton cmd_Calculate 
      Caption         =   "&Calculate"
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CheckBox chk_Division 
      Caption         =   "Division"
      Height          =   495
      Left            =   5040
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CheckBox chk_Multiplication 
      Caption         =   "Multiplication"
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CheckBox chk_Subtraction 
      Caption         =   "Subtraction"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CheckBox chk_Addition 
      Caption         =   "Addition"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txt_Variable2 
      Height          =   615
      Left            =   3960
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txt_Variable1 
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lbl_Division 
      Caption         =   "Division"
      Height          =   495
      Left            =   5040
      TabIndex        =   10
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lbl_Multiplication 
      Caption         =   "Multiplication"
      Height          =   495
      Left            =   3480
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lbl_Subtraction 
      Caption         =   "Subtraction"
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lbl_Addition 
      Caption         =   "Addition"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "frm_MathOperations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Calculate_Click()

' Checks to see if Addition box is checked
' If so, it runs values through addition equation
If chk_Addition.Value = 1 Then
    lbl_Addition.Caption = Round((Val(txt_Variable1.Text) + (txt_Variable2.Text)), 2)
Else
    lbl_Addition.Caption = "Addition"
End If


' Checks to see if Subtraction box is checked
' If so, it runs values through subtraction equation
If chk_Subtraction.Value = 1 Then
    lbl_Subtraction.Caption = Round((Val(txt_Variable1.Text) - (txt_Variable2.Text)), 2)
Else
    lbl_Subtraction.Caption = "Subtraction"
End If


' Checks to see if Multiplication box is checked
' If so, it runs values through multiplication equation
If chk_Multiplication.Value = 1 Then
    lbl_Multiplication.Caption = Round((Val(txt_Variable1.Text) * (txt_Variable2.Text)), 2)
Else
    lbl_Multiplication.Caption = "Multiplication"
End If


' Checks to see if Division box is checked
' If so, checks to see if denominator = 0
' If denominator = 0, error message is displayed
If chk_Division.Value = 1 Then
    If txt_Variable2.Text <> 0 Then
    lbl_Division.Caption = Round((Val(txt_Variable1.Text) / (txt_Variable2.Text)), 2)
' With statement sets cursor in box 1 if program runs successfully
        With txt_Variable1
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    End If
    If txt_Variable2.Text = 0 Then
    MsgBox "Cannot Divide by Zero", vbOKOnly, "Logic Error"
' With statment sets cursor in box 2 if program has error
        With txt_Variable2
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    End If
Else
    lbl_Division.Caption = "Division"
' With statement sets cursor to box 1 if program runs successfully and Division box is not used
    With txt_Variable1
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With

End If
End Sub
