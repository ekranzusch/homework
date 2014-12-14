VERSION 5.00
Begin VB.Form frm_ShapesAndColors 
   Caption         =   "Shapes and Colors"
   ClientHeight    =   7335
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   9165
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton opt_BlueCircle 
      Caption         =   "Blue"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   2280
      Width           =   855
   End
   Begin VB.OptionButton opt_YellowCircle 
      Caption         =   "Yellow"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin VB.OptionButton opt_RedCircle 
      Caption         =   "Red"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   2280
      Width           =   855
   End
   Begin VB.Frame fra_CircleColors 
      Caption         =   "Circle Colors"
      Height          =   975
      Left            =   480
      TabIndex        =   15
      Top             =   2040
      Width           =   3495
   End
   Begin VB.OptionButton opt_BlueRectangle 
      Caption         =   "Blue"
      Height          =   495
      Left            =   7440
      TabIndex        =   11
      Top             =   5520
      Width           =   735
   End
   Begin VB.OptionButton opt_YellowRectangle 
      Caption         =   "Yellow"
      Height          =   495
      Left            =   6360
      TabIndex        =   10
      Top             =   5520
      Width           =   855
   End
   Begin VB.OptionButton opt_RedRectangle 
      Caption         =   "Red"
      Height          =   495
      Left            =   5280
      TabIndex        =   9
      Top             =   5520
      Width           =   615
   End
   Begin VB.Frame fra_RectangleColors 
      Caption         =   "Rectangle Colors"
      Height          =   975
      Left            =   5040
      TabIndex        =   14
      Top             =   5280
      Width           =   3495
   End
   Begin VB.OptionButton opt_BlueOval 
      Caption         =   "Blue"
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   5520
      Width           =   615
   End
   Begin VB.OptionButton opt_YellowOval 
      Caption         =   "Yellow"
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   5520
      Width           =   855
   End
   Begin VB.OptionButton opt_RedOval 
      Caption         =   "Red"
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   5520
      Width           =   615
   End
   Begin VB.Frame fra_OvalColors 
      Caption         =   "Oval Colors"
      Height          =   975
      Left            =   480
      TabIndex        =   13
      Top             =   5280
      Width           =   3495
   End
   Begin VB.OptionButton opt_BlueSquare 
      Caption         =   "Blue"
      Height          =   495
      Left            =   7440
      TabIndex        =   5
      Top             =   2280
      Width           =   735
   End
   Begin VB.OptionButton opt_YellowSquare 
      Caption         =   "Yellow"
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   2280
      Width           =   855
   End
   Begin VB.OptionButton opt_RedSquare 
      Caption         =   "Red"
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   2280
      Width           =   735
   End
   Begin VB.Frame fra_SquareColors 
      Caption         =   "Square Colors"
      Height          =   975
      Left            =   5040
      TabIndex        =   12
      Top             =   2040
      Width           =   3495
   End
   Begin VB.Shape shp_Rectangle 
      FillColor       =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   5640
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Shape shp_Oval 
      FillColor       =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   1200
      Shape           =   2  'Oval
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Shape shp_Square 
      FillColor       =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   6120
      Shape           =   1  'Square
      Top             =   360
      Width           =   1335
   End
   Begin VB.Shape shp_Circle 
      FillColor       =   &H80000005&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   1560
      Shape           =   3  'Circle
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frm_ShapesAndColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub opt_RedCircle_Click()
shp_Circle.FillColor = vbRed
End Sub

Private Sub opt_YellowCircle_Click()
shp_Circle.FillColor = vbYellow
End Sub

Private Sub opt_BlueCircle_Click()
shp_Circle.FillColor = vbBlue
End Sub

Private Sub opt_RedSquare_Click()
shp_Square.FillColor = vbRed
End Sub

Private Sub opt_YellowSquare_Click()
shp_Square.FillColor = vbYellow
End Sub

Private Sub opt_BlueSquare_Click()
shp_Square.FillColor = vbBlue
End Sub

Private Sub opt_RedOval_Click()
shp_Oval.FillColor = vbRed
End Sub

Private Sub opt_YellowOval_Click()
shp_Oval.FillColor = vbYellow
End Sub

Private Sub opt_BlueOval_Click()
shp_Oval.FillColor = vbBlue
End Sub

Private Sub opt_RedRectangle_Click()
shp_Rectangle.FillColor = vbRed
End Sub

Private Sub opt_YellowRectangle_Click()
shp_Rectangle.FillColor = vbYellow
End Sub

Private Sub opt_BlueRectangle_Click()
shp_Rectangle.FillColor = vbBlue
End Sub
