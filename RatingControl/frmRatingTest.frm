VERSION 5.00
Begin VB.Form frmRating 
   Caption         =   " Rating Example"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   3240
   StartUpPosition =   3  'Windows Default
   Begin Project1.ucRating ucRating1 
      Height          =   435
      Left            =   240
      TabIndex        =   5
      Top             =   1050
      Width           =   2340
      _extentx        =   3519
      _extenty        =   767
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   1350
      TabIndex        =   4
      Top             =   2100
      Width           =   1125
   End
   Begin VB.TextBox txtStars 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   900
      TabIndex        =   3
      Text            =   "5"
      Top             =   2100
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Stars:"
      Height          =   195
      Left            =   390
      TabIndex        =   2
      Top             =   2130
      Width           =   405
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   1373
      TabIndex        =   1
      Top             =   510
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Value: 0"
      Height          =   195
      Left            =   1328
      TabIndex        =   0
      Top             =   300
      Width           =   585
   End
End
Attribute VB_Name = "frmRating"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  ucRating1.MaxStars = Val(txtStars)
End Sub

Private Sub UserControl11_Click()
  Label1.Caption = "Value: " & ucRating1.Value
End Sub

Private Sub UserControl11_MouseLeave()
  Label2 = "Over_Value: N/A"
End Sub

Private Sub UserControl11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label2 = "Over_Value: " & ucRating1.OverValue
End Sub

