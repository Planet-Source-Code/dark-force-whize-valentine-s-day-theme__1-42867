VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   315
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSet 
      Caption         =   "Set"
      Height          =   225
      Left            =   2565
      TabIndex        =   4
      Top             =   15
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.TextBox txtMinutes 
      Height          =   240
      Left            =   1575
      TabIndex        =   3
      Text            =   "01"
      Top             =   15
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Popup Message after         Minutes."
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   2520
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   525
      Left            =   0
      Top             =   0
      Width           =   700
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      X1              =   -60
      X2              =   1215
      Y1              =   255
      Y2              =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   " Exit       "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   30
      TabIndex        =   1
      Top             =   285
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   " Options "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   690
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      X1              =   -60
      X2              =   1215
      Y1              =   255
      Y2              =   270
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSet_Click()
PopupInterval = Val(txtMinutes)
Me.Hide
End Sub

Private Sub Label1_Click()
Label1.Visible = False
Label2.Visible = False
Line1.Visible = False
Line2.Visible = False
Shape1.Visible = False
Label3.Visible = True
txtMinutes.Visible = True
cmdSet.Visible = True
Me.Width = 3405
Me.Height = 360
Me.BackColor = vbBlack
Me.Top = WindowDimensions.Bottom * Screen.TwipsPerPixelY - 800
Me.Left = WindowDimensions.Right * Screen.TwipsPerPixelX - 3500
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.BackColor = &H8000000D
Label1.ForeColor = vbWhite
Label2.BackColor = vbWhite
Label2.ForeColor = vbBlack
End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = &H8000000D
Label2.ForeColor = vbWhite
Label1.BackColor = vbWhite
Label1.ForeColor = vbBlack
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
Unload frmHeart
Unload frmMenu
End If
End Sub
