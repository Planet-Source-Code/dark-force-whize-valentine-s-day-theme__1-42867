VERSION 5.00
Begin VB.Form frmQuotes 
   Caption         =   "Valentine Day Theme"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2955
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   30
      TabIndex        =   2
      Top             =   2415
      Width           =   3420
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   30
      TabIndex        =   1
      Top             =   1890
      Width           =   3420
   End
   Begin VB.TextBox txtQuotes 
      Height          =   1845
      Left            =   30
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   15
      Width           =   3435
   End
End
Attribute VB_Name = "frmQuotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Quote
Quotes As String * 200
End Type
Dim MyQuotes As Quote
Dim filenum As Integer

Private Sub cmdClear_Click()
txtQuotes.Text = Clear
End Sub

Private Sub cmdSave_Click()
MyQuotes.Quotes = txtQuotes.Text
Print #filenum, MyQuotes.Quotes
txtQuotes.Text = Clear
End Sub

Private Sub Form_Load()
filenum = FreeFile
Open App.Path & "\Quotes.dat" For Append Access Write As #filenum Len = Len(MyQuotes)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Close #filenum
End Sub
