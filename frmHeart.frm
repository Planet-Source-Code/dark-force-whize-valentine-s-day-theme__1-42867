VERSION 5.00
Begin VB.Form frmHeart 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   3300
   ClientLeft      =   2610
   ClientTop       =   1365
   ClientWidth     =   3300
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmHeart.frx":0000
   ScaleHeight     =   220
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   220
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrInterval 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2250
      Top             =   705
   End
   Begin VB.Timer tmrTrayIcon 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1905
      Top             =   1545
   End
   Begin VB.PictureBox picTrayIcon 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   2520
      ScaleHeight     =   315
      ScaleWidth      =   330
      TabIndex        =   0
      Top             =   495
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Timer tmrHeart 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1050
      Top             =   1410
   End
   Begin VB.Label lblQuote 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFF00&
      Height          =   1755
      Left            =   735
      TabIndex        =   1
      Top             =   600
      Width           =   2355
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgTrayIcon 
      Height          =   480
      Index           =   1
      Left            =   480
      Picture         =   "frmHeart.frx":23774
      Top             =   420
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTrayIcon 
      Height          =   480
      Index           =   0
      Left            =   495
      Picture         =   "frmHeart.frx":23A7E
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmHeart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ThanX for downloading this project. If u like it, plz dont forget to vote.
'Special Thanx to Doug Gaede for his transparent shape class. It was very
'helpful. Thanx buddy :-)
'***********************************************************************************
'** Project      : Whize Valentine's Day Theme                                    **
'** Author       : Zubair Ahmed M.                                                **
'** Description  : This program creates a heart shaped form and let it floats or
'rather bounce of the screen edges. The heart also displays a love quote. The heart
'can be minimized to the system tray by left clicking on it and can be made to stand
'still y right clicking. Once minimized the system tray, it can be restored by
'double clicking. U can right click the form when minimzed for options.U can set the
'popup interval when the heart must popup again.
'***********************************************************************************
'The Quote Type for getting the Quotes from a file.
Private Type Quote
Quotes As String * 200
End Type

'Constants for setting the form as the top most
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

'constants used for minimizing the form to the sys tray
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

'Mouse click constants
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_RBUTTONUP = &H205

Dim Displacement As Integer
Dim MoveDirection As Integer
Dim TrayIcon As NOTIFYICONDATA
Dim filenum As Integer
Dim MyQuotes As Quote
Dim FormStill As Boolean, setStillBy As Integer
Dim TotalQuotes As Integer

Private Sub Form_Load()
PopupInterval = 1
filenum = FreeFile
Open App.Path & "\Quotes.dat" For Random As #filenum Len = Len(MyQuotes)
Do While Not EOF(filenum) 'Count for no. of quotes
Get #filenum, , MyQuotes
TotalQuotes = TotalQuotes + 1
Loop
Set ShapeTheForm = New clsTransForm 'Make the form a heart shape
ShapeTheForm.ShapeMe frmHeart, RGB(255, 255, 255) 'Transparent color is white.
MoveDirection = 1
GetWindowRect GetDesktopWindow, WindowDimensions
tmrHeart.Enabled = True
Displacement = 10
TrayIcon.cbSize = Len(TrayIcon)
TrayIcon.hWnd = picTrayIcon.hWnd
TrayIcon.uId = 1&
TrayIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
TrayIcon.ucallbackMessage = WM_LBUTTONDOWN
TrayIcon.szTip = "Valentine Day Theme" & vbCr & "DoubleClick to open." & vbCr & "RightClick for options."
GetQuote
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then 'If left click minimize form
frmHeart.WindowState = 1
tmrHeart.Enabled = False
tmrInterval.Enabled = True
End If
If Button = vbRightButton And FormStill = False Then 'If right click stop the form
tmrHeart.Enabled = False
FormStill = True
ElseIf Button <> vbLeftButton Then
tmrHeart.Enabled = True
FormStill = False
End If
End Sub

Private Sub Form_Resize()
If WindowState = 1 Then
TrayIcon.hIcon = imgTrayIcon(1).Picture
Shell_NotifyIcon NIM_ADD, TrayIcon
tmrTrayIcon.Enabled = True
tmrHeart.Enabled = False
End If
End Sub

Private Sub lblQuote_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then 'If left click minimize form
frmHeart.WindowState = 1
tmrHeart.Enabled = False
tmrInterval.Enabled = True
End If
If Button = vbRightButton And FormStill = False Then 'If right click stop the form
tmrHeart.Enabled = False
FormStill = True
ElseIf Button <> vbLeftButton Then
tmrHeart.Enabled = True
FormStill = False
End If
End Sub

Private Sub picTrayIcon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Msg = X / Screen.TwipsPerPixelX
    If Msg = WM_LBUTTONDBLCLK Then 'If icon is double-clicked in sys tray
        Shell_NotifyIcon NIM_DELETE, TrayIcon
        frmHeart.WindowState = 0
        tmrTrayIcon.Enabled = False
        tmrHeart.Enabled = True
        SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    ElseIf Msg = WM_RBUTTONUP Then  'If icon is right-clicked in sys tray
        frmMenu.Top = WindowDimensions.Bottom * Screen.TwipsPerPixelY - 1000
        frmMenu.Left = WindowDimensions.Right * Screen.TwipsPerPixelX - 2000
        frmMenu.Label1.Visible = True
        frmMenu.Label2.Visible = True
        frmMenu.Line1.Visible = True
        frmMenu.Line2.Visible = True
        frmMenu.Width = 705
        frmMenu.Height = 525
        frmMenu.BackColor = vbWhite
        frmMenu.Label3.Visible = False
        frmMenu.txtMinutes.Visible = False
        frmMenu.cmdSet.Visible = False
        frmMenu.Show
        frmMenu.SetFocus
    End If
End Sub

Private Sub tmrHeart_Timer()
'This procedure checks the direction from where the form came and bounces it
'accordingly.
If MoveDirection = 1 Then
frmHeart.Left = Left + Displacement
frmHeart.Top = Top - Displacement
If (frmHeart.Top / Screen.TwipsPerPixelY) < -12 Then MoveDirection = 2
If (frmHeart.Left / Screen.TwipsPerPixelX) > 583 Then MoveDirection = 4
End If

If MoveDirection = 2 Then
frmHeart.Left = Left + Displacement
frmHeart.Top = Top + Displacement
If (frmHeart.Left / Screen.TwipsPerPixelX) > 583 Then MoveDirection = 3
If (frmHeart.Top / Screen.TwipsPerPixelY) > 362 Then MoveDirection = 1
End If

If MoveDirection = 3 Then
frmHeart.Left = Left - Displacement
frmHeart.Top = Top + Displacement
If (frmHeart.Top / Screen.TwipsPerPixelY) > 362 Then MoveDirection = 4
If (frmHeart.Left / Screen.TwipsPerPixelX) < 0 Then MoveDirection = 2
End If

If MoveDirection = 4 Then
frmHeart.Left = Left - Displacement
frmHeart.Top = Top - Displacement
If (frmHeart.Top / Screen.TwipsPerPixelY) < -12 Then MoveDirection = 5
If (frmHeart.Left / Screen.TwipsPerPixelX) < 0 Then MoveDirection = 1
End If

If MoveDirection = 5 Then
frmHeart.Left = Left - Displacement
frmHeart.Top = Top + Displacement
If (frmHeart.Left / Screen.TwipsPerPixelX) < 0 Then MoveDirection = 6
If (frmHeart.Top / Screen.TwipsPerPixelY) > 362 Then MoveDirection = 4
End If

If MoveDirection = 6 Then
frmHeart.Left = Left + Displacement
frmHeart.Top = Top + Displacement
If (frmHeart.Top / Screen.TwipsPerPixelY) > 362 Then MoveDirection = 1
If (frmHeart.Left / Screen.TwipsPerPixelX) > 583 Then MoveDirection = 5
End If

End Sub

Private Sub tmrInterval_Timer()
'This procedure checks if it is time to popup the form.
Static Interval As Integer
Interval = Interval + 1
If Interval = PopupInterval * 60 Then  'PopupInterval * 60 seconds
frmHeart.WindowState = 0
tmrHeart.Enabled = True
tmrInterval.Enabled = False
Interval = 0
GetQuote
End If
End Sub

Private Sub tmrTrayIcon_Timer()
'this procedure Animates the system tray icon
    Static Tek As Integer
    Me.Icon = imgTrayIcon(Tek).Picture
    TrayIcon.hIcon = imgTrayIcon(Tek).Picture
    Tek = Tek + 1
    If Tek = 2 Then Tek = 0
    Shell_NotifyIcon NIM_MODIFY, TrayIcon
End Sub
Private Sub Form_Unload(Cancel As Integer)
    TrayIcon.cbSize = Len(TrayIcon)
    TrayIcon.hWnd = picTrayIcon.hWnd
    TrayIcon.uId = 1&
    Shell_NotifyIcon NIM_DELETE, TrayIcon
    Set ShapeTheForm = Nothing
    Close #filenum
    End
End Sub
Private Sub GetQuote()
'This procedure gets the quote from the quotes.dat file
Static QuotesCount As Integer
Get #filenum, QuotesCount + 1, MyQuotes.Quotes
lblQuote.Caption = MyQuotes.Quotes
QuotesCount = QuotesCount + 1
If QuotesCount + 1 = TotalQuotes Then QuotesCount = 0
End Sub
