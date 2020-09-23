VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5670
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5355
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   378
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   357
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   5565
      TabIndex        =   1
      Top             =   0
      Width           =   5565
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   4950
         TabIndex        =   4
         Top             =   15
         Width           =   135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   5145
         TabIndex        =   3
         Top             =   -30
         Width           =   150
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Things To Do"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   60
         TabIndex        =   2
         Top             =   15
         Width           =   1020
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   315
      Width           =   5220
   End
   Begin VB.Line Line1 
      X1              =   1
      X2              =   357
      Y1              =   17
      Y2              =   17
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Sub Form_Initialize()
AlwaysOnBottom Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Open App.Path & "\data.dat" For Input As #1
Text1.Text = LTrim(RTrim(Input(LOF(1), 1)))
Close #1

Open App.Path & "\settings.dat" For Input As #1
varset = Split(Input(LOF(1), 1), "|")
Me.Left = varset(0)
Me.Top = varset(1)
Close #1

Me.Show
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AlwaysOnBottom Me
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
AlwaysOnBottom Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveSettings

End
End Sub

Sub SaveSettings()
On Error Resume Next
Open App.Path & "\data.dat" For Output As #1
Print #1, Text1.Text
Close #1

Open App.Path & "\settings.dat" For Output As #1
Print #1, Me.Left & "|" & Me.Top
Close #1
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AlwaysOnBottom Me
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
AlwaysOnBottom Me
End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
AlwaysOnBottom Me
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AlwaysOnBottom Me
End Sub

Private Sub Label3_Click()
SaveSettings
Label3.ForeColor = RGB(180, 0, 0)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AlwaysOnBottom Me
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
AlwaysOnBottom Me
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AlwaysOnBottom Me
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Call ReleaseCapture
lngReturnValue = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
AlwaysOnBottom Me
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
AlwaysOnBottom Me

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Label3.ForeColor = RGB(235, 235, 235)
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
AlwaysOnBottom Me
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
AlwaysOnBottom Me
End Sub
