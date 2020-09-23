VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "3D Lines"
   ClientHeight    =   7560
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   28
      Top             =   7305
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3598
            MinWidth        =   3598
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3246
            MinWidth        =   3246
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4657
            MinWidth        =   4657
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6480
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   2280
      TabIndex        =   18
      Top             =   6480
      Width           =   4095
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3000
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   255
         LargeChange     =   20
         Left            =   720
         SmallChange     =   5
         TabIndex        =   20
         Top             =   240
         Value           =   30
         Width           =   375
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   255
         LargeChange     =   5
         Left            =   2640
         TabIndex        =   19
         Top             =   240
         Value           =   1
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Step:"
         Height          =   255
         Left            =   2160
         TabIndex        =   24
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Scale:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Sample Focus Effect"
      Height          =   1455
      Left            =   6480
      TabIndex        =   14
      Top             =   4680
      Width           =   2055
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         Max             =   3
         SelStart        =   2
         Value           =   2
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Go"
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Speed"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Colors"
      Height          =   1815
      Left            =   6480
      TabIndex        =   11
      Top             =   2400
      Width           =   2175
      Begin VB.OptionButton Option1 
         Caption         =   "Option3"
         Height          =   615
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Random Colors"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Fade Effect"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Shape Shape2 
         Height          =   375
         Left            =   840
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Background"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   795
         Width           =   855
      End
      Begin VB.Shape Shape4 
         Height          =   375
         Left            =   840
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Lines"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   315
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   8040
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose Parts"
      Height          =   1455
      Left            =   6480
      TabIndex        =   2
      Top             =   480
      Width           =   2175
      Begin VB.CheckBox Check4 
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         Height          =   255
         Left            =   1440
         TabIndex        =   5
         Top             =   480
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   1080
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Up-Right"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Up-Left"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Down-Right"
         Height          =   255
         Left            =   1080
         TabIndex        =   8
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Down-Left"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Draw"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   6600
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0080&
      Height          =   6255
      Left            =   120
      ScaleHeight     =   6195
      ScaleWidth      =   6195
      TabIndex        =   1
      Top             =   120
      Width           =   6255
      Begin VB.Image Bart 
         Height          =   1095
         Left            =   2400
         Top             =   2280
         Width           =   1455
      End
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu save 
         Caption         =   "&Save Bmp"
         Shortcut        =   {F3}
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim z As Double, st As Double, w As Double
Dim i As Double, j As Double, ex As Double
Dim col As Long
Dim Fade As Integer, gr As Integer, k As Integer

Private Sub Command1_Click()
    On Error Resume Next
    Picture1.Cls
    
    z = CDbl(Text1.Text)
    st = CDbl(Text2.Text)
    
    Me.SetTheScale Me.Picture1, 0, z, z, 0
    
    Picture1.AutoRedraw = True  'Enables the AutoRedraw and the save bitmap utility
    Draw_Lines
    Command1.SetFocus
    Status_Bar
End Sub

Function SetTheScale(ByVal obj As Object, ByVal upper_left_x As Single, _
ByVal upper_left_y As Single, ByVal lower_right_x As Single, ByVal lower_right_y As Single)
    obj.ScaleLeft = upper_left_x
    obj.ScaleTop = upper_left_y
    obj.ScaleWidth = lower_right_x - upper_left_x
    obj.ScaleHeight = lower_right_y - upper_left_y
End Function

Private Sub Command2_Click()
Dim sl_val, sp1, sp2, sp3 As Double
sl_val = Slider1.Value
z = CDbl(Text1.Text)

sp1 = z / 15000             'The speed of the sample focus
sp2 = z / 4285.71428571429
sp3 = z / 2000

If sl_val = 1 Then Text2.Text = sp1
If sl_val = 2 Then Text2.Text = sp2
If sl_val = 3 Then Text2.Text = sp3
Picture1.AutoRedraw = False         'It dissbles the AutoRedraw to enable the
Command1_Click                      'focus effect
Text2.Text = 1

End Sub

Private Sub Form_Load()
Text1.Text = VScroll1.Value
Text2.Text = VScroll2.Value
col = 16744576  'The default lines color
End Sub

Private Sub Form_Resize()
On Error Resume Next
Picture1.Height = Me.ScaleHeight - 1000
Picture1.Width = Me.ScaleWidth - 2600
Command1.Top = Picture1.Height + 325
Frame4.Top = Picture1.Height + 205
Frame1.Left = Picture1.Width + 225
Frame2.Left = Picture1.Width + 225
Frame3.Left = Picture1.Width + 225
End Sub

Private Sub Label7_Click()
    On Error GoTo ColorErr
    dlgCommon.Flags = cdlCCRGBInit
    dlgCommon.Color = col
    dlgCommon.ShowColor
    
    col = dlgCommon.Color
    Command1_Click
    Command1.SetFocus
ColorErr:
    Exit Sub

End Sub

Private Sub Label8_Click()
    On Error Resume Next
    dlgCommon.Flags = cdlCCRGBInit
    dlgCommon.Color = Picture1.BackColor
    dlgCommon.ShowColor
    Picture1.BackColor = dlgCommon.Color
    Command1_Click
    Command1.SetFocus
End Sub

Private Sub Option1_Click()
Label7.Enabled = True
Label8.Enabled = True
Command1.SetFocus
End Sub

Private Sub Option2_GotFocus()
        On Error Resume Next
        Option2.Value = True
        dlgCommon.Flags = cdlCCRGBInit
        dlgCommon.Color = col
        dlgCommon.ShowColor
        dlgCommon.DialogTitle = "Choose the color from which to derive the others"
        col = dlgCommon.Color
        
        Fade = InputBox("Insert a number between 1 - 100", "Choose Fade Value", "10")
        Command1_Click
Label7.Enabled = False
Label8.Enabled = False
Command1.SetFocus
End Sub

Private Sub Option3_Click()
Label7.Enabled = False
Label8.Enabled = False
Command1_Click
Command1.SetFocus
End Sub

Private Sub save_Click()
    With dlgCommon
        .DialogTitle = "Choose a filename to save"
        .Filter = "Bitmap files (*.BMP)|*.BMP"
        .FilterIndex = 1
        .FileName = ""
        .ShowSave
        
        If .FileName = "" Then Exit Sub
        
        SavePicture Picture1.Image, .FileName
    End With

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1_Click
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Command1_Click
End Sub

Private Sub VScroll1_Change()
Text1.Text = VScroll1.Value
Text1.SetFocus
End Sub

Private Sub VScroll2_Change()
Text2.Text = VScroll2.Value
Text2.SetFocus
End Sub

Private Sub Draw_Lines()
    On Error Resume Next
    For i = 0 To z Step st
        If Option3.Value = True Then col = Rnd * RGB(255, 255, 255) 'Checks the
        If Option2.Value = True Then col = col - Fade               'chosen colors
        
        'This was a pretty complicated thing to find, I mean I knew what to do but
        'it wasn't easy to made it right. If you want to figure it out you're welcome
        If Check1.Value = 1 Then Picture1.Line (0, i)-(z - i, 0), col
        If Check2.Value = 1 Then Picture1.Line (z - i, 0)-(z, z - i), col
        If Check3.Value = 1 Then Picture1.Line (z, z - i)-(i, z), col
        If Check4.Value = 1 Then Picture1.Line (0, i)-(i, z), col
    Next i
    
End Sub

Private Sub Status_Bar()
Dim Lines As Long
Lines = (z / st) * 4
StatusBar1.Panels.Item(1).Text = "Scale: " & z & "; Step: " & st
StatusBar1.Panels.Item(2).Text = "Lines Drawn: " & Lines
End Sub
