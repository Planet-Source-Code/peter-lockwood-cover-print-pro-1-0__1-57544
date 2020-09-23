VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Main 
   BorderStyle     =   0  'None
   Caption         =   "CoverPrintPro 1.0"
   ClientHeight    =   3330
   ClientLeft      =   3225
   ClientTop       =   2835
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Main.frx":0000
   ScaleHeight     =   3330
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdback 
      Left            =   2760
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdfront 
      Left            =   240
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4815
   End
   Begin VB.Image tb 
      Height          =   255
      Left            =   120
      Picture         =   "Main.frx":3EABA
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5400
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.Image xb 
      Height          =   255
      Left            =   5400
      Picture         =   "Main.frx":43320
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image us 
      Height          =   255
      Left            =   5040
      Picture         =   "Main.frx":437D6
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image pf 
      Height          =   375
      Left            =   120
      Picture         =   "Main.frx":43C8C
      Stretch         =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image pb 
      Height          =   375
      Left            =   4200
      Picture         =   "Main.frx":454A2
      Stretch         =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image bo 
      Height          =   375
      Left            =   2640
      Picture         =   "Main.frx":46F20
      Stretch         =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image front 
      Height          =   2295
      Left            =   120
      Picture         =   "Main.frx":4899E
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2295
   End
   Begin VB.Image back 
      Height          =   2295
      Left            =   2640
      Picture         =   "Main.frx":5AB58
      Stretch         =   -1  'True
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   2880
      Width           =   975
   End
   Begin VB.Image se 
      Height          =   375
      Left            =   1440
      Picture         =   "Main.frx":72562
      Stretch         =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub back_Click()
    cdback.DialogTitle = "Open Image (Back Cover)"
    cdback.Filter = "Image Files (*.bmp;*.jpg;*.gif;*.bif)|*.bmp;*.jpg;*.gif;*.bif"
    cdback.ShowOpen
    Set back.Picture = LoadPicture(cdback.FileName)
End Sub

Private Sub Command1_Click()
cd1.ShowPrinter
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
us.Visible = False
xb.Visible = False
pf.Visible = False
bo.Visible = False
pb.Visible = False
se.Visible = False
tb.Visible = False
End Sub

Private Sub front_Click()
    cdfront.DialogTitle = "Open Image (Front Cover)"
    cdfront.Filter = "Image Files (*.bmp;*.jpg;*.gif;*.bif)|*.bmp;*.jpg;*.gif;*.bif"
    cdfront.ShowOpen
    Set front.Picture = LoadPicture(cdfront.FileName)
End Sub

Private Sub Label1_Click()
Main.WindowState = 1
us.Visible = False
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
us.Visible = True
xb.Visible = False
tb.Visible = False
End Sub

Private Sub Label2_Click()
Unload Main
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
xb.Visible = True
us.Visible = False
End Sub

Private Sub Label3_Click()
Call PrintFront
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
pf.Visible = True
se.Visible = False
End Sub

Private Sub Label4_Click()
Main.Visible = False
Settings1.Visible = True
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
se.Visible = True
pf.Visible = False
End Sub

Private Sub Label5_Click()
Call PrintBoth
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
bo.Visible = True
pb.Visible = False
End Sub

Private Sub Label6_Click()
Call PrintBack
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
pb.Visible = True
bo.Visible = False
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
tb.Visible = True
Label7.Caption = " Cover Print Pro 1.0 Beta"
FormDrag Main
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tb.Visible = False
Label7.Caption = ""
End Sub

Private Sub Label7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
tb.Visible = False
End Sub
