VERSION 5.00
Begin VB.Form Settings1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Settings1.frx":0000
   ScaleHeight     =   3330
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Image title 
      Height          =   255
      Left            =   120
      Picture         =   "Settings1.frx":25F4A
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Image cancel 
      Height          =   375
      Left            =   1920
      Picture         =   "Settings1.frx":29F90
      Stretch         =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image apply 
      Height          =   375
      Left            =   120
      Picture         =   "Settings1.frx":2BC56
      Stretch         =   -1  'True
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "Settings1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
apply.Visible = False
cancel.Visible = False
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
title.Visible = True
FormDrag Settings1
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
title.Visible = False
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
title.Visible = False
End Sub


Private Sub Label2_Click()
Unload Me
Main.Show
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
apply.Visible = True
End Sub

Private Sub Label3_Click()
Unload Me
Main.Show
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cancel.Visible = True
End Sub
