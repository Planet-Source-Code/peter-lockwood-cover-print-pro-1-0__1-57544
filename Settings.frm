VERSION 5.00
Begin VB.Form Settings 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Settings.frx":0000
   ScaleHeight     =   3330
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Picture         =   "Settings.frx":25F4A
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   1920
      Picture         =   "Settings.frx":29F90
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   120
      Picture         =   "Settings.frx":2BC56
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1455
   End
End
Attribute VB_Name = "settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
title.Visible = True
FormDrag settings
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
title.Visible = False
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
title.Visible = False
End Sub


