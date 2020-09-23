VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Formular1"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   2880
   StartUpPosition =   3  'Windows-Standard
   Begin Project1.LoSlider LoSlider1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2655
      _extentx        =   4683
      _extenty        =   661
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LoSlider1_Change()
Text1.Text = LoSlider1.Value
End Sub



Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
LoSlider1.Value = Text1.Text
End Sub
