VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Help"
   ClientHeight    =   4260
   ClientLeft      =   3720
   ClientTop       =   2340
   ClientWidth     =   5355
   Icon            =   "PlannerHelp.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4260
   ScaleWidth      =   5355
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFC0C0&
      Height          =   3585
      Left            =   105
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "PlannerHelp.frx":030A
      Top             =   105
      Width           =   5160
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   435
      Left            =   1995
      TabIndex        =   0
      Top             =   3780
      Width           =   1380
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form2.Visible = False
    Form1.Visible = True
End Sub
