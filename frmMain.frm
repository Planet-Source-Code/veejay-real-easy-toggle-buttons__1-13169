VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The easiest toggle buttons on Planet Source Code!!!"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit (This is a normal button)"
      Height          =   975
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Frame2"
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Frame1"
      Height          =   975
      Left            =   120
      Picture         =   "frmMain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3615
      Left            =   3840
      TabIndex        =   6
      Top             =   0
      Width           =   2415
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   240
         Picture         =   "frmMain.frx":0442
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   7
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label5 
         Caption         =   "THIS WAS DISCOVERED BY ACCIDENT WHEN I WANTED TO SEE WHAT THE PICTURE PROPERTY OF A CHECKBOX DOES"
         Height          =   1215
         Left            =   120
         TabIndex        =   13
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   $"frmMain.frx":0884
         Height          =   2295
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.CheckBox Check3 
         Caption         =   "Enable"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   2880
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         DisabledPicture =   "frmMain.frx":0926
         DownPicture     =   "frmMain.frx":0D68
         Height          =   975
         Left            =   240
         Picture         =   "frmMain.frx":1072
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2520
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Height          =   975
         Left            =   240
         Picture         =   "frmMain.frx":137C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "With Text"
         Height          =   975
         Left            =   240
         Picture         =   "frmMain.frx":1686
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "These are Radio Buttons"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Without Text"
         Height          =   495
         Left            =   1440
         TabIndex        =   8
         Top             =   1680
         Width           =   855
      End
   End
   Begin VB.Label Label4 
      Caption         =   "These are Check Boxes"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    Frame1.Visible = CBool(Check1.Value)
End Sub

Private Sub Check2_Click()
    Frame2.Visible = CBool(Check2.Value)
End Sub

Private Sub Check3_Click()
    Option3.Enabled = CBool(Check3.Value)
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub
