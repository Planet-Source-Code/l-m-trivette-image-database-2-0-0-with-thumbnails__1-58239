VERSION 5.00
Begin VB.Form frmAutoRename 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Auto Renaming Utility"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Rename to Match Pattern"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4335
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Text            =   "0"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Beginning Number"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Base Name (if any)"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "0001.jpg"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "Preview:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   735
   End
End
Attribute VB_Name = "frmAutoRename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    MsgBox "Sorry, this utility has not been written yet"
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub




