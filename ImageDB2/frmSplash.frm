VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1845
      Left            =   720
      Picture         =   "frmSplash.frx":0000
      ScaleHeight     =   1845
      ScaleWidth      =   2250
      TabIndex        =   5
      Top             =   600
      Width           =   2250
   End
   Begin VB.Timer Timer1 
      Interval        =   6500
      Left            =   120
      Top             =   120
   End
   Begin VB.Line Line1 
      X1              =   3720
      X2              =   6720
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Image Database"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   615
      Left            =   3480
      TabIndex        =   4
      Top             =   960
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      Height          =   4215
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   7215
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Windows 9x/NT/2000/XP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Label lblWarning 
      BackColor       =   &H80000002&
      Caption         =   $"frmSplash.frx":207C
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   240
      TabIndex        =   1
      Top             =   3240
      Width           =   6735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2004,  Trivette Productions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4320
      TabIndex        =   0
      Top             =   95
      Width           =   2775
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H80000002&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   0
      Top             =   3120
      Width           =   7215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000006&
      FillColor       =   &H80000002&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   3480
      Top             =   0
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   1845
      Left            =   -2400
      Picture         =   "frmSplash.frx":215A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2250
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ' Get and update the current version information
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Pretty much any of these events below will close the splash screen   '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(keyascii As Integer)
    Unload Me
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub imgLogo_Click()
    Unload Me
End Sub

Private Sub lblCompany_Click()
    Unload Me
End Sub

Private Sub lblCopyright_Click()
    Unload Me
End Sub

Private Sub lblLicenseTo_Click()
    Unload Me
End Sub

Private Sub lblPlatform_Click()
    Unload Me
End Sub

Private Sub lblProductName_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Image1_Click()
    Unload Me
End Sub

Private Sub Label1_Click()
    Unload Me
End Sub

Private Sub Label2_Click()
    Unload Me
End Sub

Private Sub Label3_Click()
    Unload Me
End Sub

Private Sub Label4_Click()
    Unload Me
End Sub

Private Sub lblVersion_Click()
    Unload Me
End Sub

Private Sub lblWarning_Click()
    Unload Me
End Sub

Private Sub Picture1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub
