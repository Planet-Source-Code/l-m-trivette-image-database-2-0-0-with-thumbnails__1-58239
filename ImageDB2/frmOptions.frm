VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Options"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5805
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   285
      Left            =   5040
      TabIndex        =   18
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   285
      Left            =   5040
      TabIndex        =   17
      Top             =   480
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Caption         =   "Options"
      Height          =   1695
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      Width           =   2775
      Begin VB.CheckBox Check1 
         Caption         =   "Delete Images after exporting"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   2415
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Delete Images after importing"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   2415
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Show multiple previews"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Thumbnail Size"
      Height          =   1695
      Left            =   3120
      TabIndex        =   5
      Top             =   1680
      Width           =   2415
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         Caption         =   "Symmetric"
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   1320
         Width           =   1095
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   285
         LargeChange     =   10
         Left            =   1800
         Max             =   500
         Min             =   20
         TabIndex        =   11
         Top             =   840
         Value           =   20
         Width           =   185
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   285
         LargeChange     =   10
         Left            =   1800
         Max             =   500
         Min             =   20
         TabIndex        =   8
         Top             =   480
         Value           =   20
         Width           =   185
      End
      Begin VB.Label Label4 
         Caption         =   "Height:"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Width:"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label Label2 
      Caption         =   "Default path to export to"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Image Database"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check3_Click()
    If Check3.Value = 1 Then Text4.Text = Text3.Text
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    Dim strFolder As String
    strFolder = GetFolder("Default Export Path")
    If strFolder <> "" Then
        ' Add the backslash if needed
        If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"
        Text2.Text = strFolder
    End If
End Sub

Private Sub Form_Load()
    ' Set the form controls based on the global variables
    Text1.Text = srcDB
    Text2.Text = ExportPath
    Text3.Text = ThumbWidth
    Text4.Text = ThumbHeight
    VScroll1.Value = ThumbWidth
    VScroll2.Value = ThumbHeight
    
    If DelExport = True Then Check1.Value = 1
    If DelImport = True Then Check2.Value = 1
    If MultiPreview = True Then Check4.Value = 1
    
    ' Load this from directly form the config.ini since it is not a global variable
    Check3.Value = GetValue(App.path & "\Config.ini", "Settings", "Symmetric", 0)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Save the settings to config.ini file
    SetValue App.path & "\Config.ini", "Settings", "Database", Text1.Text
    SetValue App.path & "\Config.ini", "Settings", "ExportPath", Text2.Text
    SetValue App.path & "\Config.ini", "Settings", "DelExport", CBool(Check1.Value)
    SetValue App.path & "\Config.ini", "Settings", "DelImport", CBool(Check2.Value)
    SetValue App.path & "\Config.ini", "Settings", "MultiPreview", CBool(Check4.Value)
    SetValue App.path & "\Config.ini", "Settings", "Symmetric", Check3.Value
    SetValue App.path & "\Config.ini", "Settings", "ThumbWidth", Text3.Text
    SetValue App.path & "\Config.ini", "Settings", "ThumbHeight", Text4.Text
    
    ' Reset the program global variables
    srcDB = Text1.Text
    ExportPath = Text2.Text
    ThumbWidth = Text3.Text
    ThumbHeight = Text4.Text
    DelExport = CBool(Check1.Value)
    DelImport = CBool(Check2.Value)
    MultiPreview = CBool(Check4.Value)
    
End Sub

Private Sub Text3_Change()
    If IsNumeric(Text3.Text) = True And Check3.Value = 1 Then Text4.Text = Text3.Text
End Sub

Private Sub Text3_KeyPress(keyascii As Integer)
    If keyascii <> 9 And keyascii <> 32 And keyascii <> 8 Then
        If keyascii < 48 Or keyascii > 57 Then keyascii = 0
    End If
End Sub

Private Sub Text3_LostFocus()
    ' Do error checking on thumbnail width
    If Text3.Text = "" Then Text3.Text = ThumbWidth
    If Val(Text3.Text) > 500 Then
        Text3.Text = "500"
        VScroll1.Value = 500
    End If
    VScroll1.Value = Val(Text3.Text)
    If Check3.Value = 1 Then VScroll2.Value = Val(Text3.Text)
End Sub

Private Sub Text4_Change()
    If IsNumeric(Text4.Text) = True And Check3.Value = 1 Then Text3.Text = Text4.Text
End Sub

Private Sub Text4_KeyPress(keyascii As Integer)
    If keyascii <> 9 And keyascii <> 32 And keyascii <> 8 Then
        If keyascii < 48 Or keyascii > 57 Then keyascii = 0
    End If
End Sub

Private Sub Text4_LostFocus()
    ' Do error checking on thumbnail width
    If Text4.Text = "" Then Text4.Text = ThumbHeight
    If Val(Text4.Text) > 500 Then
        Text4.Text = "500"
        VScroll2.Value = 500
    End If
    VScroll2.Value = Val(Text4.Text)
    If Check3.Value = 1 Then VScroll1.Value = Val(Text4.Text)
End Sub

Private Sub VScroll1_Change()
    Text3.Text = VScroll1.Value
    If Check3.Value = 1 Then Text4.Text = VScroll1.Value
End Sub

Private Sub VScroll2_Change()
    Text4.Text = VScroll2.Value
    If Check3.Value = 1 Then Text3.Text = VScroll2.Value
End Sub
