VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5055
   ClientLeft      =   6345
   ClientTop       =   2820
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6840
      Top             =   3480
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   7035
      TabIndex        =   2
      Top             =   4560
      Width           =   7095
      Begin VB.Image Image1 
         Height          =   180
         Left            =   0
         Picture         =   "frmSplash.frx":000C
         Top             =   0
         Width           =   405
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "KEMUDI YAKIN FURNITURE STORE"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   1455
         Left            =   2640
         TabIndex        =   7
         Top             =   1200
         Width           =   3495
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "FURNITURE ORDERING SYSTEM"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Width           =   6495
      End
      Begin VB.Image imgLogo 
         Height          =   1395
         Left            =   1080
         Picture         =   "frmSplash.frx":034D
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label lblWarning 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Warning : Copyright Reserved 2014"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   3660
         Width           =   6855
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Loading Files......."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   5160
      TabIndex        =   4
      Top             =   4200
      Width           =   1935
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim X As Integer
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Load MDIForm1
    MDIForm1.Show
    Unload Me
End Sub

Private Sub Form_Load()
    'lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
   
    File1.FileName = App.Path
    X = File1.ListCount
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
If (Image1.Left <= 6600) Then
    Image1.Left = Image1.Left + 50
Else
    Image1.Left = 0
End If
If (i <= X) Then
    Label1.Caption = File1.List(i)
    i = i + 1
Else
Load MDIForm1
    MDIForm1.Show
    Unload Me
End If
End Sub
