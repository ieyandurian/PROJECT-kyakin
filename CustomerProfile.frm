VERSION 5.00
Begin VB.Form CustomerProfile 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FURNITURE ORDERING SYSTEM"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   720
   ClientWidth     =   14175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   14175
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   3000
      TabIndex        =   5
      Top             =   2640
      Width           =   7935
      Begin VB.TextBox txtFirstName 
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2520
         TabIndex        =   14
         Top             =   960
         Width           =   3975
      End
      Begin VB.TextBox txtPhone 
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   13
         Top             =   4080
         Width           =   3975
      End
      Begin VB.TextBox txtEmail 
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   12
         Top             =   4800
         Width           =   3975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   11
         Top             =   5640
         Width           =   1695
      End
      Begin VB.TextBox txtLastName 
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2520
         TabIndex        =   10
         Top             =   1560
         Width           =   3975
      End
      Begin VB.TextBox txtStreet 
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   9
         Top             =   2160
         Width           =   3975
      End
      Begin VB.TextBox txtPostcode 
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox txtCity 
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   7
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtState 
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label lblFirstName 
         BackColor       =   &H008080FF&
         Caption         =   "First Name :"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   24
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lblStreet 
         BackColor       =   &H008080FF&
         Caption         =   "Street :"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   23
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label lblPhoneNo 
         BackColor       =   &H008080FF&
         Caption         =   "Phone Number :"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   22
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Label lblEmail 
         BackColor       =   &H008080FF&
         Caption         =   "Email :"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   21
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label lblLastName 
         BackColor       =   &H008080FF&
         Caption         =   "Last Name :"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   20
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblCustId 
         BackColor       =   &H008080FF&
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   19
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblUsername 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   18
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblPostcode 
         BackColor       =   &H008080FF&
         Caption         =   "Postcode :"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   17
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label lblCity 
         BackColor       =   &H008080FF&
         Caption         =   "City :"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   16
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label lblState 
         BackColor       =   &H008080FF&
         Caption         =   "State :"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   3480
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      TabIndex        =   4
      Top             =   8640
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14175
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "KEMUDI YAKIN FURNITURE STORE"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   27.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   13935
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "FURNITURE ORDERING SYSTEM"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   26.25
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   855
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   13695
      End
   End
   Begin VB.Label lblCustProfile 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "P R O F I L E  D E T A I L"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   1920
      Width           =   12255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000040C0&
      Height          =   7575
      Left            =   360
      Top             =   1680
      Width           =   13455
   End
End
Attribute VB_Name = "CustomerProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Call HomeCustomer.Show
    Unload Me
    MDIForm1!mnuFileCustomer.Enabled = False
    MDIForm1!mnuFileAdmin.Enabled = False
End Sub

Private Sub cmdHome_Click()
    Call HomeCustomer.Show
    Unload Me
    MDIForm1!mnuFileCustomer.Enabled = False
    MDIForm1!mnuFileAdmin.Enabled = False
End Sub
Private Sub Timer1_Timer()
Frame4.Top = Frame4.Top + 100
If Frame4.Top >= 4080 Then
Timer1.Enabled = False
Timer2.Enabled = True
End If
End Sub
Private Sub Timer2_Timer()
Frame3.Top = Frame3.Top - 100
If Frame3.Top <= -1560 Then
Timer1.Enabled = True
Timer2.Enabled = False
End If
End Sub

Private Sub cmdUpdate_Click()


OpenKYAKINDatabase
OpenUserTable 'open student table in the registration database file

    intStdno = lblUsername
    rs.Index = "user_username"
    rs.Seek "=", intStdno
    rs.Edit
    rs!USER_FNAME = txtFirstName
    rs!USER_LNAME = txtLastName
    rs!USER_STREETNAME = txtStreet
    rs!USER_POSTCODE = txtPostCode
    rs!USER_CITY = txtCity
    rs!USER_STATE = txtState
    rs!USER_PHONE = txtPhone
    rs.Update
    MsgBox "Record has been updated", vbOKOnly, "Add Record"
    rs.Close
    Set rs = Nothing
End Sub


Private Sub Form_Load()
    lblUsername.Caption = usern
    OpenKYAKINDatabase
OpenUserTable
    rs.Index = "user_username"
    rs.Seek "=", usern
    txtFirstName = rs!USER_FNAME
    txtLastName = rs!USER_LNAME
    txtStreet = rs!USER_STREETNAME
    txtPostCode = rs!USER_POSTCODE
    txtCity = rs!USER_CITY
    txtState = rs!USER_STATE
    txtPhone = rs!USER_PHONE
    txtEmail = rs!USER_EMAIL
End Sub

