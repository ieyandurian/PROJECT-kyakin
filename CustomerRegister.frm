VERSION 5.00
Begin VB.Form CustomerRegister 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FURNITURE ORDERING SYSTEM"
   ClientHeight    =   8490
   ClientLeft      =   3345
   ClientTop       =   1965
   ClientWidth     =   14295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   14295
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Height          =   1455
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   14175
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
         TabIndex        =   14
         Top             =   120
         Width           =   13695
      End
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
         TabIndex        =   13
         Top             =   720
         Width           =   13935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      Height          =   5775
      Left            =   960
      TabIndex        =   8
      Top             =   2520
      Width           =   11295
      Begin VB.TextBox txtUsername 
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
         TabIndex        =   26
         Top             =   345
         Width           =   2775
      End
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
         TabIndex        =   25
         Top             =   1470
         Width           =   2775
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   7680
         TabIndex        =   24
         Top             =   315
         Width           =   2775
      End
      Begin VB.TextBox txtContact 
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
         TabIndex        =   23
         Top             =   3195
         Width           =   2775
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
         Height          =   405
         Left            =   2520
         TabIndex        =   22
         Top             =   4065
         Width           =   2775
      End
      Begin VB.TextBox txtLastName 
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2520
         TabIndex        =   21
         Top             =   2310
         Width           =   2775
      End
      Begin VB.TextBox txtPostCode 
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   7680
         TabIndex        =   3
         Top             =   4035
         Width           =   2775
      End
      Begin VB.TextBox txtState 
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   7680
         TabIndex        =   2
         Top             =   3165
         Width           =   2775
      End
      Begin VB.TextBox txtCity 
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   7680
         TabIndex        =   1
         Top             =   2310
         Width           =   2775
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   5
         Top             =   5040
         Width           =   1695
      End
      Begin VB.CommandButton cmdReg 
         Caption         =   "REGISTER"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5880
         TabIndex        =   4
         Top             =   5040
         Width           =   1695
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7680
         TabIndex        =   0
         Top             =   1365
         Width           =   2775
      End
      Begin VB.Label lblUsername 
         BackColor       =   &H00000000&
         Caption         =   "Username :"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   720
         TabIndex        =   27
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblPass 
         BackColor       =   &H00000000&
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5880
         TabIndex        =   11
         Top             =   360
         Width           =   2175
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   1095
         Left            =   0
         Top             =   0
         Width           =   11295
      End
      Begin VB.Label lblPostcode 
         BackColor       =   &H008080FF&
         Caption         =   "PostCode :"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   20
         Top             =   4080
         Width           =   2175
      End
      Begin VB.Label lblState 
         BackColor       =   &H008080FF&
         Caption         =   "State :"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   19
         Top             =   3210
         Width           =   2175
      End
      Begin VB.Label lblCity 
         BackColor       =   &H008080FF&
         Caption         =   "City :"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   18
         Top             =   2355
         Width           =   2175
      End
      Begin VB.Label lblLastName 
         BackColor       =   &H008080FF&
         Caption         =   "Last Name :"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   17
         Top             =   2355
         Width           =   1455
      End
      Begin VB.Label lblEmail 
         BackColor       =   &H008080FF&
         Caption         =   "Email :"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   16
         Top             =   4080
         Width           =   2175
      End
      Begin VB.Label lblContact 
         BackColor       =   &H008080FF&
         Caption         =   "Contact :"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   15
         Top             =   3210
         Width           =   2175
      End
      Begin VB.Label lblAddress 
         BackColor       =   &H008080FF&
         Caption         =   "Address :"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   10
         Top             =   1485
         Width           =   2175
      End
      Begin VB.Label lblFullName 
         BackColor       =   &H008080FF&
         Caption         =   "First Name :"
         BeginProperty Font 
            Name            =   "Myriad Hebrew"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   1485
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command1 
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
      Left            =   12600
      TabIndex        =   6
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label lblRegister 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "R E G I S T R A T I O N  F O R M"
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
      Left            =   720
      TabIndex        =   7
      Top             =   1800
      Width           =   12735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000040C0&
      Height          =   7095
      Left            =   360
      Top             =   1680
      Width           =   13575
   End
End
Attribute VB_Name = "CustomerRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
    txtUsername = ""
    txtFirstName = ""
    txtLastName = ""
    txtContact = ""
    txtCity = ""
    txtState = ""
    txtPostCode = ""
    txtAddress = ""
    txtPassword = ""
    txtEmail = ""
    txtUsername.SetFocus
End Sub

Private Sub cmdHome_Click()
    Unload Me
    MDIForm1!mnuFileCustomer.Enabled = True
    MDIForm1!mnuFileCustomerLogin.Enabled = True
    MDIForm1!mnuFileAdmin.Enabled = True
End Sub
Private Sub cmdReg_Click()
OpenKYAKINDatabase
Dim struser As String
OpenUserTable
struser = txtUsername
If Not (txtUsername = "" And txtPassword = "") Then
        Set rs4 = dbkyakin.OpenRecordset("USER")
        rs.Index = "user_username"
        rs.Seek "=", struser
        If rs.NoMatch Then
        rs4.AddNew
        rs4!user_username = txtUsername
        rs4!USER_FNAME = txtFirstName
        rs4!USER_LNAME = txtLastName
        rs4!USER_PHONE = txtContact
        rs4!USER_CITY = txtCity
        rs4!USER_STATE = txtState
        rs4!USER_POSTCODE = txtPostCode
        rs4!USER_STREETNAME = txtAddress
        rs4!Password = txtPassword
        rs4!USER_EMAIL = txtEmail
        rs4!USER_ROLE = "CUSTOMER"
        rs4.Update
        MsgBox "You have been registered. Please go to the login page.", vbOKOnly, "Congratulations"
        Else
            MsgBox "User existed.", vbOKOnly, "Error"
        End If
        cmdClear_Click
        rs4.Close
        Set rs4 = Nothing
        rs.Close
        Set rs = Nothing
        Else
            MsgBox "You are required to enter username and password to register.", vbExclamation, "Error"
        End If
End Sub

Private Sub Command1_Click()
    Unload Me
    MDIForm1!mnuFileCustomer.Enabled = True
    MDIForm1!mnuFileCustomerLogin.Enabled = True
    MDIForm1!mnuFileAdmin.Enabled = True
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

