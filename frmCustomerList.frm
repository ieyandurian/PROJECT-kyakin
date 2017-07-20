VERSION 5.00
Begin VB.Form frmCustomerList 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FURNITURE ORDERING SYSTEM"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   720
   ClientWidth     =   14655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   14655
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Height          =   1455
      Left            =   240
      TabIndex        =   3
      Top             =   120
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   120
         Width           =   13695
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      TabIndex        =   1
      Top             =   7560
      Width           =   2655
   End
   Begin VB.CommandButton cmdCustomer 
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   0
      Top             =   7560
      Width           =   2655
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   14400
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   5040
      X2              =   5040
      Y1              =   2280
      Y2              =   7080
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   6600
      X2              =   6600
      Y1              =   2280
      Y2              =   7080
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   9600
      X2              =   9600
      Y1              =   2280
      Y2              =   7080
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   8400
      X2              =   8400
      Y1              =   2280
      Y2              =   7080
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   3960
      X2              =   3960
      Y1              =   2280
      Y2              =   7080
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2160
      X2              =   2160
      Y1              =   2280
      Y2              =   7080
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "LIST OF  CUSTOMER"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1800
      Width           =   13215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000040C0&
      Height          =   5535
      Left            =   120
      Top             =   1680
      Width           =   14415
   End
End
Attribute VB_Name = "frmCustomerList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
    Call HomeAdmin.Show
    Unload Me
    MDIForm1!mnuFileCustomer.Enabled = False
    MDIForm1!mnuFileAdmin.Enabled = False
End Sub

Private Sub cmdCustomer_Click()
OpenKYAKINDatabase
OpenUserTable
Print ""
Print ""
Print ""
Print ""
Print ""
Print ""
Print ""
Print ""
Print ""
Print ""
Print ""
Print ""
Print Tab(5); "FirstName"; Tab(33); "LastName"; Tab(55); "Phone No."; Tab(70); "City"; Tab(90); "State"; Tab(115); "Postcode"; Tab(130); "Street";
Print Tab(4); "==============================================================================================================================================================";
If (rs.RecordCount > 0) Then
    rs.MoveFirst
    Do While rs.EOF = False
        Print Tab(5); (rs.Fields("USER_FNAME").Value); Tab(33); (rs.Fields("USER_LNAME").Value); Tab(55); (rs.Fields("USER_PHONE").Value); Tab(70); (rs.Fields("USER_CITY").Value); Tab(90); (rs.Fields("USER_STATE").Value); Tab(115); (rs.Fields("USER_POSTCODE").Value); Tab(130); _
        (rs.Fields("USER_STREETNAME").Value);
        
        rs.MoveNext
        Loop
    Else
        MsgBox "Table is Empty", 0, "Warning"
End If
rs.Close
Set rs = Nothing
        
End Sub

