VERSION 5.00
Begin VB.Form OrderList 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FURNITURE ORDERING SYSTEM"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   14415
   Begin VB.CommandButton cmdOrder 
      Caption         =   "Order"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      TabIndex        =   4
      Top             =   7200
      Width           =   2655
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7440
      TabIndex        =   3
      Top             =   7200
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   720
         Width           =   13935
      End
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   12000
      X2              =   12000
      Y1              =   2160
      Y2              =   6960
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   10320
      X2              =   10320
      Y1              =   2160
      Y2              =   6960
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   9120
      X2              =   9120
      Y1              =   2160
      Y2              =   6960
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2880
      X2              =   2880
      Y1              =   2160
      Y2              =   6960
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1680
      X2              =   1680
      Y1              =   2160
      Y2              =   6960
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   120
      X2              =   14280
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "LIST OF ORDER"
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
      Left            =   480
      TabIndex        =   6
      Top             =   1680
      Width           =   13215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000040C0&
      Height          =   5535
      Left            =   120
      Top             =   1560
      Width           =   14175
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
      Left            =   480
      TabIndex        =   5
      Top             =   1680
      Width           =   13215
   End
End
Attribute VB_Name = "OrderList"
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

Private Sub cmdOrder_Click()
OpenKYAKINDatabase
OpenOrderTable
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
Print Tab(10); "Order ID "; Tab(25); "Product ID"; Tab(40); "Product Name"; Tab(125); "Price"; Tab(145); "Username"; Tab(165); "Order Date"; Tab(185);
Print Tab(3); "============================================================================================================================================================";
If (rs3.RecordCount > 0) Then
    rs3.MoveFirst
    Do While rs3.EOF = False
        Print Tab(10); (rs3.Fields("ORDER_ID").Value); Tab(25); (rs3.Fields("PRODUCT_ID").Value); Tab(40); (rs3.Fields("PRODUCT_NAME").Value); Tab(125); "RM "; (rs3.Fields("PRODUCT_PRICE").Value); Tab(145); (rs3.Fields("USERNAME").Value); Tab(165); (rs3.Fields("ORDER_DATE").Value); Tab(185);
        rs3.MoveNext
    Loop
Else
    MsgBox "Table is Empty", 0, "Warning"
End If
rs3.Close
Set rs3 = Nothing
End Sub

