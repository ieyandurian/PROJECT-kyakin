VERSION 5.00
Begin VB.Form Product 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FURNITURE ORDERING SYSTEM"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   720
   ClientWidth     =   14415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   14415
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Height          =   2775
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   14175
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Contact : 06-6794563"
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   495
         Left            =   0
         TabIndex        =   18
         Top             =   2280
         Width           =   13935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No 1, 46, 47, 48, 50, Jalan Tampin, Taman Bukit Emas, 70450 Seremban, Negeri Sembilan."
         BeginProperty Font 
            Name            =   "Book Antiqua"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   975
         Left            =   1680
         TabIndex        =   17
         Top             =   1320
         Width           =   11055
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "FURNITURE ORDERING SYSTEM"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   27.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   855
         Left            =   240
         TabIndex        =   16
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
            Size            =   36
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   975
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   13935
      End
   End
   Begin VB.TextBox txtProductDescription 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   12
      Top             =   5520
      Width           =   2415
   End
   Begin VB.TextBox txtProductName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   10
      Top             =   4800
      Width           =   2415
   End
   Begin VB.TextBox txtProductID 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   8
      Top             =   4080
      Width           =   2415
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
      Left            =   13080
      TabIndex        =   6
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   5
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   7200
      Width           =   1695
   End
   Begin VB.TextBox txtProductPrice 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   1
      Top             =   6240
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H000040C0&
      Caption         =   "RM"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   19
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label lblProducPrice 
      BackColor       =   &H000040C0&
      Caption         =   "Product Price :"
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
      Left            =   4920
      TabIndex        =   13
      Top             =   6240
      Width           =   2775
   End
   Begin VB.Label lblProductDescription 
      BackColor       =   &H000040C0&
      Caption         =   "Product Description :"
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
      Left            =   4920
      TabIndex        =   11
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Label lblProductName 
      BackColor       =   &H000040C0&
      Caption         =   "Product Name :"
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
      Left            =   4920
      TabIndex        =   9
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label lblProductId 
      BackColor       =   &H000040C0&
      Caption         =   "Product ID :"
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
      Left            =   4920
      TabIndex        =   7
      Top             =   4200
      Width           =   2775
   End
   Begin VB.Label lblProduct 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "P R O D U C T  F O R M"
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
      Left            =   480
      TabIndex        =   0
      Top             =   3240
      Width           =   13455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000040C0&
      Height          =   4095
      Left            =   120
      Top             =   3000
      Width           =   14175
   End
End
Attribute VB_Name = "Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    OpenKYAKINDatabase ' call procedure openregdatabase to open the Registration database
    Set rs1 = dbkyakin.OpenRecordset("PRODUCT") ' open product table in the Registration database file
        If Not txtProductID = "" Then
        rs1.AddNew 'to add
        rs1!PRODUCT_ID = txtProductID
        rs1!PRODUCT_NAME = txtProductName
        rs1!PRODUCT_DESCRIPTION = txtProductDescription
        rs1!PRODUCT_PRICE = Val(txtProductPrice)
        rs1.Update 'to add
        MsgBox "One record has been added ", vbOKOnly, "Add Record"
        Else
            MsgBox "The details is empty. Please make sure that you have inserted the data.", vbOKOnly, "No record"
        End If
        rs1.Close
        Set rs1 = Nothing
End Sub

Private Sub cmdClear_Click()
txtProductID = ""
txtProductName = ""
txtProductDescription = ""
txtProductPrice = ""
txtProductID.SetFocus
End Sub

Private Sub cmdDelete_Click()

Dim intProductID As String
'Form Load Event
'Set db=OpenDatabase(<<Your MsAccess Filename with path in Double Quotes>>)
'Example
'-------
 'Set db = OpenDatabase("d:\vb6\staff.mdb")
'Set rs= db.OpenRecordset(<<Your MsAccess Table Name in double quotes>>)
'Example
'-------


OpenKYAKINDatabase 'open the OpenEmpDatabase function in the module
OpenProductTable

intProductID = txtProductID
rs1.MoveFirst

'nostaff adalah nama index key dalam staff table
'to create index field , open table , click icon index , namakan filed index
'selalunya field Primary Key
        

rs1.Index = "product_id"
rs1.Seek "=", intProductID
If rs1.NoMatch Then
   MsgBox "sorry no record found", vbOKOnly, "sorry"
Else
   strans = MsgBox("Are Sure You Want To Delete This Record" & vbCrLf & "Product ID : " & (rs1.Fields("product_id").Value) & "  Product Name : " & (rs1.Fields("product_name").Value) & "  Product Description : " & (rs1.Fields("product_description").Value) & "  Product Price : " & (rs1.Fields("product_price").Value), vbYesNo, "Comfirmation")
   If strans = vbYes Then
    rs1.Delete
    MsgBox "One Record Been Deleted", 16, "Record Delete"
   End If
   
End If
rs1.Close
Set rs1 = Nothing

End Sub

Private Sub cmdSearch_Click()
Dim intStdno As String

OpenKYAKINDatabase 'open the OpenEmpDatabase function in the module
OpenProductTable

intProductID = txtProductID
If Not txtProductID = "" Then
    rs1.Index = "product_id"
    rs1.Seek "=", intProductID
    If rs1.NoMatch Then
        MsgBox "Sorry no record found", vbOKOnly, "sorry"
        txtProductID.SetFocus
    Else
        txtProductName = rs1!PRODUCT_NAME
        txtProductDescription = rs1!PRODUCT_DESCRIPTION
        txtProductPrice = rs1!PRODUCT_PRICE
    End If
Else
    MsgBox "The Product ID is empty. Please make sure that you have inserted the Product ID.", vbOKOnly, "Error"
End If

End Sub

Private Sub cmdUpdate_Click()
Dim strans As String

OpenKYAKINDatabase
OpenProductTable 'open student table in the registration database file

intProductID = txtProductID
If Not txtProductID = "" Then
rs1.Index = "product_id"
rs1.Seek "=", intProductID
rs1.Edit
rs1!PRODUCT_NAME = txtProductName
rs1!PRODUCT_DESCRIPTION = txtProductDescription
rs1!PRODUCT_PRICE = Val(txtProductPrice)
rs1.Update 'to add
MsgBox "The record has been updated ", vbOKOnly, "Add Record"
cmdClear_Click
Else
    MsgBox "The details is empty. Please make sure that you have inserted the data.", vbOKOnly, "No record"
End If
rs1.Close
Set rs1 = Nothing
End Sub

Private Sub cmdExit_Click()
    Call HomeAdmin.Show
    Unload Me
    MDIForm1!mnuFileCustomer.Enabled = False
    MDIForm1!mnuFileAdmin.Enabled = False
End Sub

Private Sub cmdHome_Click()
Call HomeAdmin.Show

End Sub

