VERSION 5.00
Begin VB.Form Customer 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FURNITURE ORDERING SYSTEM"
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   720
   ClientWidth     =   15510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   15510
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      Caption         =   "Product"
      Height          =   4815
      Left            =   120
      TabIndex        =   26
      Top             =   1680
      Width           =   15255
      Begin VB.OptionButton optProd1 
         BackColor       =   &H008080FF&
         Caption         =   "Amisco Bedroom Billboard Queen Platform Bed"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   2295
      End
      Begin VB.OptionButton optProd2 
         BackColor       =   &H008080FF&
         Caption         =   "Amisco Bedroom Bridge Queen Platform Bed"
         Height          =   615
         Left            =   2640
         TabIndex        =   7
         Top             =   1920
         Width           =   2295
      End
      Begin VB.OptionButton optProd3 
         BackColor       =   &H008080FF&
         Caption         =   "Amisco Bedroom Full Headboard"
         Height          =   375
         Left            =   5160
         TabIndex        =   8
         Top             =   2040
         Width           =   1935
      End
      Begin VB.OptionButton optProd4 
         BackColor       =   &H008080FF&
         Caption         =   "Amisco Youth Bedroom Twin Headboard"
         Height          =   615
         Left            =   7680
         TabIndex        =   9
         Top             =   1920
         Width           =   2055
      End
      Begin VB.OptionButton optProd5 
         BackColor       =   &H008080FF&
         Caption         =   "Leggett And Platt Bedroom Belmont Twin Headboard"
         Height          =   615
         Left            =   10200
         TabIndex        =   10
         Top             =   1920
         Width           =   1815
      End
      Begin VB.OptionButton optProd6 
         BackColor       =   &H008080FF&
         Caption         =   "Americana Dining Room Oak Side Chair"
         Height          =   375
         Left            =   12480
         TabIndex        =   11
         Top             =   2040
         Width           =   2055
      End
      Begin VB.OptionButton optProd7 
         BackColor       =   &H008080FF&
         Caption         =   "Bassett Dining Room Sideboard"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   4200
         Width           =   1815
      End
      Begin VB.OptionButton optProd8 
         BackColor       =   &H008080FF&
         Caption         =   "Signature Design Dining Room 5-Piece Counter Height Set"
         Height          =   615
         Left            =   2640
         TabIndex        =   13
         Top             =   4080
         Width           =   1935
      End
      Begin VB.OptionButton optProd9 
         BackColor       =   &H008080FF&
         Caption         =   "Signature Design Dining Room 6-Piece Dining Packaged"
         Height          =   375
         Left            =   5160
         TabIndex        =   14
         Top             =   4200
         Width           =   1695
      End
      Begin VB.OptionButton optProd10 
         BackColor       =   &H008080FF&
         Caption         =   "Nuevo Dining Room Versailles Dining Table"
         Height          =   375
         Left            =   7680
         TabIndex        =   15
         Top             =   4200
         Width           =   1935
      End
      Begin VB.OptionButton optProd11 
         BackColor       =   &H008080FF&
         Caption         =   "Signature Design Dining Room 6-Piece Dining Packaged"
         Height          =   375
         Left            =   10080
         TabIndex        =   16
         Top             =   4200
         Width           =   1695
      End
      Begin VB.OptionButton optProd12 
         BackColor       =   &H008080FF&
         Caption         =   "Corinthian Living Room Exclusive"
         Height          =   375
         Left            =   12480
         TabIndex        =   17
         Top             =   4200
         Width           =   1695
      End
      Begin VB.PictureBox Picture12 
         Height          =   1515
         Left            =   12600
         ScaleHeight     =   1515
         ScaleMode       =   0  'User
         ScaleWidth      =   1815
         TabIndex        =   38
         Top             =   2520
         Width           =   1875
         Begin VB.Image Image12 
            Height          =   1455
            Left            =   0
            Picture         =   "Customer.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1875
         End
      End
      Begin VB.PictureBox Picture11 
         Height          =   1515
         Left            =   10200
         ScaleHeight     =   1515
         ScaleMode       =   0  'User
         ScaleWidth      =   1875
         TabIndex        =   37
         Top             =   2520
         Width           =   1875
         Begin VB.Image Image11 
            Height          =   1455
            Left            =   0
            Picture         =   "Customer.frx":2D1AF
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.PictureBox Picture10 
         Height          =   1515
         Left            =   7800
         ScaleHeight     =   1515
         ScaleMode       =   0  'User
         ScaleWidth      =   1875
         TabIndex        =   36
         Top             =   2520
         Width           =   1875
         Begin VB.Image Image10 
            Height          =   1455
            Left            =   0
            Picture         =   "Customer.frx":5313E
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.PictureBox Picture9 
         Height          =   1515
         Left            =   5280
         ScaleHeight     =   1515
         ScaleMode       =   0  'User
         ScaleWidth      =   1875
         TabIndex        =   35
         Top             =   2520
         Width           =   1875
         Begin VB.Image Image9 
            Height          =   1455
            Left            =   0
            Picture         =   "Customer.frx":6D587
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.PictureBox Picture8 
         Height          =   1515
         Left            =   2760
         Picture         =   "Customer.frx":962F4
         ScaleHeight     =   1515
         ScaleMode       =   0  'User
         ScaleWidth      =   1875
         TabIndex        =   34
         Top             =   2520
         Width           =   1875
         Begin VB.Image Image8 
            Height          =   1455
            Left            =   0
            Picture         =   "Customer.frx":C2915
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.PictureBox Picture7 
         Height          =   1515
         Left            =   240
         ScaleHeight     =   1515
         ScaleMode       =   0  'User
         ScaleWidth      =   1875
         TabIndex        =   33
         Top             =   2520
         Width           =   1935
         Begin VB.Image Image7 
            Height          =   1455
            Left            =   0
            Picture         =   "Customer.frx":EEF36
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1875
         End
      End
      Begin VB.PictureBox Picture6 
         Height          =   1575
         Left            =   12600
         ScaleHeight     =   1515
         ScaleWidth      =   1815
         TabIndex        =   32
         Top             =   240
         Width           =   1875
         Begin VB.Image Image6 
            Height          =   1515
            Left            =   0
            Picture         =   "Customer.frx":1227AB
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1875
         End
      End
      Begin VB.PictureBox Picture5 
         Height          =   1575
         Left            =   10200
         ScaleHeight     =   1577.474
         ScaleMode       =   0  'User
         ScaleWidth      =   1875
         TabIndex        =   31
         Top             =   240
         Width           =   1815
         Begin VB.Image Image5 
            Height          =   1455
            Left            =   0
            Picture         =   "Customer.frx":14533F
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1755
         End
      End
      Begin VB.PictureBox Picture4 
         Height          =   1575
         Left            =   7800
         ScaleHeight     =   1515
         ScaleMode       =   0  'User
         ScaleWidth      =   1875
         TabIndex        =   30
         Top             =   240
         Width           =   1815
         Begin VB.Image Image4 
            Height          =   1515
            Left            =   0
            Picture         =   "Customer.frx":1587C9
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1755
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   1575
         Left            =   5280
         ScaleHeight     =   1515
         ScaleMode       =   0  'User
         ScaleWidth      =   1875
         TabIndex        =   29
         Top             =   240
         Width           =   1875
         Begin VB.Image Image3 
            Height          =   1515
            Left            =   0
            Picture         =   "Customer.frx":174651
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.PictureBox Picture2 
         Height          =   1575
         Left            =   2760
         ScaleHeight     =   1515
         ScaleWidth      =   1875
         TabIndex        =   28
         Top             =   240
         Width           =   1935
         Begin VB.Image Image2 
            Height          =   1515
            Left            =   0
            Picture         =   "Customer.frx":192C33
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1875
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   1575
         Left            =   240
         ScaleHeight     =   1515
         ScaleWidth      =   1875
         TabIndex        =   27
         Top             =   240
         Width           =   1935
         Begin VB.Image Image1 
            Height          =   1515
            Left            =   0
            Picture         =   "Customer.frx":1AD86D
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1875
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Height          =   1455
      Left            =   600
      TabIndex        =   23
      Top             =   120
      Width           =   14295
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
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   720
         Width           =   13935
      End
   End
   Begin VB.CommandButton cmdCart 
      Caption         =   "SUBMIT ORDER"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      TabIndex        =   4
      Top             =   7440
      Width           =   2175
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
      Left            =   10440
      TabIndex        =   5
      Top             =   8160
      Width           =   2175
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
      Left            =   14160
      TabIndex        =   19
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Label lblProdPrice 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   6120
      TabIndex        =   3
      Top             =   8760
      Width           =   2775
   End
   Begin VB.Label lblProdDesc 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   5640
      TabIndex        =   2
      Top             =   8160
      Width           =   3255
   End
   Begin VB.Label lblProdName 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   5640
      TabIndex        =   1
      Top             =   7320
      Width           =   3255
   End
   Begin VB.Label lblProdID 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   5640
      TabIndex        =   0
      Top             =   6720
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H000040C0&
      Caption         =   "RM"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   39
      Top             =   8760
      Width           =   615
   End
   Begin VB.Label lblProductId 
      BackColor       =   &H000040C0&
      Caption         =   "Product ID                    :"
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
      Left            =   2880
      TabIndex        =   22
      Top             =   6720
      Width           =   2775
   End
   Begin VB.Label lblProductName 
      BackColor       =   &H000040C0&
      Caption         =   "Product Name            :"
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
      Left            =   2880
      TabIndex        =   21
      Top             =   7320
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
      Left            =   2880
      TabIndex        =   20
      Top             =   8160
      Width           =   2775
   End
   Begin VB.Label lblProducPrice 
      BackColor       =   &H000040C0&
      Caption         =   "Product Price              :"
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
      Left            =   2880
      TabIndex        =   18
      Top             =   8760
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000040C0&
      Height          =   2655
      Left            =   2760
      Top             =   6600
      Width           =   9975
   End
End
Attribute VB_Name = "Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCart_Click()
OpenKYAKINDatabase
OpenOrderTable 'open student table in the registration database file
If Not (lblProdID = "") Then
    Set rs3 = dbkyakin.OpenRecordset("ORDER")
    rs3.AddNew
        rs3!PRODUCT_ID = lblProdID
        rs3!PRODUCT_NAME = lblProdName
        rs3!PRODUCT_PRICE = lblProdPrice
        rs3!UserName = usern
        rs3!Order_Date = Now()
        rs3.Update
        MsgBox "Order has been submited", vbOKOnly, "Add Record"
    cmdClear_Click
    rs3.Close
    Set rs3 = Nothing
    Else
        MsgBox "Please choose product to Order", vbExclamation, "Warning"
End If
    
End Sub

Private Sub cmdClear_Click()
lblProdID.Caption = ""
lblProdName.Caption = ""
lblProdDesc.Caption = ""
lblProdPrice.Caption = ""
optProd1.Value = False
optProd2.Value = False
optProd3.Value = False
optProd4.Value = False
optProd5.Value = False
optProd6.Value = False
optProd7.Value = False
optProd8.Value = False
optProd9.Value = False
optProd10.Value = False
optProd11.Value = False
optProd12.Value = False
End Sub

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

Private Sub optProd1_Click()
OpenKYAKINDatabase
OpenProductTable
If optProd1.Value = True Then
    lblProdID.Caption = "P0001"
    
    prodID = lblProdID
    rs1.Index = "PRODUCT_ID"
    rs1.Seek "=", prodID
    If rs1.NoMatch Then
        MsgBox "No record found", vbOKOnly, "Error"
    Else
        lblProdName = rs1!PRODUCT_NAME
        lblProdPrice = rs1!PRODUCT_PRICE
        lblProdDesc = rs1!PRODUCT_DESCRIPTION
    End If
End If
rs1.Close
Set rs1 = Nothing
End Sub

Private Sub optProd10_Click()
OpenKYAKINDatabase
OpenProductTable
If optProd10.Value = True Then
    lblProdID.Caption = "P0010"
    
    prodID = lblProdID
    rs1.Index = "PRODUCT_ID"
    rs1.Seek "=", prodID
    If rs1.NoMatch Then
        MsgBox "No record found", vbOKOnly, "Error"
    Else
        lblProdName = rs1!PRODUCT_NAME
        lblProdPrice = rs1!PRODUCT_PRICE
        lblProdDesc = rs1!PRODUCT_DESCRIPTION
    End If
End If
rs1.Close
Set rs1 = Nothing
End Sub

Private Sub optProd11_Click()
OpenKYAKINDatabase
OpenProductTable
If optProd11.Value = True Then
    lblProdID.Caption = "P0011"
    
    prodID = lblProdID
    rs1.Index = "PRODUCT_ID"
    rs1.Seek "=", prodID
    If rs1.NoMatch Then
        MsgBox "No record found", vbOKOnly, "Error"
    Else
        lblProdName = rs1!PRODUCT_NAME
        lblProdPrice = rs1!PRODUCT_PRICE
        lblProdDesc = rs1!PRODUCT_DESCRIPTION
    End If
End If
rs1.Close
Set rs1 = Nothing
End Sub

Private Sub optProd12_Click()
OpenKYAKINDatabase
OpenProductTable
If optProd12.Value = True Then
    lblProdID.Caption = "P0012"
    
    prodID = lblProdID
    rs1.Index = "PRODUCT_ID"
    rs1.Seek "=", prodID
    If rs1.NoMatch Then
        MsgBox "No record found", vbOKOnly, "Error"
    Else
        lblProdName = rs1!PRODUCT_NAME
        lblProdPrice = rs1!PRODUCT_PRICE
        lblProdDesc = rs1!PRODUCT_DESCRIPTION
    End If
End If
rs1.Close
Set rs1 = Nothing
End Sub

Private Sub optProd2_Click()
OpenKYAKINDatabase
OpenProductTable
If optProd2.Value = True Then
    lblProdID.Caption = "P0002"
    
    prodID = lblProdID
    rs1.Index = "PRODUCT_ID"
    rs1.Seek "=", prodID
    If rs1.NoMatch Then
        MsgBox "No record found", vbOKOnly, "Error"
    Else
        lblProdName = rs1!PRODUCT_NAME
        lblProdPrice = rs1!PRODUCT_PRICE
        lblProdDesc = rs1!PRODUCT_DESCRIPTION
    End If
End If
rs1.Close
Set rs1 = Nothing
End Sub

Private Sub optProd3_Click()
OpenKYAKINDatabase
OpenProductTable
If optProd3.Value = True Then
    lblProdID.Caption = "P0003"
    
    prodID = lblProdID
    rs1.Index = "PRODUCT_ID"
    rs1.Seek "=", prodID
    If rs1.NoMatch Then
        MsgBox "No record found", vbOKOnly, "Error"
    Else
        lblProdName = rs1!PRODUCT_NAME
        lblProdPrice = rs1!PRODUCT_PRICE
        lblProdDesc = rs1!PRODUCT_DESCRIPTION
    End If
End If
rs1.Close
Set rs1 = Nothing
End Sub

Private Sub optProd4_Click()
OpenKYAKINDatabase
OpenProductTable
If optProd4.Value = True Then
    lblProdID.Caption = "P0004"
    
    prodID = lblProdID
    rs1.Index = "PRODUCT_ID"
    rs1.Seek "=", prodID
    If rs1.NoMatch Then
        MsgBox "No record found", vbOKOnly, "Error"
    Else
        lblProdName = rs1!PRODUCT_NAME
        lblProdPrice = rs1!PRODUCT_PRICE
        lblProdDesc = rs1!PRODUCT_DESCRIPTION
    End If
End If
rs1.Close
Set rs1 = Nothing
End Sub

Private Sub optProd5_Click()
OpenKYAKINDatabase
OpenProductTable
If optProd5.Value = True Then
    lblProdID.Caption = "P0005"
    
    prodID = lblProdID
    rs1.Index = "PRODUCT_ID"
    rs1.Seek "=", prodID
    If rs1.NoMatch Then
        MsgBox "No record found", vbOKOnly, "Error"
    Else
        lblProdName = rs1!PRODUCT_NAME
        lblProdPrice = rs1!PRODUCT_PRICE
        lblProdDesc = rs1!PRODUCT_DESCRIPTION
    End If
End If
rs1.Close
Set rs1 = Nothing
End Sub

Private Sub optProd6_Click()
OpenKYAKINDatabase
OpenProductTable
If optProd6.Value = True Then
    lblProdID.Caption = "P0006"
    
    prodID = lblProdID
    rs1.Index = "PRODUCT_ID"
    rs1.Seek "=", prodID
    If rs1.NoMatch Then
        MsgBox "No record found", vbOKOnly, "Error"
    Else
        lblProdName = rs1!PRODUCT_NAME
        lblProdPrice = rs1!PRODUCT_PRICE
        lblProdDesc = rs1!PRODUCT_DESCRIPTION
    End If
End If
rs1.Close
Set rs1 = Nothing
End Sub

Private Sub optProd7_Click()
OpenKYAKINDatabase
OpenProductTable
If optProd7.Value = True Then
    lblProdID.Caption = "P0007"
    
    prodID = lblProdID
    rs1.Index = "PRODUCT_ID"
    rs1.Seek "=", prodID
    If rs1.NoMatch Then
        MsgBox "No record found", vbOKOnly, "Error"
    Else
        lblProdName = rs1!PRODUCT_NAME
        lblProdPrice = rs1!PRODUCT_PRICE
        lblProdDesc = rs1!PRODUCT_DESCRIPTION
    End If
End If
rs1.Close
Set rs1 = Nothing
End Sub

Private Sub optProd8_Click()
OpenKYAKINDatabase
OpenProductTable
If optProd8.Value = True Then
    lblProdID.Caption = "P0008"
    
    prodID = lblProdID
    rs1.Index = "PRODUCT_ID"
    rs1.Seek "=", prodID
    If rs1.NoMatch Then
        MsgBox "No record found", vbOKOnly, "Error"
    Else
        lblProdName = rs1!PRODUCT_NAME
        lblProdPrice = rs1!PRODUCT_PRICE
        lblProdDesc = rs1!PRODUCT_DESCRIPTION
    End If
End If
rs1.Close
Set rs1 = Nothing
End Sub

Private Sub optProd9_Click()
OpenKYAKINDatabase
OpenProductTable
If optProd9.Value = True Then
    lblProdID.Caption = "P0009"
    
    prodID = lblProdID
    rs1.Index = "PRODUCT_ID"
    rs1.Seek "=", prodID
    If rs1.NoMatch Then
        MsgBox "No record found", vbOKOnly, "Error"
    Else
        lblProdName = rs1!PRODUCT_NAME
        lblProdPrice = rs1!PRODUCT_PRICE
        lblProdDesc = rs1!PRODUCT_DESCRIPTION
    End If
End If
rs1.Close
Set rs1 = Nothing
End Sub
