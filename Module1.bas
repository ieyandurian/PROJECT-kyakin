Attribute VB_Name = "Module1"
Option Explicit
Public dbkyakin As Database
Public rs As Recordset
Public rs1 As Recordset
Public rs2 As Recordset
Public rs3 As Recordset
Public rs4 As Recordset
Public rs5 As Recordset
Public gintHelpFileNbr As Integer
Public usern As String
Public passn As String

 
'------------------------------------------------------------------------
Public Sub OpenKYAKINDatabase()
'------------------------------------------------------------------------
Set dbkyakin = OpenDatabase(GetAppPath() & "kyakin.mdb")
 
End Sub
Public Sub OpenProductTable()

Set rs1 = dbkyakin.OpenRecordset("PRODUCT")

End Sub
Public Sub OpenUserTable()

Set rs = dbkyakin.OpenRecordset("USER")

End Sub
Public Sub OpenOrderTable()

Set rs3 = dbkyakin.OpenRecordset("ORDER")

End Sub
Public Sub OpenDeliveryTable()

Set rs4 = dbkyakin.OpenRecordset("DELIVERY")

End Sub
Public Sub OpenInvoiceTable()

Set rs5 = dbkyakin.OpenRecordset("INVOICE")

End Sub
 
'------------------------------------------------------------------------
Public Sub CloseKYAKINDatabase()
'------------------------------------------------------------------------
 
dbkyakin.Close
Set dbkyakin = Nothing
 
End Sub
 
'------------------------------------------------------------------------
Public Sub CenterForm(pobjForm As Form)
'------------------------------------------------------------------------
 
With pobjForm
.Top = (Screen.Height - .Height) / 2
.Left = (Screen.Width - .Width) / 2
End With
 
End Sub
 
'------------------------------------------------------------------------
Public Function GetAppPath() As String
'------------------------------------------------------------------------
 
GetAppPath = IIf(Right$(App.Path, 1) = "\", App.Path, App.Path & "\")
 
End Function

