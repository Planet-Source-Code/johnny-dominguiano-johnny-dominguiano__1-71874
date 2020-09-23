Attribute VB_Name = "Module1"
Global conn As New ADODB.Connection
Global rs As New ADODB.Recordset
Global recset As New ADODB.Recordset
Global rsSubtractBookQty As New ADODB.Recordset


Public Sub connect()
Set conn = New ADODB.Connection
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\Database.mdb;"
End Sub

Public Sub setlock1(val As Boolean)
With frmMembers
    .txtID.Locked = val
    .txtLastname.Locked = val
    .txtFirstname.Locked = val
    .txtMI.Locked = val
    .cboSex.Locked = val
    .txtContactnum.Locked = val
    .cboLevel.Locked = val
    .cboYear.Locked = val
End With
End Sub
Public Sub setlock2(val As Boolean)
With frmBooks
    .txtID.Locked = val
    .txtTitle.Locked = val
    .txtEdition.Locked = val
    .DataCombo1.Locked = val
    .txtAuthor.Locked = val
    .txtPublisher.Locked = val
    .txtISBN.Locked = val
    .txtCopies.Locked = val
    .txtPages.Locked = val
    .txtCallnum.Locked = val
End With
End Sub



Public Sub setbutton1(val As Boolean)
With frmMembers
    .cmdDelete.Enabled = val
    .cmdEdit.Enabled = val
    .cmdSave.Enabled = Not val
    .cmdCancel.Enabled = Not val
    End With
End Sub

Public Sub setbutton2(val As Boolean)
With frmBooks
    .cmdDelete.Enabled = val
    .cmdEdit.Enabled = val
    .cmdSave.Enabled = Not val
    .cmdCancel.Enabled = Not val
    End With
End Sub
