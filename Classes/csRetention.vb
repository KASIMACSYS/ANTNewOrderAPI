
Public Class csRetention
    Private Flag As String = String.Empty
    Public dbl_RetAmtDeduction As Double
    Public dbl_RetAmtAddition As Double
    Public dtp_RetDueDate As Date
    Public dt_RetentionInvoiceList As DataTable

    Public Sub New()
        dbl_RetAmtAddition = 0
        dbl_RetAmtDeduction = 0
    End Sub

    Public Sub New(ByVal _Flag As String)
        Flag = _Flag
    End Sub

    Public Sub RetensionTemp()
        If dt_RetentionInvoiceList IsNot Nothing Then
            dt_RetentionInvoiceList.Clear()
        Else
            dt_RetentionInvoiceList = New DataTable
            dt_RetentionInvoiceList.Columns.Add(New DataColumn("InvoiceNo"))
            dt_RetentionInvoiceList.Columns.Add(New DataColumn("InvDate", System.Type.GetType("System.DateTime")))
            dt_RetentionInvoiceList.Columns.Add(New DataColumn("DueDate", System.Type.GetType("System.DateTime")))
            dt_RetentionInvoiceList.Columns.Add(New DataColumn("RententionAmt", System.Type.GetType("System.Double")))
            dt_RetentionInvoiceList.Columns.Add(New DataColumn("Select", System.Type.GetType("System.Boolean")))
            dt_RetentionInvoiceList.AcceptChanges()
        End If
    End Sub

End Class
