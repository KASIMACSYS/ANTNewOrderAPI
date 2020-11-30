'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Public Class csJV
    Inherits csSignature
    Public int_CID As String
    Public ObjJVMain As New csJVMain
    Public ObjJVSub As New csJVSub
    Public objproject As New csProjectDetail

    Public Function DBTemplate() As DataTable
        Dim DT_Template As New DataTable
        DT_Template.Columns.Add(New DataColumn("SlNo", System.Type.GetType("System.Int32")))
        DT_Template.Columns.Add(New DataColumn("RefNo", System.Type.GetType("System.Int32")))
        DT_Template.Columns.Add(New DataColumn("MatNo", System.Type.GetType("System.Int32")))
        DT_Template.Columns.Add(New DataColumn("Ledger"))
        DT_Template.Columns.Add(New DataColumn("LedgerID", System.Type.GetType("System.Int32")))
        DT_Template.Columns.Add(New DataColumn("Category"))
        DT_Template.Columns.Add(New DataColumn("Comment"))
        DT_Template.Columns.Add(New DataColumn("Dr", System.Type.GetType("System.Decimal")))
        DT_Template.Columns.Add(New DataColumn("Cr", System.Type.GetType("System.Decimal")))

        DT_Template.Columns("Dr").DefaultValue = 0.0
        DT_Template.Columns("Cr").DefaultValue = 0.0

        DT_Template.Columns("SlNo").AutoIncrement = True
        DT_Template.Columns("SlNo").AutoIncrementStep = 1
        DT_Template.Columns("SlNo").AutoIncrementSeed = 1
        Return DT_Template
    End Function
    Public Function TVPJVSubTemplate() As DataTable
        Dim DT_Template As New DataTable
        DT_Template.Columns.Add(New DataColumn("SlNo", System.Type.GetType("System.Int32")))
        DT_Template.Columns.Add(New DataColumn("LedgerID", System.Type.GetType("System.Int32")))
        DT_Template.Columns.Add(New DataColumn("TCDebit", System.Type.GetType("System.Decimal")))
        DT_Template.Columns.Add(New DataColumn("TCCredit", System.Type.GetType("System.Decimal")))
        DT_Template.Columns.Add("Tax")
        DT_Template.Columns.Add("TaxPercentage", System.Type.GetType("System.Double"))
        DT_Template.Columns.Add("TaxAmount", System.Type.GetType("System.Decimal"))
        DT_Template.Columns.Add("NetAmount", System.Type.GetType("System.Decimal"))
        DT_Template.Columns.Add(New DataColumn("Comment"))
        DT_Template.Columns.Add(New DataColumn("Category"))
        Return DT_Template
    End Function


    Public Shared Function DT_VouMatching() As DataTable
        DT_VouMatching = New DataTable
        DT_VouMatching.Columns.Add(New DataColumn("SlNo", GetType(Integer)))
        DT_VouMatching.Columns.Add(New DataColumn("MatNo", GetType(Integer)))
        DT_VouMatching.Columns.Add(New DataColumn("BC_Ref")) ', GetType(Integer)))
        DT_VouMatching.Columns.Add(New DataColumn("ChequeNo"))
        DT_VouMatching.Columns.Add(New DataColumn("Voucher"))
        DT_VouMatching.Columns.Add(New DataColumn("VouRef"))
        DT_VouMatching.Columns.Add(New DataColumn("PayType"))
        DT_VouMatching.Columns.Add(New DataColumn("VouType"))
        DT_VouMatching.Columns.Add(New DataColumn("Amount", System.Type.GetType("System.Decimal")))
        DT_VouMatching.Columns.Add(New DataColumn("PaidAmt", System.Type.GetType("System.Decimal")))
        DT_VouMatching.Columns.Add(New DataColumn("PDCAmt", System.Type.GetType("System.Decimal")))
        DT_VouMatching.Columns.Add(New DataColumn("BalAmt", System.Type.GetType("System.Decimal")))
        DT_VouMatching.Columns.Add(New DataColumn("PayNow", System.Type.GetType("System.Decimal")))
        DT_VouMatching.Columns.Add(New DataColumn("FullPay", GetType(Boolean)))
        DT_VouMatching.Columns.Add(New DataColumn("RefNo"))

        Return DT_VouMatching
    End Function

    Public Class csJVMain
        Public int_BusinessPeriodID As Integer
        Public str_Flag As String
        Public str_MenuID As String
        Public str_Prefix As String

        Public str_JVNo As String
        Public int_RevNo As Integer
        Public dtp_JVDate As Date
        Public str_Comment As String
        Public dbl_TCNetAmount As Decimal
        Public dbl_LCNetAmount As Decimal
        Public dbl_TCTaxAmount As Decimal
        Public str_TCCurrency As String
        Public dbl_ExchangeRate As Double
        Public bool_ApprovedStatus As Boolean
        Public str_UserComment As String = String.Empty
        Public int_StatusCancel As Integer
        Public str_BatchID As String
        Public int_LanguageCode As Integer
        Public str_RefNo As String
    End Class

    Public Class csJVSub
        Public dt_JVSub As DataTable
        Public dt_JVMatching As DataTable
        Public dt_Wages As New DataTable
        Public dt_TaxItemDetails As DataTable
    End Class
End Class
