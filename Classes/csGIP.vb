
Public Class csGIP
    Inherits csSignature

    Public str_SiteID As String
    Public objGIPMain As New csGIPMain
    Public objGIPSub As New csGIPSub
    Public objProject As csProjectDetail


    Public Sub New(ByVal CustomerSetting As Dictionary(Of String, String))
        ''If CustomerSetting.Item("Project").ToString = "True" Then
        objProject = New csProjectDetail
        ''End If
    End Sub


    Public Function DBTemplate() As DataTable
        Dim DT_Template As New DataTable
        DT_Template.Columns.Add(New DataColumn("SlNo", System.Type.GetType("System.Int32")))
        DT_Template.Columns.Add(New DataColumn("Ledger"))
        DT_Template.Columns.Add(New DataColumn("DstLedgerID", System.Type.GetType("System.Int32")))
        DT_Template.Columns.Add(New DataColumn("Price", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("Amount", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("Comment"))
        DT_Template.Columns.Add(New DataColumn("ItemDesc"))
        DT_Template.Columns.Add(New DataColumn("BaseUnit", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("VouQty", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("PrimaryQty", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("BaseUnitPrice", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("TCDiscountAmount", System.Type.GetType("System.Double")))
        DT_Template.Columns.Add(New DataColumn("Unit"))
        DT_Template.Columns.Add(New DataColumn("ItemCode"))
        DT_Template.Columns.Add(New DataColumn("IsConsumable", GetType(Boolean)))

        DT_Template.Columns("BaseUnit").DefaultValue = 1
        DT_Template.Columns("VouQty").DefaultValue = 1
        DT_Template.Columns("PrimaryQty").DefaultValue = 1
        DT_Template.Columns("Price").DefaultValue = 0
        DT_Template.Columns("BaseUnitPrice").DefaultValue = 0
        DT_Template.Columns("Amount").DefaultValue = 0
        DT_Template.Columns("IsConsumable").DefaultValue = False

        DT_Template.Columns("SlNo").AutoIncrement = True
        DT_Template.Columns("SlNo").AutoIncrementStep = 1
        DT_Template.Columns("SlNo").AutoIncrementSeed = 1
        Return DT_Template
    End Function

    Public Function VouMatching() As DataTable
        Dim dt_VouMatching As New DataTable
        dt_VouMatching.Columns.Add(New DataColumn("SlNo", GetType(Integer)))
        dt_VouMatching.Columns.Add(New DataColumn("BC_Ref", GetType(Integer)))
        dt_VouMatching.Columns.Add(New DataColumn("ChequeNo"))
        dt_VouMatching.Columns.Add(New DataColumn("Voucher"))
        dt_VouMatching.Columns.Add(New DataColumn("VouRef"))
        dt_VouMatching.Columns.Add(New DataColumn("PayType"))
        dt_VouMatching.Columns.Add(New DataColumn("PDCType"))
        dt_VouMatching.Columns.Add(New DataColumn("Amount", System.Type.GetType("System.Double")))
        dt_VouMatching.Columns.Add(New DataColumn("RcvdAmt", System.Type.GetType("System.Double")))
        dt_VouMatching.Columns.Add(New DataColumn("PDCAmt", System.Type.GetType("System.Double")))
        dt_VouMatching.Columns.Add(New DataColumn("BalAmt", System.Type.GetType("System.Double")))
        dt_VouMatching.Columns.Add(New DataColumn("PayNow", System.Type.GetType("System.Double")))
        dt_VouMatching.Columns.Add(New DataColumn("FullPay", GetType(Boolean)))
        dt_VouMatching.Columns.Add(New DataColumn("RefNo"))
        Return dt_VouMatching
    End Function
End Class

Public Class csGIPMain
    Public str_Flag As String
    Public str_MenuID As String
    Public str_Prefix As String

    Public int_BusinessPeriodID As Integer
    Public str_GIPNo As String
    Public int_RevNo As Integer
    Public str_LpoNo As String
    Public str_SrcLedgerID As Integer
    Public str_Alias As String
    Public dtp_InvDate As Date
    Public bool_IsCashInvoice As Boolean
    Public bool_AffectInventory As Boolean
    Public str_PaymentStatus As String
    Public str_Comment As String
    Public dbl_TCTotalAmount As Double
    Public dbl_TCDisAmount As Double
    Public dbl_TCDiscountAmount As Double
    Public dbl_TCTaxAmount As Double
    Public dbl_TCNetAmount As Double
    Public dbl_TCPDCAmount As Double
    Public dbl_TCPaidAmount As Double
    Public dbl_LCNetAmount As Double
    Public dbl_LCPDCAmount As Double
    Public str_TCCurrency As String
    Public dbl_ExchangeRate As Double
    Public str_RefNo As String
    Public int_Aging As Integer
    Public str_PayTerm As String
    Public str_DelivAddress As String
    Public int_CashLedger As Integer
    Public dbl_CashTendered As Double

    Public str_WHID As String
End Class

Public Class csGIPSub
    Public dt_ExpenseLedger As DataTable
    Public dt_InvMatching As DataTable
End Class
