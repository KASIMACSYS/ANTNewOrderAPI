'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Public Class csGeneralVoucher
    Inherits csSignature
    Public int_CID As Integer

    Public objGVMain As New GVMain
    Public DT_GVDetails As New DataTable
    Public dt_VouMatching As DataTable
    Public objProject As csProjectDetail
    Public dt_TaxItemDetails As DataTable

    Public Sub New(ByVal CustomerSetting As Dictionary(Of String, String))
        'If CustomerSetting.Item("Project").ToString = "True" Then
        objProject = New csProjectDetail
        ' End If
    End Sub

    Public Function DT_GVDetailsTemplate() As DataTable
        DT_GVDetailsTemplate = New DataTable
        DT_GVDetailsTemplate.Columns.Add(New DataColumn("SlNo", System.Type.GetType("System.Int32")))
        DT_GVDetailsTemplate.Columns.Add(New DataColumn("DstLedger"))
        DT_GVDetailsTemplate.Columns.Add(New DataColumn("DstLedgerID", System.Type.GetType("System.Int32")))
        DT_GVDetailsTemplate.Columns.Add(New DataColumn("TCAmount", System.Type.GetType("System.Double")))
        DT_GVDetailsTemplate.Columns.Add(New DataColumn("LCAmount", System.Type.GetType("System.Double")))
        DT_GVDetailsTemplate.Columns.Add(New DataColumn("TRNNo"))
        DT_GVDetailsTemplate.Columns.Add("Tax")
        DT_GVDetailsTemplate.Columns.Add("TaxPercentage", System.Type.GetType("System.Double"))
        DT_GVDetailsTemplate.Columns.Add("TaxAmount", System.Type.GetType("System.Decimal"))
        DT_GVDetailsTemplate.Columns.Add("NetAmount", System.Type.GetType("System.Decimal"))
        DT_GVDetailsTemplate.Columns.Add(New DataColumn("Comment"))
        Return DT_GVDetailsTemplate
    End Function
End Class

Public Class GVMain
    Public int_BusinessPeriodID As Integer
    Public str_Prefix As String
    Public str_MenuID As String

    Public str_VouNo As String
    Public int_RevNo As Integer
    Public str_Type As String
    Public str_VouRef As String
    Public int_SrcLedgerID As Integer
    Public dtp_VouDate As Date
    Public dbl_TCAmount As Double
    Public dbl_LCAmount As Double
    Public str_TCCurrency As String
    Public dbl_TCTaxAmount As Double
    Public dbl_ExchangeRate As Double
    Public str_Comment As String
    Public str_LedgerDepartment As String

    Public dt_GenVou As DataTable
    Public str_CreatedBy As String

    Public str_FormType As String
    Public str_Flag As String
    Public str_Desc1 As String
    Public str_Desc2 As String
    Public str_Desc3 As String
    Public str_Desc4 As String
    Public str_Desc5 As String
    Public str_Desc6 As String
    Public str_Desc7 As String
    Public str_Desc8 As String
    Public int_StatusCancel As Integer
End Class

Public Class csDTTemplate4GV

    Public Shared Function Template_VouMatching() As DataTable
        Template_VouMatching = New DataTable
        Template_VouMatching.Columns.Add(New DataColumn("SlNo", GetType(Integer)))
        Template_VouMatching.Columns.Add(New DataColumn("MatNo", GetType(Integer)))
        Template_VouMatching.Columns.Add(New DataColumn("BC_Ref", GetType(Integer)))
        Template_VouMatching.Columns.Add(New DataColumn("ChequeNo"))
        Template_VouMatching.Columns.Add(New DataColumn("Date_", GetType(Date)))
        Template_VouMatching.Columns.Add(New DataColumn("Voucher"))
        Template_VouMatching.Columns.Add(New DataColumn("VouRef"))
        Template_VouMatching.Columns.Add(New DataColumn("PayType"))
        Template_VouMatching.Columns.Add(New DataColumn("VouType"))
        Template_VouMatching.Columns.Add(New DataColumn("Amount", System.Type.GetType("System.Decimal")))
        Template_VouMatching.Columns.Add(New DataColumn("PaidAmt", System.Type.GetType("System.Decimal")))
        Template_VouMatching.Columns.Add(New DataColumn("PDCAmt", System.Type.GetType("System.Decimal")))
        Template_VouMatching.Columns.Add(New DataColumn("BalAmt", System.Type.GetType("System.Decimal")))
        Template_VouMatching.Columns.Add(New DataColumn("PayNow", System.Type.GetType("System.Decimal")))
        Template_VouMatching.Columns.Add(New DataColumn("FullPay", GetType(Boolean)))
        Template_VouMatching.Columns.Add(New DataColumn("RefNo"))
        Template_VouMatching.Columns("BC_Ref").DefaultValue = 0
        Return Template_VouMatching
    End Function

    Public Shared Function DT_4Dialog() As DataTable
        DT_4Dialog = New DataTable
        DT_4Dialog.Columns.Add(New DataColumn("SlNo", GetType(Integer)))
        DT_4Dialog.Columns.Add(New DataColumn("BC_Ref", GetType(Integer)))
        DT_4Dialog.Columns.Add(New DataColumn("Date_", GetType(Date)))
        DT_4Dialog.Columns.Add(New DataColumn("ChequeNo"))
        DT_4Dialog.Columns.Add(New DataColumn("Voucher"))
        DT_4Dialog.Columns.Add(New DataColumn("VouRef"))
        DT_4Dialog.Columns.Add(New DataColumn("PayType"))
        DT_4Dialog.Columns.Add(New DataColumn("VouType"))
        DT_4Dialog.Columns.Add(New DataColumn("Amount", System.Type.GetType("System.Decimal")))
        DT_4Dialog.Columns.Add(New DataColumn("PaidAmt", System.Type.GetType("System.Decimal")))
        DT_4Dialog.Columns.Add(New DataColumn("PDCAmt", System.Type.GetType("System.Decimal")))
        DT_4Dialog.Columns.Add(New DataColumn("BalAmt", System.Type.GetType("System.Decimal")))
        DT_4Dialog.Columns.Add(New DataColumn("PayNow", System.Type.GetType("System.Decimal")))
        DT_4Dialog.Columns("PayType").DefaultValue = False
        DT_4Dialog.Columns("ChequeNo").DefaultValue = False
        Return DT_4Dialog
    End Function
End Class
