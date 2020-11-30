'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Public Class csPayCert
    Inherits csSignature
    Public str_CID As String
    Public objPayCertSub As New csPayCertSub
    Public objPayCertMain As New csPayCertMain
    Public objproject As New csProjectDetail
    Public Sub New(ByVal CustomerSetting As Dictionary(Of String, String))
        'If CustomerSetting.Item("useProject").ToString = "True" Then
        objproject = New csProjectDetail
        'End If
    End Sub
    Public Class csPayCertMain
        Public str_Flag As String
        Public str_MenuID As String
        Public int_BusinessPeriod As Integer
        Public str_PCNo As String
        Public dtp_PCDate As Date
        Public int_LedgerID As Integer
        Public str_MerchantName As String
        Public int_PCStatus As Integer
        Public dbl_TCAmount As Double
        Public str_Comment As String
        Public dtp_INVDate As Date

        Public bool_ApprovedStatus As Boolean
        Public str_PayCertPrefix As String
        Public str_PayTerm As String
        Public int_RevNo As Integer
        Public dt_getpip As DataTable
        Public dbl_LCNetAmount As Double
        Public str_TCCurrency As String
        Public dbl_ExchangeRate As Double
        Public int_Aging As Integer
        Public dbl_TotApproveAmt As Double
        Public str_setFlag As String
        Public str_createdby As String
        Public str_Desc1 As String
        Public str_Desc2 As String
        Public str_Desc3 As String
        Public str_Desc4 As String
        Public str_Desc5 As String
        Public str_Desc6 As String
        Public str_Desc7 As String
        Public str_Desc8 As String

        Public dbl_AdvanceAmount As Double

    End Class
    Public Class csPayCertSub
        'Public str_INVNo As String
        'Public str_PIP As String
        'Public str_LPO As String
        'Public dbl_TCAmount As Double
        'Public bool_Approve As Boolean
        'Public str_Project As String
        Public dt_PayCert As DataTable
    End Class
End Class
