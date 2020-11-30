'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Imports Classes

Public Class DAL_CounterPurchase
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef Obj As csPurchaseInvoice, ByRef ErrorNo As Integer, ByRef ErrString As String)
        ErrorNo = 0
        ErrString = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetCounterPurchaseDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure

            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.objPurInvMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CPNo", Obj.objPurInvMain.str_InvNo)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.objPurInvMain.str_Flag)
            BaseConn.cmd.Parameters.Add("@CashLedger", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@CashTendered", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

           

            If Obj.objPurInvMain.str_Flag = "CP" Then
                Obj.objPurInvMain.int_CashLedger = BaseConn.cmd.Parameters("@CashLedger").Value
                Obj.objPurInvMain.dbl_CashTendered = BaseConn.cmd.Parameters("@CashTendered").Value

                Obj.objPurInvMain.str_SrcLedgerID = ds.Tables(0).Rows(0)("SrcLedgerID").ToString()
                Obj.objPurInvMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
                Obj.objPurInvMain.str_InvRef = ds.Tables(0).Rows(0)("InvRef").ToString()
                Obj.objPurInvMain.str_InvNo = ds.Tables(0).Rows(0)("InvNo").ToString()
                Obj.objPurInvMain.str_LpoNo = ds.Tables(0).Rows(0)("LPONo").ToString()
                Obj.objPurInvMain.dtp_InvDate = ds.Tables(0).Rows(0)("InvDate").ToString()
                Obj.objPurInvMain.dtp_DueDate = ds.Tables(0).Rows(0)("DueDate").ToString()
                Obj.objPurInvMain.int_Aging = ds.Tables(0).Rows(0)("Aging").ToString()
                Obj.objPurInvMain.str_PayTerm = ds.Tables(0).Rows(0)("PaymentTerm").ToString()
                Obj.objPurInvMain.str_InvoiceType = ds.Tables(0).Rows(0)("InvoiceType").ToString()
                Obj.objPurInvMain.dbl_TCAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
                Obj.objPurInvMain.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
                Obj.objPurInvMain.dbl_TCMiscPercentage = ds.Tables(0).Rows(0)("TCMiscPercentage").ToString()
                Obj.objPurInvMain.dbl_TCMiscAmount = ds.Tables(0).Rows(0)("TCMiscAmount").ToString()
                'Obj.objPurInvMain.dbl_TCTaxAmount = ds.Tables(0).Rows(0)("TCTaxAmount").ToString()
                Obj.objPurInvMain.dbl_TCItemTaxAmount = ds.Tables(0).Rows(0)("TCItemTaxAmount")
                Obj.objPurInvMain.dbl_TCInvoiceTaxAmount = ds.Tables(0).Rows(0)("TCInvTaxAmount")
                Obj.objPurInvMain.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
                Obj.objPurInvMain.dbl_TCAdjAmount = ds.Tables(0).Rows(0)("TCAdjAmount").ToString()
                Obj.objPurInvMain.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
                Obj.objPurInvMain.dbl_TCPDCAmount = ds.Tables(0).Rows(0)("TCPDCAmount").ToString()
                Obj.objPurInvMain.dbl_LCLandingCost = ds.Tables(0).Rows(0)("LCLandingCost").ToString()
                Obj.objPurInvMain.bool_PurchaseInvoice = ds.Tables(0).Rows(0)("CounterPurchase").ToString()
                Obj.objPurInvMain.bool_IsCashPurchase = ds.Tables(0).Rows(0)("IsCashPurchase").ToString()
                Obj.objPurInvMain.bool_AffectInventory = ds.Tables(0).Rows(0)("AffectInventory").ToString()
                Obj.objPurInvMain.str_PaymentStatus = ds.Tables(0).Rows(0)("PaymentStatus").ToString()
                Obj.objPurInvMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
                Obj.objPurInvMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
                Obj.objPurInvMain.int_StatusCancel = ds.Tables(0).Rows(0)("StatusCancel").ToString()
                Obj.objPurInvMain.str_DeliveryAddress = ds.Tables(0).Rows(0)("DeliveryAddress").ToString()

                Obj.objPurInvMain.str_WHID = ds.Tables(0).Rows(0)("WHID").ToString()

                Obj.objPurInvMain.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
                Obj.objPurInvMain.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
                Obj.objPurInvMain.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
                Obj.objPurInvMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
                Obj.objPurInvMain.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
                Obj.objPurInvMain.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
                Obj.objPurInvMain.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
                Obj.objPurInvMain.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()

                Obj.objPurInvMain.dbl_LCPDCAmount = 0
                Obj.objPurInvMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()

                Obj.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
                Obj.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
                Obj.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
                Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
                Obj.objPurInvMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()

                Obj.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
                Obj.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
                Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus")

                Obj.objPurInvMain.str_CurrencyID = ds.Tables(0).Rows(0)("TCCurrency").ToString()
                Obj.objPurInvMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()
                Obj.objPurInvMain.str_MiscText = ds.Tables(0).Rows(0)("MiscText").ToString()
                Obj.objPurInvMain.str_DiscText = ds.Tables(0).Rows(0)("DiscText").ToString()
                Obj.objPurInvMain.str_ItemTaxCode = ds.Tables(0).Rows(0)("ItemTaxCode")
                Obj.objPurInvMain.str_InvoiceTaxCode = ds.Tables(0).Rows(0)("InvoiceTaxCode")
                Obj.objPurInvMain.str_PermitNo = ds.Tables(0).Rows(0)("PermitNo").ToString()
                Obj.objPurInvMain.dbl_ItemDiscPercentage = ds.Tables(0).Rows(0)("ItemDiscPercentage").ToString()
                Obj.objPurInvMain.str_InvoiceTaxXML = ds.Tables(0).Rows(0)("InvoiceTaxDetails")
                Obj.objPurInvMain.str_MerchantRef = ds.Tables(0).Rows(0)("MerchantRef").ToString()
                Obj.objPurInvMain.str_ContactPerson = ds.Tables(0).Rows(0)("ContactPerson").ToString()
                If ds.Tables(1).Rows.Count > 0 Then
                    Obj.objPurInvSub.dt_PurInv = ds.Tables(1)
                End If

                If ds.Tables(2).Rows.Count > 0 Then
                    Obj.objProject.str_ProjectID = ds.Tables(2).Rows(0)("ProjectID").ToString()
                    Obj.objProject.str_ProjectLocation = ds.Tables(2).Rows(0)("ProjectLocation").ToString()
                    Obj.objProject.str_WorkOrderNo = ds.Tables(2).Rows(0)("WorkOrderNo").ToString()
                Else
                    Obj.objProject.str_ProjectID = ""
                    Obj.objProject.str_ProjectLocation = ""
                    Obj.objProject.str_WorkOrderNo = ""
                End If

                If ds.Tables(3).Rows.Count > 0 Then
                    Obj.objPurInvMain.dt_InvoiceAccounts = ds.Tables(3)
                End If
                Obj.DTBatch = ds.Tables(4)
                Obj.objPurInvMain.dt_TaxItemDetails = ds.Tables(5)
            Else
                Obj.objPurInvMain.str_SrcLedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
                Obj.objPurInvMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
                Obj.objPurInvMain.dtp_InvDate = Date.Now
                Obj.objPurInvMain.dtp_DueDate = ds.Tables(0).Rows(0)("LPODate1").ToString()
                Obj.objPurInvMain.int_Aging = ds.Tables(0).Rows(0)("Aging").ToString()
                Obj.objPurInvMain.str_PayTerm = ds.Tables(0).Rows(0)("PayTerm").ToString()
                Obj.objPurInvMain.dbl_TCAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
                Obj.objPurInvMain.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
                Obj.objPurInvMain.dbl_TCMiscAmount = ds.Tables(0).Rows(0)("TCMiscAmount").ToString()
                Obj.objPurInvMain.dbl_TCMiscPercentage = ds.Tables(0).Rows(0)("TCMiscPercentage").ToString()
                'Obj.objPurInvMain.dbl_TCTaxAmount = ds.Tables(0).Rows(0)("TCTaxAmount").ToString()
                Obj.objPurInvMain.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
                Obj.objPurInvMain.dbl_TCAdjAmount = ds.Tables(0).Rows(0)("TCAdjAmount").ToString()
                Obj.objPurInvMain.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
                Obj.objPurInvMain.dbl_TCPDCAmount = 0 ' ds.Tables(0).Rows(0)("TCPDCAmount").ToString()
                Obj.objPurInvMain.bool_PurchaseInvoice = 0 'ds.Tables(0).Rows(0)("CounterSales").ToString()
                Obj.objPurInvMain.bool_IsCashPurchase = 0 'ds.Tables(0).Rows(0)("IsCashSales").ToString()
                Obj.objPurInvMain.bool_AffectInventory = 1 ' ds.Tables(0).Rows(0)("AffectInventory").ToString()
                Obj.objPurInvMain.str_PaymentStatus = "" ' ds.Tables(0).Rows(0)("PaymentStatus").ToString()
                Obj.objPurInvMain.int_RevNo = 0 ' ds.Tables(0).Rows(0)("RevNo").ToString()
                Obj.objPurInvMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
                'Obj.objPurInvMain.str_SalesManID = ds.Tables(0).Rows(0)("SalesManID").ToString()
                Obj.objPurInvMain.dbl_ItemDiscPercentage = ds.Tables(0).Rows(0)("ItemDiscPercentage").ToString()
                Obj.objPurInvMain.str_LpoNo = "" ' ds.Tables(0).Rows(0)("LpoNo").ToString()
                'Obj.objPurInvMain.str_DONo = "" ' ds.Tables(0).Rows(0)("DoNo").ToString()
                Obj.objPurInvMain.str_LpoNo = ds.Tables(0).Rows(0)("LPONo").ToString()
                'Obj.objPurInvMain.dbl_LCNetCostPrice = 0 'ds.Tables(0).Rows(0)("LCNetCostAmount").ToString()
                Obj.objPurInvMain.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString()
                'Obj.objPurInvMain.dbl_LCPDCAmount = ds.Tables(0).Rows(0)("LCPDCAmount").ToString()
                Obj.objPurInvMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()
                Obj.objPurInvMain.str_MiscText = ds.Tables(0).Rows(0)("MiscText").ToString
                Obj.objPurInvMain.str_DiscText = ds.Tables(0).Rows(0)("DiscText").ToString()
                Obj.objPurInvMain.str_WHID = 101 ' ds.Tables(0).Rows(0)("WHID").ToString()
                Obj.objPurInvMain.str_UserComment = ds.Tables(0).Rows(0)("UserComment").ToString()

                Obj.objPurInvMain.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
                Obj.objPurInvMain.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
                Obj.objPurInvMain.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
                Obj.objPurInvMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
                Obj.objPurInvMain.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
                Obj.objPurInvMain.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
                Obj.objPurInvMain.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
                Obj.objPurInvMain.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()
                Obj.objPurInvMain.str_ItemTaxCode = ds.Tables(0).Rows(0)("ItemTaxCode")
                Obj.objPurInvMain.str_InvoiceTaxCode = ds.Tables(0).Rows(0)("InvoiceTaxCode")
                Obj.objPurInvMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()

                Obj.objPurInvMain.str_CurrencyID = ds.Tables(0).Rows(0)("TCCurrency").ToString()
                Obj.objPurInvMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()
                Obj.objPurInvMain.str_InvoiceTaxXML = ds.Tables(0).Rows(0)("InvoiceTaxDetails")

                If ds.Tables(1).Rows.Count > 0 Then
                    Obj.objPurInvSub.dt_PurInv = ds.Tables(1)
                End If

                If ds.Tables(2).Rows.Count > 0 Then
                    Obj.objProject.str_ProjectID = ds.Tables(2).Rows(0)("ProjectID").ToString()
                    Obj.objProject.str_ProjectLocation = ds.Tables(2).Rows(0)("ProjectLocation").ToString()
                    Obj.objProject.str_WorkOrderNo = ds.Tables(2).Rows(0)("WorkOrderNo").ToString()
                Else
                    Obj.objProject.str_ProjectID = ""
                    Obj.objProject.str_ProjectLocation = ""
                    Obj.objProject.str_WorkOrderNo = ""
                End If
                'Obj.objPurInvMain.dt_TaxItemDetails = ds.Tables(3)
            End If
        Catch ex As Exception
            ErrorNo = 1
            ErrString = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub


    Public Function Update_CP(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef CPNo As String, ByRef intRevNo As Integer, ByVal obj As csPurchaseInvoice, ByRef _OutSMS As String, ByRef _OutEmail As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("InvoiceAccountUpdate_CP", BaseConn.cnn)

            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.int_CID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.objPurInvMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.objPurInvMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.objPurInvMain.str_Prefix)
            'BaseConn.cmd.Parameters.AddWithValue("@MenuID", "Menu_305")

            BaseConn.cmd.Parameters.AddWithValue("@PIPNo", obj.objPurInvMain.str_PurInvNo)
            BaseConn.cmd.Parameters.AddWithValue("@LPONo", obj.objPurInvMain.str_LpoNo)
            BaseConn.cmd.Parameters.AddWithValue("@InvoiceType", obj.objPurInvMain.str_InvoiceType)

            BaseConn.cmd.Parameters.AddWithValue("@InvRef", obj.objPurInvMain.str_InvRef)
            BaseConn.cmd.Parameters.AddWithValue("@InvNo", obj.objPurInvMain.str_InvNo)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.objPurInvMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@SrcLedgerID", obj.objPurInvMain.str_SrcLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Alias", obj.objPurInvMain.str_Alias)
            BaseConn.cmd.Parameters.AddWithValue("@InvDate", obj.objPurInvMain.dtp_InvDate)
            BaseConn.cmd.Parameters.AddWithValue("@DueDate", obj.objPurInvMain.dtp_DueDate)
            BaseConn.cmd.Parameters.AddWithValue("@Aging", obj.objPurInvMain.int_Aging)
            BaseConn.cmd.Parameters.AddWithValue("@PaymentTerm", obj.objPurInvMain.str_PayTerm)

            BaseConn.cmd.Parameters.AddWithValue("@CounterPurchase", obj.objPurInvMain.bool_PurchaseInvoice)
            BaseConn.cmd.Parameters.AddWithValue("@IsCashPurchase", obj.objPurInvMain.bool_IsCashPurchase)
            BaseConn.cmd.Parameters.AddWithValue("@AffectInventory", obj.objPurInvMain.bool_AffectInventory)
            BaseConn.cmd.Parameters.AddWithValue("@PIPStatus", obj.objPurInvMain.str_InvoiceStatus)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.objPurInvMain.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@TCAmount", obj.objPurInvMain.dbl_TCAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDisAmount", obj.objPurInvMain.dbl_TCDisAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDiscountAmount", obj.objPurInvMain.dbl_TCDiscountAmount)

            BaseConn.cmd.Parameters.AddWithValue("@TCMiscAmount", obj.objPurInvMain.dbl_TCMiscAmount)
            'BaseConn.cmd.Parameters.AddWithValue("@TCTaxAmount", obj.objPurInvMain.dbl_TCTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCItemTaxAmount", obj.objPurInvMain.dbl_TCItemTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCInvoiceTaxAmount", obj.objPurInvMain.dbl_TCInvoiceTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCAdjAmount", obj.objPurInvMain.dbl_TCAdjAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCNetAmount", obj.objPurInvMain.dbl_TCNetAmount)
            'BaseConn.cmd.Parameters.AddWithValue("@TCPDCAmount", obj.objPurInvMain.dbl_TCPDCAmount)
            BaseConn.cmd.Parameters.AddWithValue("@LCNetAmount", obj.objPurInvMain.dbl_LCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@LCLandingCost", obj.objPurInvMain.dbl_LCLandingCost)
            BaseConn.cmd.Parameters.AddWithValue("@StatusCancel", obj.objPurInvMain.int_StatusCancel)
            BaseConn.cmd.Parameters.AddWithValue("@MiscText", obj.objPurInvMain.str_MiscText)
            BaseConn.cmd.Parameters.AddWithValue("@DiscText", obj.objPurInvMain.str_DiscText)
            BaseConn.cmd.Parameters.AddWithValue("@DeliveryAddress", obj.objPurInvMain.str_DeliveryAddress)

            'BaseConn.cmd.Parameters.AddWithValue("@LCPDCAmount", obj.objPurInvMain.dbl_LCPDCAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", obj.objPurInvMain.str_CurrencyID)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", obj.objPurInvMain.dbl_ExchangeRate)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.bool_ApprovedStatus)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedLevel", obj.ApprovedLevel)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedComment", obj.ApprovedComment)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedHigherLevel", obj.ApprovedHigherLevel)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.objPurInvMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CashorCredit", obj.objPurInvMain.str_CashorCredit)
            BaseConn.cmd.Parameters.AddWithValue("@CashLedger", obj.objPurInvMain.int_CashLedger)
            BaseConn.cmd.Parameters.AddWithValue("@CashTendered", obj.objPurInvMain.dbl_CashTendered)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectID", obj.objProject.str_ProjectID)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectLocation", obj.objProject.str_ProjectLocation)
            BaseConn.cmd.Parameters.AddWithValue("@WorkOrderNo", obj.objProject.str_WorkOrderNo)
            BaseConn.cmd.Parameters.AddWithValue("@UserComment", obj.objPurInvMain.str_UserComment)
            BaseConn.cmd.Parameters.AddWithValue("@WHID", obj.objPurInvMain.str_WHID)

            ''AM Specific
            BaseConn.cmd.Parameters.AddWithValue("@Desc1", obj.objPurInvMain.str_Desc1)
            BaseConn.cmd.Parameters.AddWithValue("@Desc2", obj.objPurInvMain.str_Desc2)
            BaseConn.cmd.Parameters.AddWithValue("@Desc3", obj.objPurInvMain.str_Desc3)
            BaseConn.cmd.Parameters.AddWithValue("@Desc4", obj.objPurInvMain.str_Desc4)
            BaseConn.cmd.Parameters.AddWithValue("@Desc5", obj.objPurInvMain.str_Desc5)
            BaseConn.cmd.Parameters.AddWithValue("@Desc6", obj.objPurInvMain.str_Desc6)
            BaseConn.cmd.Parameters.AddWithValue("@Desc7", obj.objPurInvMain.str_Desc7)
            BaseConn.cmd.Parameters.AddWithValue("@Desc8", obj.objPurInvMain.str_Desc8)
            BaseConn.cmd.Parameters.AddWithValue("@ItemTaxCode", obj.objPurInvMain.str_ItemTaxCode)
            BaseConn.cmd.Parameters.AddWithValue("@InvoiceTaxCode", obj.objPurInvMain.str_InvoiceTaxCode)

            BaseConn.cmd.Parameters.AddWithValue("@PermitNo", obj.objPurInvMain.str_PermitNo)
            BaseConn.cmd.Parameters.AddWithValue("@MerchantRef", obj.objPurInvMain.str_MerchantRef)
            BaseConn.cmd.Parameters.AddWithValue("@ContactPerson", obj.objPurInvMain.str_ContactPerson)
            BaseConn.cmd.Parameters.AddWithValue("@ItemDiscPercentage", obj.objPurInvMain.dbl_ItemDiscPercentage)
            BaseConn.cmd.Parameters.AddWithValue("@TCMiscPercentage", obj.objPurInvMain.dbl_TCMiscPercentage)
            BaseConn.cmd.Parameters.AddWithValue("@InvoiceTaxXML", obj.objPurInvMain.str_InvoiceTaxXML)
            BaseConn.cmd.Parameters.AddWithValue("@LanguageCode", obj.objPurInvMain.int_LanguageCode)

            BaseConn.cmd.Parameters.AddWithValue("@CounterPurchaseDT", obj.objPurInvSub.dt_PurInv)
            BaseConn.cmd.Parameters.AddWithValue("@MatchingDT", obj.objPurInvSub.dt_PurInvMatching)
            BaseConn.cmd.Parameters.AddWithValue("@InvAccDetDT", obj.objPurInvSub.dt_InvoiceAccounts)
            BaseConn.cmd.Parameters.AddWithValue("@ItemBatchDT", obj.DTBatch)
            BaseConn.cmd.Parameters.AddWithValue("@InvTaxAmountDT", obj.objPurInvMain.dt_TaxItemDetails)

            BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutSMSMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutEmailMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output

            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            CPNo = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            _OutSMS = BaseConn.cmd.Parameters("@OutSMSMsgText").Value.ToString
            _OutEmail = BaseConn.cmd.Parameters("@OutEmailMsgText").Value.ToString

        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.int_CID)
            ObjDalGeneral.Elog_Insert(obj.int_CID, _StrDBPath, _StrDBPwd, obj.objPurInvMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "CP", Err.Number, "Error in" & obj.objPurInvMain.str_Flag & " : " & obj.objPurInvMain.str_InvNo & "  ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_CP = _ErrString
    End Function

End Class


Public Class DAL_PIP
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub GetMRV4PIP(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _StrSiteID As String, ByVal _BusinessPeriodID As Integer, ByVal _Flag As String, _
                         ByVal _VouNo As String, ByVal _LedgerID As String, ByVal _strTCCurrency As String, ByRef DTUnInvoiced As DataTable, ByRef DTInvoicedDO As DataTable, ByVal _ExchangeRate As Double, ByVal _OtherForm As Boolean)
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetMRV4PIP]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _StrSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", _VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", _strTCCurrency)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", _ExchangeRate)
            BaseConn.cmd.Parameters.AddWithValue("@OtherForm", _OtherForm)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            DTUnInvoiced = ds.Tables(0)
            If _Flag = "EDIT" Then
                DTInvoicedDO = ds.Tables(1)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub Get_Structure(ByVal _DBPath As String, ByVal _DBPwd As String, ByRef Obj As csPurchaseInvoice, ByRef _ErrNo As Integer)
        Try
            _ErrNo = 0
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetPIPDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@CPNo", Obj.objPurInvMain.str_PurInvNo)
            BaseConn.cmd.Parameters.Add("@CashLedger", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@CashTendered", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            Obj.objPurInvMain.int_CashLedger = BaseConn.cmd.Parameters("@CashLedger").Value
            Obj.objPurInvMain.dbl_CashTendered = BaseConn.cmd.Parameters("@CashTendered").Value
            Obj.objPurInvMain.str_SrcLedgerID = ds.Tables(0).Rows(0)("SrcLedgerID").ToString()
            Obj.objPurInvMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
            Obj.objPurInvMain.dtp_InvDate = ds.Tables(0).Rows(0)("InvDate").ToString()
            Obj.objPurInvMain.dtp_DueDate = ds.Tables(0).Rows(0)("DueDate").ToString()
            Obj.objPurInvMain.int_Aging = ds.Tables(0).Rows(0)("Aging").ToString()
            Obj.objPurInvMain.str_PayTerm = ds.Tables(0).Rows(0)("PaymentTerm").ToString()
            Obj.objPurInvMain.str_InvoiceType = ds.Tables(0).Rows(0)("InvoiceType").ToString()
            Obj.objPurInvMain.dbl_TCAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
            Obj.objPurInvMain.dbl_TCMiscPercentage = ds.Tables(0).Rows(0)("TCMiscPercentage").ToString()
            Obj.objPurInvMain.dbl_TCMiscAmount = ds.Tables(0).Rows(0)("TCMiscAmount").ToString()
            Obj.objPurInvMain.dbl_TCItemTaxAmount = ds.Tables(0).Rows(0)("TCItemTaxAmount")
            Obj.objPurInvMain.dbl_TCInvoiceTaxAmount = ds.Tables(0).Rows(0)("TCInvTaxAmount")
            Obj.objPurInvMain.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
            Obj.objPurInvMain.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
            Obj.objPurInvMain.dbl_TCAdjAmount = ds.Tables(0).Rows(0)("TCAdjAmount").ToString()
            Obj.objPurInvMain.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
            Obj.objPurInvMain.dbl_TCPDCAmount = ds.Tables(0).Rows(0)("TCPDCAmount").ToString()
            Obj.objPurInvMain.bool_PurchaseInvoice = ds.Tables(0).Rows(0)("CounterPurchase").ToString()
            Obj.objPurInvMain.bool_IsCashPurchase = ds.Tables(0).Rows(0)("IsCashPurchase").ToString()
            Obj.objPurInvMain.bool_AffectInventory = ds.Tables(0).Rows(0)("AffectInventory").ToString()
            Obj.objPurInvMain.str_PaymentStatus = ds.Tables(0).Rows(0)("PaymentStatus").ToString()
            Obj.objPurInvMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
            Obj.objPurInvMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.objPurInvMain.int_StatusCancel = ds.Tables(0).Rows(0)("StatusCancel")
            Obj.objPurInvMain.str_UserComment = ds.Tables(0).Rows(0)("UserComment").ToString()

            Obj.objPurInvMain.str_InvNo = ds.Tables(0).Rows(0)("InvNo").ToString()
            Obj.objPurInvMain.str_InvRef = ds.Tables(0).Rows(0)("InvRef").ToString()


            Obj.objPurInvMain.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString()
            'Obj.objPurInvMain.dbl_LCPDCAmount = ds.Tables(0).Rows(0)("LCPDCAmount").ToString()
            Obj.objPurInvMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()

            Obj.objPurInvMain.str_ItemTaxCode = ds.Tables(0).Rows(0)("ItemTaxCode")
            Obj.objPurInvMain.str_InvoiceTaxCode = ds.Tables(0).Rows(0)("InvoiceTaxCode")
            Obj.objPurInvMain.str_InvoiceTaxXML = ds.Tables(0).Rows(0)("InvoiceTaxDetails")

            Obj.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
            Obj.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
            Obj.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
            Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
            Obj.objPurInvMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()

            Obj.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
            Obj.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
            Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus")

            Obj.objPurInvMain.str_CurrencyID = ds.Tables(0).Rows(0)("TCCurrency").ToString()
            Obj.objPurInvMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()
            Obj.objPurInvMain.str_MiscText = ds.Tables(0).Rows(0)("MiscText").ToString()
            Obj.objPurInvMain.str_DiscText = ds.Tables(0).Rows(0)("DiscText").ToString()
            Obj.objPurInvMain.bool_TaxFileReturn = ds.Tables(0).Rows(0)("TaxReturnFiled")
            Obj.objPurInvMain.str_PermitNo = ds.Tables(0).Rows(0)("PermitNo").ToString()

            If ds.Tables(1).Rows.Count > 0 Then
                Obj.objProject.str_ProjectID = ds.Tables(1).Rows(0)("ProjectID").ToString()
                Obj.objProject.str_ProjectLocation = ds.Tables(1).Rows(0)("ProjectLocation").ToString()
                Obj.objProject.str_WorkOrderNo = ds.Tables(1).Rows(0)("WorkOrderNo").ToString()
            End If

            If ds.Tables(2).Rows.Count > 0 Then
                Obj.objPurInvMain.dt_InvoiceAccounts = ds.Tables(2)
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try


    End Sub

    Public Function Update_CP(ByVal _strPath As String, ByVal _strPwd As String, ByVal obj As csPurchaseInvoice, ByRef CPNo As String, ByRef intRevNo As Integer,
                             ByRef _OutSMS As String, ByRef _OutEmail As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("InvoiceAccountUpdate", BaseConn.cnn)

            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.int_CID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.objPurInvMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.objPurInvMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.objPurInvMain.str_Prefix)

            BaseConn.cmd.Parameters.AddWithValue("@PIPNo", obj.objPurInvMain.str_PurInvNo)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.objPurInvMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@DoNo", obj.objPurInvMain.str_InvNo)
            BaseConn.cmd.Parameters.AddWithValue("@LpoNo", obj.objPurInvMain.str_InvRef)

            BaseConn.cmd.Parameters.AddWithValue("@SrcLedgerID", obj.objPurInvMain.str_SrcLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Alias", obj.objPurInvMain.str_Alias)
            BaseConn.cmd.Parameters.AddWithValue("@InvDate", obj.objPurInvMain.dtp_InvDate)
            BaseConn.cmd.Parameters.AddWithValue("@DueDate", obj.objPurInvMain.dtp_DueDate)
            BaseConn.cmd.Parameters.AddWithValue("@Aging", obj.objPurInvMain.int_Aging)
            BaseConn.cmd.Parameters.AddWithValue("@PaymentTerm", obj.objPurInvMain.str_PayTerm)

            BaseConn.cmd.Parameters.AddWithValue("@InvoiceType", obj.objPurInvMain.str_InvoiceType)
            BaseConn.cmd.Parameters.AddWithValue("@CounterPurchase", obj.objPurInvMain.bool_PurchaseInvoice)
            BaseConn.cmd.Parameters.AddWithValue("@IsCashPurchase", obj.objPurInvMain.bool_IsCashPurchase)
            BaseConn.cmd.Parameters.AddWithValue("@AffectInventory", obj.objPurInvMain.bool_AffectInventory)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.objPurInvMain.str_Comment)

            BaseConn.cmd.Parameters.AddWithValue("@TCAmount", obj.objPurInvMain.dbl_TCAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDisAmount", obj.objPurInvMain.dbl_TCDisAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDiscountAmount", obj.objPurInvMain.dbl_TCDiscountAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCMiscPercentage", obj.objPurInvMain.dbl_TCMiscPercentage)
            BaseConn.cmd.Parameters.AddWithValue("@TCMiscAmount", obj.objPurInvMain.dbl_TCMiscAmount)
            'BaseConn.cmd.Parameters.AddWithValue("@TCTaxAmount", obj.objPurInvMain.dbl_TCTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCAdjAmount", obj.objPurInvMain.dbl_TCAdjAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCNetAmount", obj.objPurInvMain.dbl_TCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@LCNetAmount", obj.objPurInvMain.dbl_LCNetAmount)

            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", obj.objPurInvMain.str_CurrencyID)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", obj.objPurInvMain.dbl_ExchangeRate)


            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.bool_ApprovedStatus)

            BaseConn.cmd.Parameters.AddWithValue("@ApprovedLevel", obj.ApprovedLevel)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedComment", obj.ApprovedComment)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedHigherLevel", obj.ApprovedHigherLevel)

            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.objPurInvMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@StatusCancel", obj.objPurInvMain.int_StatusCancel)
            BaseConn.cmd.Parameters.AddWithValue("@MiscText", obj.objPurInvMain.str_MiscText)
            BaseConn.cmd.Parameters.AddWithValue("@DiscText", obj.objPurInvMain.str_DiscText)
            BaseConn.cmd.Parameters.AddWithValue("@UserComment", obj.objPurInvMain.str_UserComment)

            BaseConn.cmd.Parameters.AddWithValue("@CashorCredit", obj.objPurInvMain.str_CashorCredit)
            BaseConn.cmd.Parameters.AddWithValue("@CashLedger", obj.objPurInvMain.int_CashLedger)
            BaseConn.cmd.Parameters.AddWithValue("@CashTendered", obj.objPurInvMain.dbl_CashTendered)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectID", obj.objProject.str_ProjectID)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectLocation", obj.objProject.str_ProjectLocation)
            BaseConn.cmd.Parameters.AddWithValue("@WorkOrderNo", obj.objProject.str_WorkOrderNo)

            BaseConn.cmd.Parameters.AddWithValue("@PermitNo", obj.objPurInvMain.str_PermitNo)

            BaseConn.cmd.Parameters.AddWithValue("@InvoiceTaxCode", obj.objPurInvMain.str_InvoiceTaxCode)
            BaseConn.cmd.Parameters.AddWithValue("@TCItemTaxAmount", obj.objPurInvMain.dbl_TCItemTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCInvoiceTaxAmount", obj.objPurInvMain.dbl_TCInvoiceTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@InvoiceTaxXML", obj.objPurInvMain.str_InvoiceTaxXML)
            BaseConn.cmd.Parameters.AddWithValue("@LanguageCode", obj.objPurInvMain.int_LanguageCode)

            BaseConn.cmd.Parameters.AddWithValue("@MatchingDT", obj.objPurInvSub.dt_PurInvMatching)
            BaseConn.cmd.Parameters.AddWithValue("@MRVDT", obj.objPurInvMain.dt_MRVNo4PIP)
            BaseConn.cmd.Parameters.AddWithValue("@InvTaxAmountDT", obj.objPurInvMain.dt_TaxItemDetails)

            BaseConn.cmd.Parameters.AddWithValue("@InvAccDetDT", obj.objPurInvMain.dt_InvoiceAccounts)

            BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output

            BaseConn.cmd.Parameters.Add("@OutSMSMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutEmailMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output

            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            CPNo = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            _OutSMS = BaseConn.cmd.Parameters("@OutSMSMsgText").Value.ToString
            _OutEmail = BaseConn.cmd.Parameters("@OutEmailMsgText").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.int_CID)
            ObjDalGeneral.Elog_Insert(obj.int_CID, _strPath, _strPwd, obj.objPurInvMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "PIP", Err.Number, "Error in " & obj.objPurInvMain.str_Flag & "ED : " & obj.objPurInvMain.str_PurInvNo & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_CP = _ErrString
    End Function



    Public Sub ImportPIPfromExcel(ByVal _strPath As String, ByVal _strPwd As String, ByVal _SiteID As String, ByVal _BSID As Integer, _
                                       ByVal _JVLedgerID As Integer, ByVal _PIPMainDT As DataTable, _
                              ByVal _CreatedBy As String, ByRef ErrNo As Integer, ByRef _ErrDesc As String)
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("ImportPIPfromExcel", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BSID)
            BaseConn.cmd.Parameters.AddWithValue("@JVLedgerID", _JVLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@PIPMainDT", _PIPMainDT)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", _CreatedBy)

            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 2000
            BaseConn.cmd.ExecuteNonQuery()

            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            _ErrDesc = _ErrString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(_SiteID)
            ObjDalGeneral.Elog_Insert(_SiteID, _strPath, _strPwd, _BSID, _CreatedBy, Date.Now, "", "PIP", Err.Number, "Error in Import from Excel :", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

    End Sub
End Class
