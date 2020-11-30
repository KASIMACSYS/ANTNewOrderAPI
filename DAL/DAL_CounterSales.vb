'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes

Public Class DAL_CounterSales
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByVal _DBPath As String, ByVal _DBPwd As String, ByRef Obj As csSalesInvoice, ByRef ErrNo As Integer, ByRef ErrMsg As String)
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetCounterSalesDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CSNo", Obj.objSalInvMain.str_SalInvNo)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.objSalInvMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.objSalInvMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@RevisionHistoryNo", Obj.objSalInvMain.int_RevisionHistoryNo)
            BaseConn.cmd.Parameters.Add("@CashLedger", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@CashTendered", SqlDbType.Float).Direction = ParameterDirection.Output

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            If Obj.objSalInvMain.str_Flag = "CS" Then
                Obj.objSalInvMain.int_CashLedger = BaseConn.cmd.Parameters("@CashLedger").Value
                Obj.objSalInvMain.dbl_CashTendered = BaseConn.cmd.Parameters("@CashTendered").Value

                Obj.objSalInvMain.str_SrcLedgerID = ds.Tables(0).Rows(0)("SrcLedgerID").ToString()
                Obj.objSalInvMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
                Obj.objSalInvMain.dtp_InvDate = ds.Tables(0).Rows(0)("InvDate").ToString()
                Obj.objSalInvMain.dtp_DueDate = ds.Tables(0).Rows(0)("DueDate").ToString()
                Obj.objSalInvMain.int_Aging = ds.Tables(0).Rows(0)("Aging").ToString()
                Obj.objSalInvMain.str_PayTerm = ds.Tables(0).Rows(0)("PaymentTerm").ToString()
                Obj.objSalInvMain.dbl_TCAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
                Obj.objSalInvMain.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
                Obj.objSalInvMain.dbl_TCMiscAmount = ds.Tables(0).Rows(0)("TCMiscAmount").ToString()
                Obj.objSalInvMain.dbl_TCMiscPercentage = ds.Tables(0).Rows(0)("TCMiscPercentage").ToString()
                'Obj.objSalInvMain.dbl_TCTaxAmount = ds.Tables(0).Rows(0)("TCTaxAmount").ToString()
                Obj.objSalInvMain.dbl_TCItemTaxAmount = ds.Tables(0).Rows(0)("TCItemTaxAmount")
                Obj.objSalInvMain.dbl_TCInvoiceTaxAmount = ds.Tables(0).Rows(0)("TCInvTaxAmount")
                Obj.objSalInvMain.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
                Obj.objSalInvMain.dbl_TCAdjAmount = ds.Tables(0).Rows(0)("TCAdjAmount").ToString()
                Obj.objSalInvMain.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
                Obj.objSalInvMain.dbl_TCPDCAmount = ds.Tables(0).Rows(0)("TCPDCAmount").ToString()
                Obj.objSalInvMain.bool_SalesInvoice = ds.Tables(0).Rows(0)("CounterSales").ToString()
                Obj.objSalInvMain.bool_IsCashSales = ds.Tables(0).Rows(0)("IsCashSales").ToString()
                Obj.objSalInvMain.bool_AffectInventory = ds.Tables(0).Rows(0)("AffectInventory").ToString()
                Obj.objSalInvMain.str_PaymentStatus = ds.Tables(0).Rows(0)("PaymentStatus").ToString()
                Obj.objSalInvMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
                Obj.objSalInvMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
                Obj.objSalInvMain.str_DeliveryAddress = ds.Tables(0).Rows(0)("DeliveryAddress").ToString()
                Obj.objSalInvMain.str_SalesManID = ds.Tables(0).Rows(0)("SalesManID").ToString()
                Obj.objSalInvMain.str_InvoiceType = ds.Tables(0).Rows(0)("InvoiceType").ToString()

                Obj.objSalInvMain.str_LpoNo = ds.Tables(0).Rows(0)("LpoNo").ToString()
                Obj.objSalInvMain.str_DONo = ds.Tables(0).Rows(0)("DoNo").ToString()
                Obj.objSalInvMain.str_SalOrd = ds.Tables(0).Rows(0)("SalOrd").ToString()
                Obj.objSalInvMain.str_InvRef = ds.Tables(0).Rows(0)("InvRef").ToString()
                Obj.objSalInvMain.dbl_LCNetCostPrice = ds.Tables(0).Rows(0)("LCNetCostAmount").ToString()
                Obj.objSalInvMain.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString()
                Obj.objSalInvMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()
                Obj.objSalInvMain.int_StatusCancel = ds.Tables(0).Rows(0)("StatusCancel").ToString()
                Obj.objSalInvMain.str_InvoiceStatus = ds.Tables(0).Rows(0)("SISStatus").ToString()
                Obj.objSalInvMain.str_WHID = ds.Tables(0).Rows(0)("WHID").ToString()

                Obj.objSalInvMain.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
                Obj.objSalInvMain.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
                Obj.objSalInvMain.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
                Obj.objSalInvMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
                Obj.objSalInvMain.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
                Obj.objSalInvMain.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
                Obj.objSalInvMain.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
                Obj.objSalInvMain.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()

                Obj.objSalInvMain.str_ItemTaxCode = ds.Tables(0).Rows(0)("ItemTaxCode")
                Obj.objSalInvMain.str_InvoiceTaxCode = ds.Tables(0).Rows(0)("InvoiceTaxCode")
                Obj.objSalInvMain.dbl_ItemDiscPercentage = ds.Tables(0).Rows(0)("ItemDiscPercentage")
                Obj.objSalInvMain.str_InvoiceTaxXML = ds.Tables(0).Rows(0)("InvoiceTaxDetails")


                If ds.Tables(0).Rows(0)("RetentionDueDate").ToString = "" Then
                    Obj.objRetention.dtp_RetDueDate = Date.Now
                Else
                    Obj.objRetention.dtp_RetDueDate = ds.Tables(0).Rows(0)("RetentionDueDate").ToString()
                End If


                Obj.objRetention.dbl_RetAmtDeduction = ds.Tables(0).Rows(0)("RetentionAmountDeduction").ToString()
                Obj.objRetention.dbl_RetAmtAddition = ds.Tables(0).Rows(0)("RetentionAmountAddition").ToString()

                Obj.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
                Obj.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
                Obj.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
                Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
                Obj.objSalInvMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()

                Obj.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
                Obj.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
                Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus").ToString()

                Obj.objSalInvMain.str_CurrencyID = ds.Tables(0).Rows(0)("TCCurrency").ToString()
                Obj.objSalInvMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()
                Obj.objSalInvMain.dbl_LCNetCostPrice = ds.Tables(0).Rows(0)("LCNetAmount").ToString() 'TODO
                Obj.objSalInvMain.str_MiscText = ds.Tables(0).Rows(0)("MiscText").ToString()
                Obj.objSalInvMain.str_DiscText = ds.Tables(0).Rows(0)("DiscText").ToString()
                Obj.objSalInvMain.str_UserComment = ds.Tables(0).Rows(0)("UserComment").ToString()
                Obj.objSalInvMain.str_ApproverComment = ds.Tables(0).Rows(0)("ApproverComment").ToString()
                Obj.objSalInvMain.bool_TaxFileReturn = ds.Tables(0).Rows(0)("TaxReturnFiled")
                Obj.objSalInvMain.str_Country = ds.Tables(0).Rows(0)("Country").ToString()
                Obj.objSalInvMain.str_Filter3 = ds.Tables(0).Rows(0)("Filter3").ToString()
                Obj.objSalInvMain.str_Filter4 = ds.Tables(0).Rows(0)("Filter4").ToString()
                Obj.objSalInvMain.int_LanguageCode = ds.Tables(0).Rows(0)("LanguageCode").ToString()
                If ds.Tables(1).Rows.Count > 0 Then
                    Obj.objSalInvSub.dt_SalInv = ds.Tables(1)
                End If

                If ds.Tables(2).Rows.Count > 0 Then
                    Obj.objProject.str_ProjectID = ds.Tables(2).Rows(0)("ProjectID").ToString()
                    Obj.objProject.str_ProjectLocation = ds.Tables(2).Rows(0)("ProjectLocation").ToString()
                    Obj.objProject.str_WorkOrderNo = ds.Tables(2).Rows(0)("WorkOrderNo").ToString()
                End If

                If ds.Tables(3).Rows.Count > 0 Then
                    Obj.objSalInvMain.dt_InvoiceAccounts = ds.Tables(3)
                End If
                Obj.objSalInvMain.dt_TaxItemDetails = ds.Tables(4)

                Obj.DTBatch = ds.Tables(5)

                If ds.Tables(6).Rows.Count > 0 Then
                    Obj.objSalInvMain.str_RTF_Description = ds.Tables(6).Rows(0)("Description").ToString()
                Else
                    Obj.objSalInvMain.str_RTF_Description = ""
                End If

                Obj.DTItemExtraDetails = ds.Tables(7)

            ElseIf Obj.objSalInvMain.str_Flag = "EZERP" Then
                Obj.objSalInvMain.str_SrcLedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
                Obj.objSalInvMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
                Obj.objSalInvMain.dtp_DueDate = ds.Tables(0).Rows(0)("VouDate").ToString()
                Obj.objSalInvMain.dtp_InvDate = Date.Now
                Obj.objRetention.dtp_RetDueDate = Date.Now
                Obj.objSalInvMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()
                Obj.objSalInvMain.str_CurrencyID = ds.Tables(0).Rows(0)("TCCurrency").ToString()


                If ds.Tables(1).Rows.Count > 0 Then
                    Obj.objSalInvSub.dt_SalInv = ds.Tables(1)
                End If

            Else
                Obj.objSalInvMain.str_SrcLedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
                Obj.objSalInvMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
                Obj.objSalInvMain.dtp_InvDate = Date.Now
                Obj.objSalInvMain.dtp_DueDate = ds.Tables(0).Rows(0)("SODate").ToString()
                Obj.objSalInvMain.int_Aging = ds.Tables(0).Rows(0)("Aging").ToString()
                Obj.objSalInvMain.str_PayTerm = ds.Tables(0).Rows(0)("PayTerm").ToString()
                Obj.objSalInvMain.dbl_TCAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
                Obj.objSalInvMain.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
                Obj.objSalInvMain.dbl_TCMiscAmount = ds.Tables(0).Rows(0)("TCMiscAmount").ToString()
                Obj.objSalInvMain.dbl_TCMiscPercentage = ds.Tables(0).Rows(0)("TCMiscPercentage").ToString()
                'Obj.objSalInvMain.dbl_TCTaxAmount = ds.Tables(0).Rows(0)("TCTaxAmount").ToString()
                Obj.objSalInvMain.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
                Obj.objSalInvMain.dbl_TCAdjAmount = ds.Tables(0).Rows(0)("TCAdjAmount").ToString()
                Obj.objSalInvMain.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
                Obj.objSalInvMain.dbl_TCPDCAmount = 0 ' ds.Tables(0).Rows(0)("TCPDCAmount").ToString()
                Obj.objSalInvMain.bool_SalesInvoice = 0 'ds.Tables(0).Rows(0)("CounterSales").ToString()
                Obj.objSalInvMain.str_PaymentStatus = "" ' ds.Tables(0).Rows(0)("PaymentStatus").ToString()
                Obj.objSalInvMain.int_RevNo = 0 ' ds.Tables(0).Rows(0)("RevNo").ToString()
                Obj.objSalInvMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
                Obj.objSalInvMain.str_SalesManID = ds.Tables(0).Rows(0)("SalesManID").ToString()
                Obj.objSalInvMain.str_DeliveryAddress = ds.Tables(0).Rows(0)("DeliveryAddress").ToString()
                Obj.objSalInvMain.str_LpoNo = "" ' ds.Tables(0).Rows(0)("LpoNo").ToString()
                Obj.objSalInvMain.str_DONo = "" ' ds.Tables(0).Rows(0)("DoNo").ToString()
                Obj.objSalInvMain.str_SalOrd = ds.Tables(0).Rows(0)("SalOrd").ToString()
                Obj.objSalInvMain.dbl_LCNetCostPrice = 0 'ds.Tables(0).Rows(0)("LCNetCostAmount").ToString()
                Obj.objSalInvMain.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString()
                Obj.objSalInvMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()
                Obj.objSalInvMain.str_MiscText = ds.Tables(0).Rows(0)("MiscText").ToString()
                Obj.objSalInvMain.str_DiscText = ds.Tables(0).Rows(0)("DiscText").ToString()
                Obj.objSalInvMain.str_WHID = 101 ' ds.Tables(0).Rows(0)("WHID").ToString()

                Obj.objSalInvMain.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
                Obj.objSalInvMain.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
                Obj.objSalInvMain.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
                Obj.objSalInvMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
                Obj.objSalInvMain.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
                Obj.objSalInvMain.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
                Obj.objSalInvMain.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
                Obj.objSalInvMain.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()

                Obj.objSalInvMain.str_ItemTaxCode = ds.Tables(0).Rows(0)("ItemTaxCode")
                Obj.objSalInvMain.str_InvoiceTaxCode = ds.Tables(0).Rows(0)("InvoiceTaxCode")
                Obj.objSalInvMain.dbl_TCInvoiceTaxAmount = ds.Tables(0).Rows(0)("TCInvTaxAmount")
                Obj.objSalInvMain.dbl_ItemDiscPercentage = ds.Tables(0).Rows(0)("ItemDiscPercentage")
                Obj.objSalInvMain.str_InvoiceTaxXML = ds.Tables(0).Rows(0)("InvoiceTaxDetails")

                Obj.objRetention.dtp_RetDueDate = Date.Now
                Obj.objRetention.dbl_RetAmtDeduction = 0
                Obj.objRetention.dbl_RetAmtAddition = 0

                Obj.objSalInvMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()

                Obj.objSalInvMain.str_CurrencyID = ds.Tables(0).Rows(0)("TCCurrency").ToString()
                Obj.objSalInvMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()
                Obj.objSalInvMain.dbl_LCNetCostPrice = ds.Tables(0).Rows(0)("LCNetAmount").ToString() 'TODO

                Obj.objSalInvMain.str_UserComment = ds.Tables(0).Rows(0)("UserComment").ToString()
                Obj.objSalInvMain.str_ApproverComment = ds.Tables(0).Rows(0)("ApproverComment").ToString()
                Obj.objSalInvMain.int_LanguageCode = ds.Tables(0).Rows(0)("LanguageCode").ToString()

                If ds.Tables(1).Rows.Count > 0 Then
                    Obj.objSalInvSub.dt_SalInv = ds.Tables(1)
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
                Obj.objSalInvMain.dt_TaxItemDetails = ds.Tables(3)

                If ds.Tables(4).Rows.Count > 0 Then
                    Obj.objSalInvMain.str_RTF_Description = ds.Tables(4).Rows(0)("Description").ToString()
                Else
                    Obj.objSalInvMain.str_RTF_Description = ""
                End If

                Obj.DTItemExtraDetails = ds.Tables(5)
            End If
        Catch ex As Exception
            ErrNo = 1
            ErrMsg = ex.Message ' "Problem in Updating Invoice"
        Finally
            BaseConn.Close()
        End Try

    End Sub

    Public Function Update_CS(ByVal _strPath As String, ByVal _strPwd As String, ByRef CSNo As String, ByRef intRevNo As Integer, ByVal obj As csSalesInvoice, ByRef _OutSMS As String, ByRef _OutEmail As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strPwd)
            If obj.objSalInvMain.str_MenuID = "Menu_325" Then
                BaseConn.cmd = New SqlClient.SqlCommand("sp_EZERP_InvoiceAccountUpdate_CS", BaseConn.cnn)
                BaseConn.cmd.Parameters.AddWithValue("@PresNo", obj.objSalInvMain.str_InvRef)
            Else
                BaseConn.cmd = New SqlClient.SqlCommand("InvoiceAccountUpdate_CS", BaseConn.cnn)
            End If
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.int_CID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.objSalInvMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.objSalInvMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.objSalInvMain.str_Prefix)

            BaseConn.cmd.Parameters.AddWithValue("@SISNo", obj.objSalInvMain.str_SalInvNo)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.objSalInvMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@DoNo", obj.objSalInvMain.str_DONo)
            BaseConn.cmd.Parameters.AddWithValue("@LpoNo", obj.objSalInvMain.str_LpoNo)
            BaseConn.cmd.Parameters.AddWithValue("@SalOrd", obj.objSalInvMain.str_SalOrd)
            BaseConn.cmd.Parameters.AddWithValue("@SrcLedgerID", obj.objSalInvMain.str_SrcLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Alias", obj.objSalInvMain.str_Alias)
            BaseConn.cmd.Parameters.AddWithValue("@InvDate", obj.objSalInvMain.dtp_InvDate)
            BaseConn.cmd.Parameters.AddWithValue("@DueDate", obj.objSalInvMain.dtp_DueDate)
            BaseConn.cmd.Parameters.AddWithValue("@Aging", obj.objSalInvMain.int_Aging)
            BaseConn.cmd.Parameters.AddWithValue("@PaymentTerm", obj.objSalInvMain.str_PayTerm)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManID", obj.objSalInvMain.str_SalesManID)
            BaseConn.cmd.Parameters.AddWithValue("@InvoiceType", obj.objSalInvMain.str_InvoiceType)
            BaseConn.cmd.Parameters.AddWithValue("@CounterSales", obj.objSalInvMain.bool_SalesInvoice)
            BaseConn.cmd.Parameters.AddWithValue("@IsCashSales", obj.objSalInvMain.bool_IsCashSales)
            BaseConn.cmd.Parameters.AddWithValue("@AffectInventory", obj.objSalInvMain.bool_AffectInventory)
            BaseConn.cmd.Parameters.AddWithValue("@SISStatus", obj.objSalInvMain.str_InvoiceStatus)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.objSalInvMain.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@TCAmount", obj.objSalInvMain.dbl_TCAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDisAmount", obj.objSalInvMain.dbl_TCDisAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDiscountAmount", obj.objSalInvMain.dbl_TCDiscountAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCMiscAmount", obj.objSalInvMain.dbl_TCMiscAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCMiscPercentage", obj.objSalInvMain.dbl_TCMiscPercentage)
            'BaseConn.cmd.Parameters.AddWithValue("@TCTaxAmount", obj.objSalInvMain.dbl_TCInvoiceTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCNetAmount", obj.objSalInvMain.dbl_TCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@LCNetAmount", obj.objSalInvMain.dbl_LCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", obj.objSalInvMain.str_CurrencyID)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", obj.objSalInvMain.dbl_ExchangeRate)
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
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.objSalInvMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CashorCredit", obj.objSalInvMain.str_CashorCredit)
            BaseConn.cmd.Parameters.AddWithValue("@CashLedger", obj.objSalInvMain.int_CashLedger)
            BaseConn.cmd.Parameters.AddWithValue("@CashTendered", obj.objSalInvMain.dbl_CashTendered)
            BaseConn.cmd.Parameters.AddWithValue("@MiscText", obj.objSalInvMain.str_MiscText)
            BaseConn.cmd.Parameters.AddWithValue("@DiscText", obj.objSalInvMain.str_DiscText)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectID", obj.objProject.str_ProjectID)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectLocation", obj.objProject.str_ProjectLocation)
            BaseConn.cmd.Parameters.AddWithValue("@WorkOrderNo", obj.objProject.str_WorkOrderNo)
            BaseConn.cmd.Parameters.AddWithValue("@WHID", obj.objSalInvMain.str_WHID)
            BaseConn.cmd.Parameters.AddWithValue("@StatusCancel", obj.objSalInvMain.int_StatusCancel)
            ''AM Specific
            BaseConn.cmd.Parameters.AddWithValue("@Desc1", obj.objSalInvMain.str_Desc1)
            BaseConn.cmd.Parameters.AddWithValue("@Desc2", obj.objSalInvMain.str_Desc2)
            BaseConn.cmd.Parameters.AddWithValue("@Desc3", obj.objSalInvMain.str_Desc3)
            BaseConn.cmd.Parameters.AddWithValue("@Desc4", obj.objSalInvMain.str_Desc4)
            BaseConn.cmd.Parameters.AddWithValue("@Desc5", obj.objSalInvMain.str_Desc5)
            BaseConn.cmd.Parameters.AddWithValue("@Desc6", obj.objSalInvMain.str_Desc6)
            BaseConn.cmd.Parameters.AddWithValue("@Desc7", obj.objSalInvMain.str_Desc7)
            BaseConn.cmd.Parameters.AddWithValue("@Desc8", obj.objSalInvMain.str_Desc8)
            BaseConn.cmd.Parameters.AddWithValue("@DeliveryAddress", obj.objSalInvMain.str_DeliveryAddress)
            BaseConn.cmd.Parameters.AddWithValue("@UserComment", IIf(obj.str_UserComment = Nothing, "", obj.str_UserComment))
            BaseConn.cmd.Parameters.AddWithValue("@ApproverComment", obj.str_ApprovedComment)
            BaseConn.cmd.Parameters.AddWithValue("@ItemDiscPercentage", obj.objSalInvMain.dbl_ItemDiscPercentage)
            BaseConn.cmd.Parameters.AddWithValue("@Country", obj.objSalInvMain.str_Country)
            BaseConn.cmd.Parameters.AddWithValue("@Filter3", obj.objSalInvMain.str_Filter3)
            BaseConn.cmd.Parameters.AddWithValue("@Filter4", obj.objSalInvMain.str_Filter4)
            BaseConn.cmd.Parameters.AddWithValue("@RetentionDueDate", obj.objRetention.dtp_RetDueDate)
            BaseConn.cmd.Parameters.AddWithValue("@RetAmtDeduction", obj.objRetention.dbl_RetAmtDeduction)
            BaseConn.cmd.Parameters.AddWithValue("@RetAmtAddition", obj.objRetention.dbl_RetAmtAddition)
            BaseConn.cmd.Parameters.AddWithValue("@ItemTaxCode", obj.objSalInvMain.str_ItemTaxCode)
            BaseConn.cmd.Parameters.AddWithValue("@InvoiceTaxCode", obj.objSalInvMain.str_InvoiceTaxCode)
            BaseConn.cmd.Parameters.AddWithValue("@TCItemTaxAmount", obj.objSalInvMain.dbl_TCItemTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCInvoiceTaxAmount", obj.objSalInvMain.dbl_TCInvoiceTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@InvoiceTaxXML", obj.objSalInvMain.str_InvoiceTaxXML)
            BaseConn.cmd.Parameters.AddWithValue("@LanguageCode", obj.objSalInvMain.int_LanguageCode)


            BaseConn.cmd.Parameters.AddWithValue("@TCAdjAmount", obj.objSalInvMain.dbl_TCAdjAmount)


            If obj.objSalInvMain.str_Flag <> "Approve" Then
                If obj.objRetention.dt_RetentionInvoiceList.Columns.Contains("InvDate") Then
                    obj.objRetention.dt_RetentionInvoiceList.Columns.Remove("InvDate")
                End If
                If obj.objRetention.dt_RetentionInvoiceList.Columns.Contains("Comment") Then
                    obj.objRetention.dt_RetentionInvoiceList.Columns.Remove("Comment")
                End If
                BaseConn.cmd.Parameters.AddWithValue("@RetenMatchingDT", obj.objRetention.dt_RetentionInvoiceList)
            End If

            'If obj.objSalInvMain.str_MenuID = "Menu_317" Or obj.objSalInvMain.str_MenuID = "Menu_317_1" Then
            BaseConn.cmd.Parameters.AddWithValue("@InvAccDetDT", obj.objSalInvMain.dt_InvoiceAccounts)
            'End If

            BaseConn.cmd.Parameters.AddWithValue("@CounterSalesDT", obj.objSalInvSub.dt_SalInv)
            BaseConn.cmd.Parameters.AddWithValue("@MatchingDT", obj.objSalInvSub.dt_SalInvMatching)
            BaseConn.cmd.Parameters.AddWithValue("@ItemBatchDT", obj.DTBatch)
            BaseConn.cmd.Parameters.AddWithValue("@InvTaxAmountDT", obj.objSalInvMain.dt_TaxItemDetails)

            BaseConn.cmd.Parameters.AddWithValue("@RTF_Description", obj.objSalInvMain.str_RTF_Description)
            BaseConn.cmd.Parameters.AddWithValue("@VoucherItemExtraDetailsDT", obj.DTItemExtraDetails)

            BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutSMSMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutEmailMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output

            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            CSNo = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            _OutSMS = BaseConn.cmd.Parameters("@OutSMSMsgText").Value.ToString
            _OutEmail = BaseConn.cmd.Parameters("@OutEmailMsgText").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.int_CID)
            ObjDalGeneral.Elog_Insert(obj.int_CID, _strPath, _strPwd, obj.objSalInvMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "CS", ErrNo, "Error in " & obj.objSalInvMain.str_Flag & " : " & obj.objSalInvMain.str_SalInvNo & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_CS = _ErrString
    End Function
End Class


Public Class DAL_SIS
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub GetDO4SIS(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _StrSiteID As String, ByVal _BusinessPeriodID As Integer, ByVal _Flag As String, _
                         ByVal _VouNo As String, ByVal _LedgerID As String, ByVal _strTCCurency As String, ByRef DTUnInvoiced As DataTable, _
                         ByRef DTInvoicedDO As DataTable, ByVal _GrpID As Integer, ByVal _ExchangeRate As Double)
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetDO4SIS]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _StrSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", _VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", _strTCCurency)
            BaseConn.cmd.Parameters.AddWithValue("@GrpID", _GrpID)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", _ExchangeRate)

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

    Public Sub Get_Structure(ByVal _DBPath As String, ByVal _DBPwd As String, ByRef Obj As csSalesInvoice)
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetSISDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@CSNo", Obj.objSalInvMain.str_SalInvNo)
            BaseConn.cmd.Parameters.AddWithValue("@DONO", Obj.objSalInvMain.str_DONo)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.objSalInvMain.str_Flag)

            BaseConn.cmd.Parameters.Add("@CashLedger", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@CashTendered", SqlDbType.Decimal).Direction = ParameterDirection.Output

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            If Obj.objSalInvMain.str_Flag <> "DO" Then
                Obj.objSalInvMain.int_CashLedger = BaseConn.cmd.Parameters("@CashLedger").Value
                Obj.objSalInvMain.dbl_CashTendered = BaseConn.cmd.Parameters("@CashTendered").Value

                Obj.objSalInvMain.str_SrcLedgerID = ds.Tables(0).Rows(0)("SrcLedgerID").ToString()
                Obj.objSalInvMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
                Obj.objSalInvMain.dtp_InvDate = ds.Tables(0).Rows(0)("InvDate").ToString()
                Obj.objSalInvMain.dtp_DueDate = ds.Tables(0).Rows(0)("DueDate").ToString()
                Obj.objSalInvMain.int_Aging = ds.Tables(0).Rows(0)("Aging").ToString()
                Obj.objSalInvMain.str_PayTerm = ds.Tables(0).Rows(0)("PaymentTerm").ToString()
                Obj.objSalInvMain.dbl_TCAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
                Obj.objSalInvMain.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
                Obj.objSalInvMain.dbl_TCMiscPercentage = ds.Tables(0).Rows(0)("TCMiscPercentage").ToString()
                Obj.objSalInvMain.dbl_TCMiscAmount = ds.Tables(0).Rows(0)("TCMiscAmount").ToString()
                Obj.objSalInvMain.dbl_TCItemTaxAmount = ds.Tables(0).Rows(0)("TCItemTaxAmount")
                Obj.objSalInvMain.dbl_TCInvoiceTaxAmount = ds.Tables(0).Rows(0)("TCInvTaxAmount")
                Obj.objSalInvMain.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
                Obj.objSalInvMain.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
                Obj.objSalInvMain.dbl_TCAdjAmount = ds.Tables(0).Rows(0)("TCAdjAmount").ToString()
                Obj.objSalInvMain.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
                Obj.objSalInvMain.dbl_TCPDCAmount = ds.Tables(0).Rows(0)("TCPDCAmount").ToString()
                Obj.objSalInvMain.bool_SalesInvoice = ds.Tables(0).Rows(0)("CounterSales").ToString()
                Obj.objSalInvMain.bool_IsCashSales = ds.Tables(0).Rows(0)("IsCashSales").ToString()
                Obj.objSalInvMain.bool_AffectInventory = ds.Tables(0).Rows(0)("AffectInventory").ToString()
                Obj.objSalInvMain.str_PaymentStatus = ds.Tables(0).Rows(0)("PaymentStatus").ToString()
                Obj.objSalInvMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
                Obj.objSalInvMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
                Obj.objSalInvMain.str_SalesManID = ds.Tables(0).Rows(0)("SalesManID").ToString()
                Obj.objSalInvMain.str_InvoiceType = ds.Tables(0).Rows(0)("InvoiceType").ToString()
                Obj.objSalInvMain.int_StatusCancel = ds.Tables(0).Rows(0)("StatusCancel").ToString()

                Obj.objSalInvMain.str_LpoNo = ds.Tables(0).Rows(0)("LpoNo").ToString()
                Obj.objSalInvMain.str_DONo = ds.Tables(0).Rows(0)("DoNo").ToString()

                Obj.objSalInvMain.dbl_LCNetCostPrice = ds.Tables(0).Rows(0)("LCNetAmount").ToString()
                'Obj.objSalInvMain.dbl_LCPDCAmount = ds.Tables(0).Rows(0)("LCPDCAmount").ToString()
                Obj.objSalInvMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()

                Obj.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
                Obj.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
                Obj.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
                Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
                Obj.objSalInvMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()

                Obj.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
                Obj.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
                Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus").ToString()

                Obj.objSalInvMain.str_CurrencyID = ds.Tables(0).Rows(0)("TCCurrency").ToString()
                Obj.objSalInvMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()
                Obj.objSalInvMain.dbl_LCNetCostPrice = ds.Tables(0).Rows(0)("LCNetAmount").ToString() 'TODO
                Obj.objSalInvMain.str_MiscText = ds.Tables(0).Rows(0)("MiscText").ToString() 'TODO
                Obj.objSalInvMain.str_DiscText = ds.Tables(0).Rows(0)("DiscText").ToString() 'TODO
                Obj.objSalInvMain.str_ContactPerson = ds.Tables(0).Rows(0)("ContactPerson").ToString()
                Obj.objSalInvMain.str_DeliveryAddress = ds.Tables(0).Rows(0)("DeliveryAddress").ToString()
                Obj.objSalInvMain.bool_TaxFileReturn = ds.Tables(0).Rows(0)("TaxReturnFiled")
                Obj.objSalInvMain.str_Country = ds.Tables(0).Rows(0)("Country").ToString()
                Obj.objSalInvMain.str_Filter3 = ds.Tables(0).Rows(0)("Filter3").ToString()
                Obj.objSalInvMain.str_Filter4 = ds.Tables(0).Rows(0)("Filter4").ToString()

                Obj.objSalInvMain.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
                Obj.objSalInvMain.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
                Obj.objSalInvMain.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
                Obj.objSalInvMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
                Obj.objSalInvMain.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
                Obj.objSalInvMain.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
                Obj.objSalInvMain.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
                Obj.objSalInvMain.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()
                Obj.objSalInvMain.str_Consignee = ds.Tables(0).Rows(0)("Consignee").ToString()

                Obj.objSalInvMain.str_ItemTaxCode = ds.Tables(0).Rows(0)("ItemTaxCode")
                Obj.objSalInvMain.str_InvoiceTaxCode = ds.Tables(0).Rows(0)("InvoiceTaxCode")
                Obj.objSalInvMain.str_InvoiceTaxXML = ds.Tables(0).Rows(0)("InvoiceTaxDetails")
            End If

            If ds.Tables(1).Rows.Count > 0 Then
                Obj.objProject.str_ProjectID = ds.Tables(1).Rows(0)("ProjectID").ToString()
                Obj.objProject.str_ProjectLocation = ds.Tables(1).Rows(0)("ProjectLocation").ToString()
                Obj.objProject.str_WorkOrderNo = ds.Tables(1).Rows(0)("WorkOrderNo").ToString()
            End If

            If ds.Tables(2).Rows.Count > 0 Then
                Obj.objSalInvMain.dt_InvoiceAccounts = ds.Tables(2)
            End If

            If Obj.objSalInvMain.str_Flag = "DO" Then
                If ds.Tables(3).Rows.Count > 0 Then
                    Obj.objSalInvMain.int_Aging = ds.Tables(3).Rows(0)("Aging").ToString()
                    Obj.objSalInvMain.str_PayTerm = ds.Tables(3).Rows(0)("PayTerm").ToString()
                    Obj.objSalInvMain.str_Comment = ds.Tables(3).Rows(0)("Comment").ToString()
                    Obj.objSalInvMain.str_DeliveryAddress = ds.Tables(3).Rows(0)("DeliveryAddress").ToString()
                    Obj.objSalInvMain.str_LpoNo = ds.Tables(3).Rows(0)("MerchantRef").ToString()
                    Obj.objSalInvMain.str_SalesManID = ds.Tables(3).Rows(0)("SalesManID").ToString()
                    Obj.objSalInvMain.str_ContactPerson = ds.Tables(3).Rows(0)("ContactPerson").ToString()
                    Obj.objProject.str_ProjectID = ds.Tables(3).Rows(0)("ProjectID").ToString()
                    Obj.objProject.str_ProjectLocation = ds.Tables(3).Rows(0)("ProjectLocation").ToString()
                    Obj.objProject.str_WorkOrderNo = ds.Tables(3).Rows(0)("WorkOrderNo").ToString()
                    Obj.objSalInvMain.str_MiscText = ds.Tables(3).Rows(0)("MiscText").ToString
                    Obj.objSalInvMain.str_DiscText = ds.Tables(3).Rows(0)("DiscText").ToString()
                    Obj.objSalInvMain.str_Consignee = ds.Tables(3).Rows(0)("Consignee").ToString()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try


    End Sub

    Public Function Update_CS(ByVal _strPath As String, ByVal _strPwd As String, ByVal obj As csSalesInvoice, ByRef CSNo As String, ByRef intRevNo As Integer,
                            ByRef _OutSMS As String, ByRef _OutEmail As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strPwd)

            BaseConn.cmd = New SqlClient.SqlCommand("InvoiceAccountUpdate_SIS", BaseConn.cnn)

            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.int_CID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.objSalInvMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.objSalInvMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.objSalInvMain.str_Prefix)

            BaseConn.cmd.Parameters.AddWithValue("@SISNo", obj.objSalInvMain.str_SalInvNo)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.objSalInvMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@DoNo", obj.objSalInvMain.str_DONo)
            BaseConn.cmd.Parameters.AddWithValue("@LpoNo", obj.objSalInvMain.str_LpoNo)
            BaseConn.cmd.Parameters.AddWithValue("@SrcLedgerID", obj.objSalInvMain.str_SrcLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Alias", obj.objSalInvMain.str_Alias)
            BaseConn.cmd.Parameters.AddWithValue("@InvDate", obj.objSalInvMain.dtp_InvDate)
            BaseConn.cmd.Parameters.AddWithValue("@DueDate", obj.objSalInvMain.dtp_DueDate)
            BaseConn.cmd.Parameters.AddWithValue("@Aging", obj.objSalInvMain.int_Aging)
            BaseConn.cmd.Parameters.AddWithValue("@PaymentTerm", obj.objSalInvMain.str_PayTerm)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManID", obj.objSalInvMain.str_SalesManID)
            BaseConn.cmd.Parameters.AddWithValue("@InvoiceType", obj.objSalInvMain.str_InvoiceType)
            BaseConn.cmd.Parameters.AddWithValue("@CounterSales", obj.objSalInvMain.bool_SalesInvoice)
            BaseConn.cmd.Parameters.AddWithValue("@IsCashSales", obj.objSalInvMain.bool_IsCashSales)
            BaseConn.cmd.Parameters.AddWithValue("@AffectInventory", obj.objSalInvMain.bool_AffectInventory)
            BaseConn.cmd.Parameters.AddWithValue("@StatusCancel", obj.objSalInvMain.int_StatusCancel)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.objSalInvMain.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@Desc1", obj.objSalInvMain.str_Desc1)
            BaseConn.cmd.Parameters.AddWithValue("@Desc2", obj.objSalInvMain.str_Desc2)
            BaseConn.cmd.Parameters.AddWithValue("@Desc3", obj.objSalInvMain.str_Desc3)
            BaseConn.cmd.Parameters.AddWithValue("@Desc4", obj.objSalInvMain.str_Desc4)
            BaseConn.cmd.Parameters.AddWithValue("@Desc5", obj.objSalInvMain.str_Desc5)
            BaseConn.cmd.Parameters.AddWithValue("@Desc6", obj.objSalInvMain.str_Desc6)
            BaseConn.cmd.Parameters.AddWithValue("@Desc7", obj.objSalInvMain.str_Desc7)
            BaseConn.cmd.Parameters.AddWithValue("@Desc8", obj.objSalInvMain.str_Desc8)

            BaseConn.cmd.Parameters.AddWithValue("@TCAmount", obj.objSalInvMain.dbl_TCAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDisAmount", obj.objSalInvMain.dbl_TCDisAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDiscountAmount", obj.objSalInvMain.dbl_TCDiscountAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCMiscPercentage", obj.objSalInvMain.dbl_TCMiscPercentage)
            BaseConn.cmd.Parameters.AddWithValue("@TCMiscAmount", obj.objSalInvMain.dbl_TCMiscAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCAdjAmount", obj.objSalInvMain.dbl_TCAdjAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCNetAmount", obj.objSalInvMain.dbl_TCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@LCNetAmount", obj.objSalInvMain.dbl_LCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@LCNetCostAmount", obj.objSalInvMain.dbl_LCNetCostPrice)
            BaseConn.cmd.Parameters.AddWithValue("@LCNetProfit", obj.objSalInvMain.dbl_LCNetProfit)
            BaseConn.cmd.Parameters.AddWithValue("@MiscText", obj.objSalInvMain.str_MiscText)
            BaseConn.cmd.Parameters.AddWithValue("@DiscText", obj.objSalInvMain.str_DiscText)
            BaseConn.cmd.Parameters.AddWithValue("@Consignee", obj.objSalInvMain.str_Consignee)
            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", obj.objSalInvMain.str_CurrencyID)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", obj.objSalInvMain.dbl_ExchangeRate)

            BaseConn.cmd.Parameters.AddWithValue("@InvoiceTaxCode", obj.objSalInvMain.str_InvoiceTaxCode)
            BaseConn.cmd.Parameters.AddWithValue("@TCItemTaxAmount", obj.objSalInvMain.dbl_TCItemTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCInvoiceTaxAmount", obj.objSalInvMain.dbl_TCInvoiceTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@InvoiceTaxXML", obj.objSalInvMain.str_InvoiceTaxXML)


            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.bool_ApprovedStatus)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.objSalInvMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CashorCredit", obj.objSalInvMain.str_CashorCredit)
            BaseConn.cmd.Parameters.AddWithValue("@CashLedger", obj.objSalInvMain.int_CashLedger)
            BaseConn.cmd.Parameters.AddWithValue("@CashTendered", obj.objSalInvMain.dbl_CashTendered)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectID", obj.objProject.str_ProjectID)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectLocation", obj.objProject.str_ProjectLocation)
            BaseConn.cmd.Parameters.AddWithValue("@WorkOrderNo", obj.objProject.str_WorkOrderNo)
            BaseConn.cmd.Parameters.AddWithValue("@ContactPerson", obj.objSalInvMain.str_ContactPerson)
            BaseConn.cmd.Parameters.AddWithValue("@DeliveryAddress", obj.objSalInvMain.str_DeliveryAddress)
            BaseConn.cmd.Parameters.AddWithValue("@MatchingDT", obj.objSalInvSub.dt_SalInvMatching)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedLevel", obj.ApprovedLevel)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedComment", obj.ApprovedComment)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedHigherLevel", obj.ApprovedHigherLevel)

            If obj.objSalInvMain.str_MenuID = "ERP_165" Then
                BaseConn.cmd.Parameters.AddWithValue("@InvAccDetDT", obj.objSalInvMain.dt_InvoiceAccounts)
            End If

            BaseConn.cmd.Parameters.AddWithValue("@Country", obj.objSalInvMain.str_Country)
            BaseConn.cmd.Parameters.AddWithValue("@Filter3", obj.objSalInvMain.str_Filter3)
            BaseConn.cmd.Parameters.AddWithValue("@Filter4", obj.objSalInvMain.str_Filter4)
            BaseConn.cmd.Parameters.AddWithValue("@DODT", obj.objSalInvMain.dt_DONo4SIS)
            BaseConn.cmd.Parameters.AddWithValue("@InvTaxAmountDT", obj.objSalInvMain.dt_TaxItemDetails)

            BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output

            BaseConn.cmd.Parameters.Add("@OutSMSMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutEmailMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output

            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            CSNo = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            _OutSMS = BaseConn.cmd.Parameters("@OutSMSMsgText").Value.ToString
            _OutEmail = BaseConn.cmd.Parameters("@OutEmailMsgText").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.int_CID)
            ObjDalGeneral.Elog_Insert(obj.int_CID, _strPath, _strPwd, obj.objSalInvMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "SIS", Err.Number, "Error in " & obj.objSalInvMain.str_Flag & " : " & obj.objSalInvMain.str_SalInvNo & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_CS = _ErrString
    End Function

    Public Sub ImportSISfromExcel(ByVal _strPath As String, ByVal _strPwd As String, ByVal _SiteID As String, ByVal _BSID As Integer, _
                                       ByVal _JVLedgerID As Integer, ByVal _SISMainDT As DataTable, _
                              ByVal _CreatedBy As String, ByRef ErrNo As Integer, ByRef _ErrDesc As String)
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("ImportSISfromExcel", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BSID)
            BaseConn.cmd.Parameters.AddWithValue("@JVLedgerID", _JVLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@SISMainDT", _SISMainDT)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", _CreatedBy)

            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 2000
            BaseConn.cmd.ExecuteNonQuery()

            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            _ErrDesc = _ErrString
        Catch ex As Exception
            _ErrDesc = ex.Message
            ObjDalGeneral = New DAL_General(_SiteID)
            ObjDalGeneral.Elog_Insert(_SiteID, _strPath, _strPwd, _BSID, _CreatedBy, Date.Now, "", "SIS", Err.Number, "Error in Import from Excel :", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

    End Sub
End Class

