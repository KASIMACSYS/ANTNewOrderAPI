'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes

Public Class DAL_SalesOrder
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef Obj As csSalesOrder, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef iRC As Integer, ByRef ErrStr As String)
        iRC = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetSalesOrder]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.Add("@SalOrd", SqlDbType.VarChar).Value = Obj.objSalesOrderMain.str_SalOrd
            BaseConn.cmd.Parameters.Add("@BusinessID", SqlDbType.Int).Value = Obj.objSalesOrderMain.int_BusinessPeriodID
            BaseConn.cmd.Parameters.Add("@Flag", SqlDbType.VarChar).Value = Obj.objSalesOrderMain.str_Flag
            BaseConn.cmd.Parameters.Add("@CID", SqlDbType.VarChar).Value = Obj.int_CID
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            If Obj.objSalesOrderMain.str_Flag = "QUOTATION" Then
                Obj.objSalesOrderMain.dtp_SODate = Date.Now
                Obj.objSalesOrderMain.str_ContactPerson = ds.Tables(0).Rows(0)("Contact").ToString()
                Obj.objSalesOrderMain.dtp_QuotationDate = ds.Tables(0).Rows(0)("QtnDate").ToString()
                Obj.objSalesOrderMain.str_ExpiryDays = ds.Tables(0).Rows(0)("ExpiryDays").ToString()
                Obj.objSalesOrderMain.str_MerchantRef = ""
                Obj.objSalesOrderMain.str_DeliveryAddress = ds.Tables(0).Rows(0)("DeliveryIn").ToString()
                Obj.objSalesOrderMain.str_Consignee = ""
                Obj.objSalesOrderMain.str_SalesType = ""
                Obj.objSalesOrderMain.str_DeliveryCountry = ""

                'Obj.ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus").ToString() 'TODO
            Else
                Obj.objSalesOrderMain.str_SOStatus = ds.Tables(0).Rows(0)("SOStatus").ToString()
                Obj.objSalesOrderMain.str_MerchantRef = ds.Tables(0).Rows(0)("MerchantRef").ToString()
                Obj.objSalesOrderMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
                Obj.objSalesOrderMain.dtp_SODate = ds.Tables(0).Rows(0)("SODate").ToString()
                Obj.objSalesOrderMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
                Obj.objSalesOrderMain.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
                Obj.objSalesOrderMain.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
                Obj.objSalesOrderMain.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
                Obj.objSalesOrderMain.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()
                Obj.objSalesOrderMain.str_DeliveryAddress = ds.Tables(0).Rows(0)("DeliveryAddress").ToString()
                Obj.objSalesOrderMain.str_ContactPerson = ds.Tables(0).Rows(0)("ContactPerson").ToString()
                Obj.objSalesOrderMain.int_StatusCancel = ds.Tables(0).Rows(0)("StatusCancel").ToString()
                Obj.objSalesOrderMain.str_Consignee = ds.Tables(0).Rows(0)("Consignee").ToString()
                Obj.objSalesOrderMain.str_SalesType = ds.Tables(0).Rows(0)("SalesType").ToString()
                Obj.objSalesOrderMain.str_DeliveryCountry = ds.Tables(0).Rows(0)("DeliveryCountry").ToString()

                Obj.ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
                Obj.ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
                Obj.ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus").ToString() 'TODO
                Obj.objSalesOrderMain.str_WHID = ds.Tables(0).Rows(0)("WHID").ToString()
            End If
            Obj.objSalesOrderMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString() ' max business periodID rom Business tabl
            Obj.objSalesOrderMain.str_QtnNum = ds.Tables(0).Rows(0)("QtnNo").ToString()
            Obj.objSalesOrderMain.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
            Obj.objSalesOrderMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
            Obj.objSalesOrderMain.int_Aging = ds.Tables(0).Rows(0)("Aging").ToString()
            Obj.objSalesOrderMain.str_PayTerm = ds.Tables(0).Rows(0)("PayTerm").ToString()
            Obj.objSalesOrderMain.str_Indref = ds.Tables(0).Rows(0)("IndRef").ToString()
            Obj.objSalesOrderMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.objSalesOrderMain.str_Contact = ds.Tables(0).Rows(0)("Contact").ToString()
            Obj.objSalesOrderMain.str_SalesManID = ds.Tables(0).Rows(0)("SalesManID").ToString()
            Obj.objSalesOrderMain.str_TCCurrency = ds.Tables(0).Rows(0)("TCCurrency").ToString()
            Obj.objSalesOrderMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()
            Obj.objSalesOrderMain.dbl_TCAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
            Obj.objSalesOrderMain.dbl_TCItemTaxAmount = ds.Tables(0).Rows(0)("TCItemTaxAmount")
            Obj.objSalesOrderMain.dbl_TCInvoiceTaxAmount = ds.Tables(0).Rows(0)("TCInvTaxAmount")
            Obj.objSalesOrderMain.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
            Obj.objSalesOrderMain.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
            Obj.objSalesOrderMain.dbl_TCMiscPercentage = ds.Tables(0).Rows(0)("TCMiscPercentage").ToString()
            Obj.objSalesOrderMain.dbl_TCAdjAmount = ds.Tables(0).Rows(0)("TCAdjAmount").ToString()
            Obj.objSalesOrderMain.dbl_TCMiscAmount = ds.Tables(0).Rows(0)("TCMiscAmount").ToString()
            Obj.objSalesOrderMain.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
            Obj.objSalesOrderMain.str_MiscText = ds.Tables(0).Rows(0)("MiscText").ToString()
            Obj.objSalesOrderMain.str_DiscText = ds.Tables(0).Rows(0)("DiscText").ToString()
            Obj.objSalesOrderMain.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
            Obj.objSalesOrderMain.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
            Obj.objSalesOrderMain.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
            Obj.objSalesOrderMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
            Obj.objSalesOrderMain.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
            Obj.objSalesOrderMain.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
            Obj.objSalesOrderMain.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
            Obj.objSalesOrderMain.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()

            'Obj.objSalesOrderMain.str_TaxCode = ds.Tables(0).Rows(0)("TaxCode").ToString()
            Obj.objSalesOrderMain.str_ItemTaxCode = ds.Tables(0).Rows(0)("ItemTaxCode")
            Obj.objSalesOrderMain.str_InvoiceTaxCode = ds.Tables(0).Rows(0)("InvoiceTaxCode")
            Obj.objSalesOrderMain.str_InvoiceTaxXML = ds.Tables(0).Rows(0)("InvoiceTaxDetails")

            Obj.objSalesOrderMain.dbl_ItemDiscPercentage = ds.Tables(0).Rows(0)("ItemDiscPercentage").ToString()
            Obj.objSalesOrderMain.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString() 'TODO

            Obj.objSalesOrderMain.str_UserComment = ds.Tables(0).Rows(0)("UserComment").ToString()
            Obj.objSalesOrderMain.str_ApproverComment = ds.Tables(0).Rows(0)("ApproverComment").ToString()
            Obj.objSalesOrderMain.int_LanguageCode = ds.Tables(0).Rows(0)("LanguageCode")

            Obj.objSalesorderSub.dt_SalesOrderItemDetails = ds.Tables(1)

            If ds.Tables(2).Rows.Count > 0 Then
                Obj.objproject.str_ProjectID = ds.Tables(2).Rows(0)("ProjectID").ToString()
                Obj.objproject.str_ProjectLocation = ds.Tables(2).Rows(0)("ProjectLocation").ToString()
                Obj.objproject.str_WorkOrderNo = ds.Tables(2).Rows(0)("WorkOrderNo").ToString()
            Else
                Obj.objproject.str_ProjectID = ""
                Obj.objproject.str_ProjectLocation = ""
                Obj.objproject.str_WorkOrderNo = ""
            End If

            'Obj.objSalesOrderMain.dt_TaxItemDetails = ds.Tables(3)
            If Obj.objSalesOrderMain.str_Flag <> "QUOTATION" Then
                If ds.Tables(3).Rows.Count > 0 Then
                    Obj.DTBatch = ds.Tables(3)
                End If
            End If

            If ds.Tables(4).Rows.Count > 0 Then
                Obj.objSalesOrderMain.str_RTF_Description = ds.Tables(4).Rows(0)("Description").ToString()
            Else
                Obj.objSalesOrderMain.str_RTF_Description = ""
            End If

            Obj.DTItemExtraDetails = ds.Tables(5)

        Catch ex As Exception
            iRC = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Update_SalesOrder(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef SONo As String, ByRef intRevNo As Integer, ByVal obj As csSalesOrder, ByRef _OutSMS As String, ByRef _OutEmail As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("SalesOrderUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure

            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.objSalesOrderMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.objSalesOrderMain.str_FormPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.objSalesOrderMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.objSalesOrderMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.int_CID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@SalOrd", obj.objSalesOrderMain.str_SalOrd)
            BaseConn.cmd.Parameters.AddWithValue("@SODate", obj.objSalesOrderMain.dtp_SODate)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.objSalesOrderMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@QtnNo", obj.objSalesOrderMain.str_QtnNum)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.objSalesOrderMain.int_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Alias", obj.objSalesOrderMain.str_Alias)

            BaseConn.cmd.Parameters.AddWithValue("@Aging", obj.objSalesOrderMain.int_Aging)
            BaseConn.cmd.Parameters.AddWithValue("@PayTerm", obj.objSalesOrderMain.str_PayTerm)
            BaseConn.cmd.Parameters.AddWithValue("@Indref", obj.objSalesOrderMain.str_Indref)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.objSalesOrderMain.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@SOStatus", obj.objSalesOrderMain.str_SOStatus)
            BaseConn.cmd.Parameters.AddWithValue("@MerchantRef", obj.objSalesOrderMain.str_MerchantRef)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManID", obj.objSalesOrderMain.str_SalesManID)
            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", obj.objSalesOrderMain.str_TCCurrency)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", obj.objSalesOrderMain.dbl_ExchangeRate)
            BaseConn.cmd.Parameters.AddWithValue("@DeliveryAddress", obj.objSalesOrderMain.str_DeliveryAddress)
            BaseConn.cmd.Parameters.AddWithValue("@ContactPerson", obj.objSalesOrderMain.str_ContactPerson)
            BaseConn.cmd.Parameters.AddWithValue("@StatusCancel", obj.objSalesOrderMain.int_StatusCancel)

            BaseConn.cmd.Parameters.AddWithValue("@TCAmount", obj.objSalesOrderMain.dbl_TCAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDisAmount", obj.objSalesOrderMain.dbl_TCDisAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDiscountAmount", obj.objSalesOrderMain.dbl_TCDiscountAmount)
            BaseConn.cmd.Parameters.AddWithValue("@InvoiceTaxXML", obj.objSalesOrderMain.str_InvoiceTaxXML)
            BaseConn.cmd.Parameters.AddWithValue("@TCItemTaxAmount", obj.objSalesOrderMain.dbl_TCItemTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCInvoiceTaxAmount", obj.objSalesOrderMain.dbl_TCInvoiceTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCAdjAmount", obj.objSalesOrderMain.dbl_TCAdjAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCNetAmount", obj.objSalesOrderMain.dbl_TCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCMiscPercentage", obj.objSalesOrderMain.dbl_TCMiscPercentage)
            BaseConn.cmd.Parameters.AddWithValue("@TCMiscAmount", obj.objSalesOrderMain.dbl_TCMiscAmount)
            BaseConn.cmd.Parameters.AddWithValue("@LCNetAmount", obj.objSalesOrderMain.dbl_LCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@MiscText", obj.objSalesOrderMain.str_MiscText)
            BaseConn.cmd.Parameters.AddWithValue("@DiscText", obj.objSalesOrderMain.str_DiscText)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.ApprovedStatus)

            BaseConn.cmd.Parameters.AddWithValue("@ApprovedLevel", obj.ApprovedLevel)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedComment", obj.ApprovedComment)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedHigherLevel", obj.ApprovedHigherLevel)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectID", obj.objproject.str_ProjectID)
            BaseConn.cmd.Parameters.AddWithValue("@WorkOrderNo", obj.objproject.str_WorkOrderNo)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectLocation", obj.objproject.str_ProjectLocation)

            BaseConn.cmd.Parameters.AddWithValue("@Desc1", obj.objSalesOrderMain.str_Desc1)
            BaseConn.cmd.Parameters.AddWithValue("@Desc2", obj.objSalesOrderMain.str_Desc2)
            BaseConn.cmd.Parameters.AddWithValue("@Desc3", obj.objSalesOrderMain.str_Desc3)
            BaseConn.cmd.Parameters.AddWithValue("@Desc4", obj.objSalesOrderMain.str_Desc4)
            BaseConn.cmd.Parameters.AddWithValue("@Desc5", obj.objSalesOrderMain.str_Desc5)
            BaseConn.cmd.Parameters.AddWithValue("@Desc6", obj.objSalesOrderMain.str_Desc6)
            BaseConn.cmd.Parameters.AddWithValue("@Desc7", obj.objSalesOrderMain.str_Desc7)
            BaseConn.cmd.Parameters.AddWithValue("@Desc8", obj.objSalesOrderMain.str_Desc8)
            BaseConn.cmd.Parameters.AddWithValue("@WHID", obj.objSalesOrderMain.str_WHID)
            BaseConn.cmd.Parameters.AddWithValue("@Consignee", obj.objSalesOrderMain.str_Consignee)
            BaseConn.cmd.Parameters.AddWithValue("@SalesType", obj.objSalesOrderMain.str_SalesType)
            BaseConn.cmd.Parameters.AddWithValue("@DeliveryCountry", obj.objSalesOrderMain.str_DeliveryCountry)
            BaseConn.cmd.Parameters.AddWithValue("@RTF_Description", obj.objSalesOrderMain.str_RTF_Description)

            BaseConn.cmd.Parameters.AddWithValue("@ItemTaxCode", obj.objSalesOrderMain.str_ItemTaxCode)
            BaseConn.cmd.Parameters.AddWithValue("@InvoiceTaxCode", obj.objSalesOrderMain.str_InvoiceTaxCode)

            BaseConn.cmd.Parameters.AddWithValue("@ItemDiscPercentage", obj.objSalesOrderMain.dbl_ItemDiscPercentage)

            BaseConn.cmd.Parameters.AddWithValue("@UserComment", obj.objSalesOrderMain.str_UserComment)
            BaseConn.cmd.Parameters.AddWithValue("@ApproverComment", obj.objSalesOrderMain.str_ApproverComment)
            BaseConn.cmd.Parameters.AddWithValue("@LanguageCode", obj.objSalesOrderMain.int_LanguageCode)

            BaseConn.cmd.Parameters.AddWithValue("@SalesOrderItemDetailsDT", obj.objSalesorderSub.dt_SalesOrderItemDetails)
            BaseConn.cmd.Parameters.AddWithValue("@VoucherItemExtraDetailsDT", obj.DTItemExtraDetails)
            BaseConn.cmd.Parameters.AddWithValue("@InvTaxAmountDT", obj.objSalesOrderMain.dt_TaxItemDetails)
            BaseConn.cmd.Parameters.AddWithValue("@ItemBatchDT", obj.DTBatch)
            BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutSMSMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutEmailMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output

            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            SONo = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            _OutSMS = BaseConn.cmd.Parameters("@OutSMSMsgText").Value.ToString
            _OutEmail = BaseConn.cmd.Parameters("@OutEmailMsgText").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.int_CID)
            ObjDalGeneral.Elog_Insert(obj.int_CID, _StrDBPath, _StrDBPwd, obj.objSalesOrderMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "SalesOrder", Err.Number, "Error in " & obj.objSalesOrderMain.str_Flag & " : " & obj.objSalesOrderMain.str_SalOrd & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Update_SalesOrder = _ErrString
        Return Update_SalesOrder
    End Function

    Public Function GetSalesmanOrder(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _CID As Integer, ByVal _SalesmanID As Integer,
                                       ByVal _Status As String, ByRef iRC As Integer, ByRef ErrStr As String) As DataTable
        GetSalesmanOrder = New DataTable
        iRC = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[MA_GetSalesmanOrder]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@SalesmanID", _SalesmanID)
            BaseConn.cmd.Parameters.AddWithValue("@Status", _Status)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            GetSalesmanOrder = ds.Tables(0)
        Catch ex As Exception
            iRC = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try

        Return GetSalesmanOrder
    End Function

    Public Function MA_OrderDashboard(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _CID As Integer, ByVal _SalesmanID As Integer, ByRef iRC As Integer, ByRef ErrStr As String) As DataSet
        'MA_QuotationDashboard = New DataSet
        Dim ds As New DataSet

        iRC = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[MA_OrderDashboard]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@SalesmanID", _SalesmanID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(ds)
            'MA_QuotationDashboard = ds.Tables(0)
        Catch ex As Exception
            iRC = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try

        Return ds
    End Function
End Class
