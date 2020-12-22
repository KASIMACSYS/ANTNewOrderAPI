'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes


Public Class DAL_Quotation
    'Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Function GetSalesmanQuotation(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _CID As Integer, ByVal _SalesmanID As Integer,
                                        ByVal _FromDate As Date, ByVal _ToDate As Date, ByVal _Status As String, ByRef iRC As Integer, ByRef ErrStr As String) As DataTable
        GetSalesmanQuotation = New DataTable
        iRC = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[MA_GetQuotationAgainstSalesman]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@SalesmanID", _SalesmanID)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Status", _Status)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            GetSalesmanQuotation = ds.Tables(0)
        Catch ex As Exception
            iRC = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try

        Return GetSalesmanQuotation
    End Function

    Public Function GetSalesmanQuotation1(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _CID As Integer, ByVal _SalesmanID As Integer,
                                       ByVal _Status As String, ByRef iRC As Integer, ByRef ErrStr As String) As DataTable
        GetSalesmanQuotation1 = New DataTable
        iRC = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[MA_GetQuotationAgainstSalesman1]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@SalesmanID", _SalesmanID)
            BaseConn.cmd.Parameters.AddWithValue("@Status", _Status)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            GetSalesmanQuotation1 = ds.Tables(0)
        Catch ex As Exception
            iRC = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try

        Return GetSalesmanQuotation1
    End Function

    Public Function MA_QuotationDashboard(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _CID As Integer, ByVal _SalesmanID As Integer, ByRef iRC As Integer, ByRef ErrStr As String) As DataTable
        MA_QuotationDashboard = New DataTable
        iRC = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[MA_QuotationDashboard]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@SalesmanID", _SalesmanID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            MA_QuotationDashboard = ds.Tables(0)
        Catch ex As Exception
            iRC = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try

        Return MA_QuotationDashboard
    End Function

    Public Sub Get_Structure(ByRef Obj As csQuotation, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef iRC As Integer, ByRef ErrStr As String)
        iRC = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetQuotationDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@QtnNo", Obj.objQuotationMain.str_QtnNo)
            'BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.objQuotationMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_CID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.objQuotationMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@RevisionHistoryNo", Obj.objQuotationMain.int_RevisionHistoryNo)
            'BaseConn.cmd.Parameters.AddWithValue("@MerchantID", Obj.objQuotationMain.int_LedgerID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            If Obj.objQuotationMain.str_Flag = "ESTIMATION" Then
                Obj.objQuotationMain.dtp_QtnDate = ds.Tables(0).Rows(0)("EstDate").ToString()
                Obj.objQuotationMain.Str_QtnStatus = ds.Tables(0).Rows(0)("Status").ToString()
                Obj.objQuotationMain.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
                Obj.objQuotationMain.str_SalesManID = ds.Tables(0).Rows(0)("SalesManID").ToString()
                Obj.objQuotationMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
                Obj.objQuotationMain.str_TCCurrency = "AED"
                Obj.objQuotationMain.dbl_ExchangeRate = 1
                Obj.objQuotationMain.str_RTF_Description = ""
                Obj.objQuotationMain.str_ItemTaxCode = "TAX"
                Obj.objQuotationMain.str_InvoiceTaxCode = "TAX"
                Obj.objQuotationMain.str_InvoiceTaxXML = "TAX"
                Obj.objQuotationMain.str_MiscText = "Misc"
                Obj.objQuotationMain.str_DiscText = "Disc"
            Else
                Obj.objQuotationMain.str_EstNo = ds.Tables(0).Rows(0)("EstNo").ToString()
                Obj.objQuotationMain.dtp_QtnDate = ds.Tables(0).Rows(0)("QtnDate").ToString()
                Obj.objQuotationMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
                Obj.objQuotationMain.Str_QtnStatus = ds.Tables(0).Rows(0)("QtnStatus").ToString()
                Obj.objQuotationMain.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
                Obj.objQuotationMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
                Obj.objQuotationMain.int_Aging = ds.Tables(0).Rows(0)("Aging").ToString()
                Obj.objQuotationMain.str_PayTerm = ds.Tables(0).Rows(0)("PayTerm").ToString()
                Obj.objQuotationMain.str_IndRef = ds.Tables(0).Rows(0)("IndRef").ToString()
                Obj.objQuotationMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
                Obj.objQuotationMain.str_Contact = ds.Tables(0).Rows(0)("Contact").ToString()
                Obj.objQuotationMain.str_SalesManID = ds.Tables(0).Rows(0)("SalesManID").ToString()
                Obj.objQuotationMain.str_TCCurrency = ds.Tables(0).Rows(0)("TCCurrency").ToString()
                Obj.objQuotationMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()
                Obj.objQuotationMain.dbl_TCAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
                Obj.objQuotationMain.dbl_TCItemTaxAmount = ds.Tables(0).Rows(0)("TCItemTaxAmount")
                Obj.objQuotationMain.dbl_TCInvoiceTaxAmount = ds.Tables(0).Rows(0)("TCInvTaxAmount")
                Obj.objQuotationMain.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
                Obj.objQuotationMain.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
                Obj.objQuotationMain.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
                Obj.objQuotationMain.dbl_TCMiscAmount = ds.Tables(0).Rows(0)("TCMiscAmount").ToString()
                Obj.objQuotationMain.dbl_TCMiscPercentage = ds.Tables(0).Rows(0)("TCMiscPercentage").ToString()
                Obj.objQuotationMain.dbl_TCAdjAmount = ds.Tables(0).Rows(0)("TCAdjAmount").ToString()
                Obj.objQuotationMain.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString()
                Obj.objQuotationMain.int_StatusCancel = ds.Tables(0).Rows(0)("StatusCancel").ToString()
                Obj.objQuotationMain.str_MiscText = ds.Tables(0).Rows(0)("MiscText").ToString()
                Obj.objQuotationMain.str_DiscText = ds.Tables(0).Rows(0)("DiscText").ToString()
                Obj.objQuotationMain._XMLCustomData = ds.Tables(0).Rows(0)("XMLData1").ToString()

                'Obj.objQuotationMain.str_Surface = ds.Tables(0).Rows(0)("Desc1").ToString()
                Obj.objQuotationMain.str_DeliverIn = ds.Tables(0).Rows(0)("DeliveryIn").ToString()
                Obj.objQuotationMain.str_QtnValidity = ds.Tables(0).Rows(0)("QuotationValidity").ToString()

                Obj.objQuotationMain.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
                Obj.objQuotationMain.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
                Obj.objQuotationMain.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
                Obj.objQuotationMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
                Obj.objQuotationMain.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
                Obj.objQuotationMain.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
                Obj.objQuotationMain.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
                Obj.objQuotationMain.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()

                'Obj.objQuotationMain.str_TaxCode = ds.Tables(0).Rows(0)("TaxCode").ToString()
                Obj.objQuotationMain.str_ItemTaxCode = ds.Tables(0).Rows(0)("ItemTaxCode")
                Obj.objQuotationMain.str_InvoiceTaxCode = ds.Tables(0).Rows(0)("InvoiceTaxCode")
                Obj.objQuotationMain.str_InvoiceTaxXML = ds.Tables(0).Rows(0)("InvoiceTaxDetails")

                Obj.objQuotationMain.dbl_ItemDiscPercentage = ds.Tables(0).Rows(0)("ItemDiscPercentage").ToString()

                Obj.objQuotationMain.str_ExpiryDays = ds.Tables(0).Rows(0)("ExpiryDays").ToString()

                Obj.objQuotationMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()
                Obj.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
                Obj.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
                Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus").ToString()

                Obj.objQuotationMain.str_UserComment = ds.Tables(0).Rows(0)("UserComment").ToString()
                Obj.objQuotationMain.str_ApproverComment = ds.Tables(0).Rows(0)("ApproverComment").ToString()

            End If


            Obj.objQuotationSub.dt_Quotation = ds.Tables(1)

            If ds.Tables(2).Rows.Count > 0 Then
                Obj.objproject.str_ProjectID = ds.Tables(2).Rows(0)("ProjectID").ToString()
                Obj.objproject.str_ProjectLocation = ds.Tables(2).Rows(0)("ProjectLocation").ToString()
                Obj.objproject.str_WorkOrderNo = ds.Tables(2).Rows(0)("WorkOrderNo").ToString()
            Else
                Obj.objproject.str_ProjectID = ""
                Obj.objproject.str_ProjectLocation = ""
                Obj.objproject.str_WorkOrderNo = ""
            End If

            'Obj.objQuotationMain.dt_TaxItemDetails = ds.Tables(3)
            If Obj.objQuotationMain.str_Flag = "QUOTATION" Then
                Obj.DTItemExtraDetails = ds.Tables(4)

                If ds.Tables(5).Rows.Count > 0 Then
                    Obj.objQuotationMain.str_RTF_Description = ds.Tables(5).Rows(0)("Description").ToString()
                Else
                    Obj.objQuotationMain.str_RTF_Description = ""
                End If
            End If



        Catch ex As Exception
            iRC = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Update_Quotation(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef QtnNo As String, ByRef intRevNo As Integer, ByVal obj As csQuotation, ByRef _OutSMS As String, ByRef _OutEmail As String, ByRef ErrNo As Integer) As String
        'ObjDalGeneral = New DAL_General(obj.str_CID)
        'Dim JsonString As String = ObjDalGeneral.ClassToJSon(obj.objQuotationMain)

        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("QuotationUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_CID) 'obj.str_SiteID
            'BaseConn.cmd.Parameters.AddWithValue("@Json", JsonString)

            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.objQuotationMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.objQuotationMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.objQuotationMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.objQuotationMain.str_FormPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@LanguageCode", obj.objQuotationMain.int_LanguageCode)

            BaseConn.cmd.Parameters.AddWithValue("@QtnNo", obj.objQuotationMain.str_QtnNo)
            BaseConn.cmd.Parameters.AddWithValue("@EstNo", obj.objQuotationMain.str_EstNo)
            BaseConn.cmd.Parameters.AddWithValue("@QtnDate", obj.objQuotationMain.dtp_QtnDate)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.objQuotationMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.objQuotationMain.int_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Alias", obj.objQuotationMain.str_Alias)
            BaseConn.cmd.Parameters.AddWithValue("@Aging", obj.objQuotationMain.int_Aging)
            BaseConn.cmd.Parameters.AddWithValue("@PayTerm", obj.objQuotationMain.str_PayTerm)
            BaseConn.cmd.Parameters.AddWithValue("@IndRef", obj.objQuotationMain.str_IndRef)
            BaseConn.cmd.Parameters.AddWithValue("@QtnStatus", obj.objQuotationMain.Str_QtnStatus)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.objQuotationMain.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@StatusCancel", obj.objQuotationMain.int_StatusCancel)

            BaseConn.cmd.Parameters.AddWithValue("@Contact", obj.objQuotationMain.str_Contact)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManID", obj.objQuotationMain.str_SalesManID)
            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", obj.objQuotationMain.str_TCCurrency)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", obj.objQuotationMain.dbl_ExchangeRate)

            BaseConn.cmd.Parameters.AddWithValue("@TCAmount", obj.objQuotationMain.dbl_TCAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDisAmount", obj.objQuotationMain.dbl_TCDisAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDiscountAmount", obj.objQuotationMain.dbl_TCDiscountAmount)
            BaseConn.cmd.Parameters.AddWithValue("@InvoiceTaxXML", obj.objQuotationMain.str_InvoiceTaxXML)
            BaseConn.cmd.Parameters.AddWithValue("@TCItemTaxAmount", obj.objQuotationMain.dbl_TCItemTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCInvoiceTaxAmount", obj.objQuotationMain.dbl_TCInvoiceTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCNetAmount", obj.objQuotationMain.dbl_TCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCMiscAmount", obj.objQuotationMain.dbl_TCMiscAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCMiscPercentage", obj.objQuotationMain.dbl_TCMiscPercentage)
            BaseConn.cmd.Parameters.AddWithValue("@TCAdjAmount", obj.objQuotationMain.dbl_TCAdjAmount)
            BaseConn.cmd.Parameters.AddWithValue("@LCNetAmount", obj.objQuotationMain.dbl_LCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@MiscText", obj.objQuotationMain.str_MiscText)
            BaseConn.cmd.Parameters.AddWithValue("@DiscText", obj.objQuotationMain.str_DiscText)
            BaseConn.cmd.Parameters.AddWithValue("@XMLDATA1", obj.objQuotationMain._XMLCustomData)

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

            BaseConn.cmd.Parameters.AddWithValue("@ProjectID", obj.objproject.str_ProjectID)
            BaseConn.cmd.Parameters.AddWithValue("@WorkOrderNo", obj.objproject.str_WorkOrderNo)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectLocation", obj.objproject.str_ProjectLocation)

            'BaseConn.cmd.Parameters.AddWithValue("@Desc1", obj.objQuotationMain.str_Surface)
            BaseConn.cmd.Parameters.AddWithValue("@DeliveryIn", obj.objQuotationMain.str_DeliverIn)
            BaseConn.cmd.Parameters.AddWithValue("@QuotationValidity", obj.objQuotationMain.str_QtnValidity)

            BaseConn.cmd.Parameters.AddWithValue("@Desc1", obj.objQuotationMain.str_Desc1)
            BaseConn.cmd.Parameters.AddWithValue("@Desc2", obj.objQuotationMain.str_Desc2)
            BaseConn.cmd.Parameters.AddWithValue("@Desc3", obj.objQuotationMain.str_Desc3)
            BaseConn.cmd.Parameters.AddWithValue("@Desc4", obj.objQuotationMain.str_Desc4)
            BaseConn.cmd.Parameters.AddWithValue("@Desc5", obj.objQuotationMain.str_Desc5)
            BaseConn.cmd.Parameters.AddWithValue("@Desc6", obj.objQuotationMain.str_Desc6)
            BaseConn.cmd.Parameters.AddWithValue("@Desc7", obj.objQuotationMain.str_Desc7)
            BaseConn.cmd.Parameters.AddWithValue("@Desc8", obj.objQuotationMain.str_Desc8)

            BaseConn.cmd.Parameters.AddWithValue("@ItemTaxCode", obj.objQuotationMain.str_ItemTaxCode)
            BaseConn.cmd.Parameters.AddWithValue("@InvoiceTaxCode", obj.objQuotationMain.str_InvoiceTaxCode)

            BaseConn.cmd.Parameters.AddWithValue("@ItemDiscPercentage", obj.objQuotationMain.dbl_ItemDiscPercentage)

            BaseConn.cmd.Parameters.AddWithValue("@ExpiryDays", obj.objQuotationMain.str_ExpiryDays)
            BaseConn.cmd.Parameters.AddWithValue("@UserComment", obj.objQuotationMain.str_UserComment)
            BaseConn.cmd.Parameters.AddWithValue("@ApproverComment", obj.objQuotationMain.str_ApproverComment)
            BaseConn.cmd.Parameters.AddWithValue("@RTF_Description", obj.objQuotationMain.str_RTF_Description)

            BaseConn.cmd.Parameters.AddWithValue("@QuotationItemDetailsDT", obj.objQuotationSub.dt_Quotation)
            BaseConn.cmd.Parameters.AddWithValue("@VoucherItemExtraDetailsDT", obj.DTItemExtraDetails)
            BaseConn.cmd.Parameters.AddWithValue("@InvTaxAmountDT", obj.objQuotationMain.dt_TaxItemDetails)

            BaseConn.cmd.Parameters.Add("@QtnNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output

            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output

            BaseConn.cmd.Parameters.Add("@OutSMSMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutEmailMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output

            BaseConn.cmd.ExecuteNonQuery()

            QtnNo = BaseConn.cmd.Parameters("@QtnNoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString

            _OutSMS = BaseConn.cmd.Parameters("@OutSMSMsgText").Value.ToString
            _OutEmail = BaseConn.cmd.Parameters("@OutEmailMsgText").Value.ToString

        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_CID)
            ObjDalGeneral.Elog_Insert(obj.str_CID, _StrDBPath, _StrDBPwd, obj.objQuotationMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "Quotation", Err.Number, "Error in " & obj.objQuotationMain.str_Flag & " : " & obj.objQuotationMain.str_QtnNo & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_Quotation = _ErrString
    End Function


End Class
