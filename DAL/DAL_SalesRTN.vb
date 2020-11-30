'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes

Public Class DAL_SalesRTN
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef Obj As csSalesRTN, ByVal _strPath As String, ByVal _strPwd As String, ByRef iRC As Integer, ByRef ErrStr As String)
        iRC = 0
        ErrStr = ""
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetSalesRTNDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.objSalRtnMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@SRNo", Obj.objSalRtnMain.str_SRNo)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.objSalRtnMain.str_Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            If Obj.objSalRtnMain.str_Flag = "SRT" Then
                Obj.objSalRtnMain.str_InvRef = ds.Tables(0).Rows(0)("InvRef").ToString()
                Obj.objSalRtnMain.str_DoRef = ds.Tables(0).Rows(0)("DoRef").ToString()
                Obj.objSalRtnMain.str_LpoRef = ds.Tables(0).Rows(0)("LpoRef").ToString()
                Obj.objSalRtnMain.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
                Obj.objSalRtnMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
                Obj.objSalRtnMain.dtp_RTNDate1 = ds.Tables(0).Rows(0)("RTNDate1")
                Obj.objSalRtnMain.dtp_RTNDate2 = ds.Tables(0).Rows(0)("RTNDate2")
                Obj.objSalRtnMain.str_SalesManID = ds.Tables(0).Rows(0)("SalesManID").ToString()
                Obj.objSalRtnMain.int_Aging = ds.Tables(0).Rows(0)("Aging").ToString()
                Obj.objSalRtnMain.str_PayTerm = ds.Tables(0).Rows(0)("PayTerm").ToString()
                Obj.objSalRtnMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
                Obj.objSalRtnMain.dbl_TCAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
                Obj.objSalRtnMain.dbl_TCItemTaxAmount = ds.Tables(0).Rows(0)("TCItemTaxAmount")
                Obj.objSalRtnMain.dbl_TCInvoiceTaxAmount = ds.Tables(0).Rows(0)("TCInvTaxAmount")
                Obj.objSalRtnMain.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
                Obj.objSalRtnMain.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
                Obj.objSalRtnMain.dbl_TCAdjAMount = ds.Tables(0).Rows(0)("TCAdjAmount").ToString()
                Obj.objSalRtnMain.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
                Obj.objSalRtnMain.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString()
                Obj.objSalRtnMain.str_TCCurrency = ds.Tables(0).Rows(0)("TCCurrency").ToString()
                Obj.objSalRtnMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()
                Obj.objSalRtnMain.int_StatusCancel = ds.Tables(0).Rows(0)("StatusCancel").ToString()
                Obj.objSalRtnMain.str_RtnVouType = ds.Tables(0).Rows(0)("RtnVouType").ToString()
                Obj.objSalRtnMain.str_RtnVouTypeNo = ds.Tables(0).Rows(0)("RtnVouTypeNo").ToString()

                Obj.objSalRtnMain.str_WHID = ds.Tables(0).Rows(0)("WHID").ToString()

                Obj.objSalRtnMain.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
                Obj.objSalRtnMain.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
                Obj.objSalRtnMain.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
                Obj.objSalRtnMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
                Obj.objSalRtnMain.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
                Obj.objSalRtnMain.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
                Obj.objSalRtnMain.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
                Obj.objSalRtnMain.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()
                'Obj.objSalRtnMain.str_TaxCode = ds.Tables(0).Rows(0)("TaxCode").ToString()

                Obj.objSalRtnMain.str_ItemTaxCode = ds.Tables(0).Rows(0)("ItemTaxCode")
                Obj.objSalRtnMain.str_InvoiceTaxCode = ds.Tables(0).Rows(0)("InvoiceTaxCode")
                'Obj.objSalInvMain.dbl_ItemDiscPercentage = ds.Tables(0).Rows(0)("ItemDiscPercentage")
                Obj.objSalRtnMain.str_InvoiceTaxXML = ds.Tables(0).Rows(0)("InvoiceTaxDetails")


                Obj.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
                Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
                Obj.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
                Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
                Obj.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
                Obj.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
                Obj.objSalRtnMain.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus")
                Obj.objSalRtnMain.str_UserComment = ds.Tables(0).Rows(0)("UserComment").ToString()
                Obj.objSalRtnMain.bool_TaxFileReturn = ds.Tables(0).Rows(0)("TaxReturnFiled")
                Obj.objSalRtnMain.str_DiscText = ds.Tables(0).Rows(0)("DiscText").ToString()
                If ds.Tables(1).Rows.Count > 0 Then
                    Obj.objSalRtnSub.dt_SalRtn = ds.Tables(1)
                End If

                If ds.Tables(2).Rows.Count > 0 Then
                    Obj.objproject.str_ProjectID = ds.Tables(2).Rows(0)("ProjectID").ToString()
                    Obj.objproject.str_ProjectLocation = ds.Tables(2).Rows(0)("ProjectLocation").ToString()
                    Obj.objproject.str_WorkOrderNo = ds.Tables(2).Rows(0)("WorkOrderNo").ToString()
                End If

                If ds.Tables(3).Rows.Count > 0 Then
                    Obj.objSalRtnMain.dt_InvoiceAccounts = ds.Tables(3)
                End If
                Obj.objSalRtnMain.dt_TaxItemDetails = ds.Tables(4)
                Obj.DTBatch = ds.Tables(5)
            ElseIf Obj.objSalRtnMain.str_Flag = "DO" Then
                Obj.objSalRtnMain.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
                Obj.objSalRtnMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
                Obj.objSalRtnMain.dtp_RTNDate1 = Date.Now
                Obj.objSalRtnMain.dtp_RTNDate2 = Date.Now
                Obj.objSalRtnMain.str_SalesManID = ds.Tables(0).Rows(0)("SalesManID").ToString()
                Obj.objSalRtnMain.int_Aging = ds.Tables(0).Rows(0)("Aging").ToString()
                Obj.objSalRtnMain.str_PayTerm = ds.Tables(0).Rows(0)("PayTerm").ToString()
                Obj.objSalRtnMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
                Obj.objSalRtnMain.dbl_TCAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
                Obj.objSalRtnMain.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
                Obj.objSalRtnMain.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
                Obj.objSalRtnMain.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
                Obj.objSalRtnMain.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString()
                Obj.objSalRtnMain.str_TCCurrency = ds.Tables(0).Rows(0)("TCCurrency").ToString()
                Obj.objSalRtnMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()
                Obj.objSalRtnMain.str_WHID = ds.Tables(0).Rows(0)("WHID").ToString()
                Obj.objSalRtnMain.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
                Obj.objSalRtnMain.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
                Obj.objSalRtnMain.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
                Obj.objSalRtnMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
                Obj.objSalRtnMain.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
                Obj.objSalRtnMain.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
                Obj.objSalRtnMain.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
                Obj.objSalRtnMain.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()
                Obj.objSalRtnMain.str_ItemTaxCode = ds.Tables(0).Rows(0)("TaxCode").ToString()
                Obj.objSalRtnMain.str_DiscText = ds.Tables(0).Rows(0)("DiscText").ToString()
                If ds.Tables(1).Rows.Count > 0 Then
                    Obj.objSalRtnSub.dt_SalRtn = ds.Tables(1)
                End If
                If ds.Tables(2).Rows.Count > 0 Then
                    Obj.objproject.str_ProjectID = ds.Tables(2).Rows(0)("ProjectID").ToString()
                    Obj.objproject.str_ProjectLocation = ds.Tables(2).Rows(0)("ProjectLocation").ToString()
                    Obj.objproject.str_WorkOrderNo = ds.Tables(2).Rows(0)("WorkOrderNo").ToString()
                End If
                Obj.objSalRtnMain.dt_TaxItemDetails = ds.Tables(3)
                Obj.DTBatch = ds.Tables(4)
            ElseIf Obj.objSalRtnMain.str_Flag = "CS" Then
                Obj.objSalRtnMain.int_LedgerID = ds.Tables(0).Rows(0)("SrcLedgerID").ToString()
                Obj.objSalRtnMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
                Obj.objSalRtnMain.dtp_RTNDate1 = Date.Now
                Obj.objSalRtnMain.dtp_RTNDate2 = Date.Now
                Obj.objSalRtnMain.str_SalesManID = ds.Tables(0).Rows(0)("SalesManID").ToString()
                Obj.objSalRtnMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
                Obj.objSalRtnMain.dbl_TCAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
                Obj.objSalRtnMain.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
                Obj.objSalRtnMain.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
                Obj.objSalRtnMain.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
                Obj.objSalRtnMain.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString()
                Obj.objSalRtnMain.str_TCCurrency = ds.Tables(0).Rows(0)("TCCurrency").ToString()
                Obj.objSalRtnMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()
                Obj.objSalRtnMain.str_WHID = ds.Tables(0).Rows(0)("WHID").ToString()
                Obj.objSalRtnMain.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
                Obj.objSalRtnMain.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
                Obj.objSalRtnMain.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
                Obj.objSalRtnMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
                Obj.objSalRtnMain.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
                Obj.objSalRtnMain.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
                Obj.objSalRtnMain.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
                Obj.objSalRtnMain.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()
                Obj.objSalRtnMain.str_ItemTaxCode = ds.Tables(0).Rows(0)("TaxCode").ToString()
                Obj.objSalRtnMain.str_DiscText = ds.Tables(0).Rows(0)("DiscText").ToString()
                If ds.Tables(1).Rows.Count > 0 Then
                    Obj.objSalRtnSub.dt_SalRtn = ds.Tables(1)
                End If

                If ds.Tables(2).Rows.Count > 0 Then
                    Obj.objproject.str_ProjectID = ds.Tables(2).Rows(0)("ProjectID").ToString()
                    Obj.objproject.str_ProjectLocation = ds.Tables(2).Rows(0)("ProjectLocation").ToString()
                    Obj.objproject.str_WorkOrderNo = ds.Tables(2).Rows(0)("WorkOrderNo").ToString()
                End If

                If ds.Tables(3).Rows.Count > 0 Then
                    Obj.objSalRtnMain.dt_InvoiceAccounts = ds.Tables(3)
                End If
                Obj.objSalRtnMain.dt_TaxItemDetails = ds.Tables(4)
                Obj.DTBatch = ds.Tables(5)
            End If
        Catch ex As Exception

            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try

    End Sub

    Public Function Update_SalRtn(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef SrtNo As String, ByRef intRevNo As Integer, ByVal obj As csSalesRTN, ByRef _OutSMS As String, ByRef _OutEmail As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("SalesRTNUpdated", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.objSalRtnMain.int_BusinessPeriodID)

            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.objSalRtnMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.objSalRtnMain.str_SalesRtnPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@SRNo", obj.objSalRtnMain.str_SRNo)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.objSalRtnMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@RTNDate1", obj.objSalRtnMain.dtp_RTNDate1)
            BaseConn.cmd.Parameters.AddWithValue("@RTNDate2", obj.objSalRtnMain.dtp_RTNDate2)
            BaseConn.cmd.Parameters.AddWithValue("@InvRef", obj.objSalRtnMain.str_InvRef)
            BaseConn.cmd.Parameters.AddWithValue("@DoRef", obj.objSalRtnMain.str_DoRef)
            BaseConn.cmd.Parameters.AddWithValue("@LpoRef", obj.objSalRtnMain.str_LpoRef)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.objSalRtnMain.int_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Alias", obj.objSalRtnMain.str_Alias)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManID", obj.objSalRtnMain.str_SalesManID)
            BaseConn.cmd.Parameters.AddWithValue("@Aging", obj.objSalRtnMain.int_Aging)
            BaseConn.cmd.Parameters.AddWithValue("@PayTerm", obj.objSalRtnMain.str_PayTerm)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.objSalRtnMain.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@TCAmount", obj.objSalRtnMain.dbl_TCAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDisAmount", obj.objSalRtnMain.dbl_TCDisAmount)
            BaseConn.cmd.Parameters.AddWithValue("@InvoiceTaxXML", obj.objSalRtnMain.str_InvoiceTaxXML)
            BaseConn.cmd.Parameters.AddWithValue("@TCItemTaxAmount", obj.objSalRtnMain.dbl_TCItemTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCInvoiceTaxAmount", obj.objSalRtnMain.dbl_TCInvoiceTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDiscountAmount", obj.objSalRtnMain.dbl_TCDiscountAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCAdjAmount", obj.objSalRtnMain.dbl_TCAdjAMount)
            BaseConn.cmd.Parameters.AddWithValue("@TCNetAmount", obj.objSalRtnMain.dbl_TCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@LCNetAmount", obj.objSalRtnMain.dbl_LCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", obj.objSalRtnMain.str_TCCurrency)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", obj.objSalRtnMain.dbl_ExchangeRate)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.objSalRtnMain.bool_ApprovedStatus)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedLevel", obj.ApprovedLevel)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedComment", obj.ApprovedComment)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedHigherLevel", obj.ApprovedHigherLevel)
            BaseConn.cmd.Parameters.AddWithValue("@RtnStatus", obj.objSalRtnMain.bool_RtnStatus)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectID", obj.objproject.str_ProjectID)
            BaseConn.cmd.Parameters.AddWithValue("@WorkOrderNo", obj.objproject.str_WorkOrderNo)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectLocation", obj.objproject.str_ProjectLocation)
            BaseConn.cmd.Parameters.AddWithValue("@ItemTaxCode", obj.objSalRtnMain.str_ItemTaxCode)
            BaseConn.cmd.Parameters.AddWithValue("@InvoiceTaxCode", obj.objSalRtnMain.str_InvoiceTaxCode)
            BaseConn.cmd.Parameters.AddWithValue("@LanguageCode", obj.objSalRtnMain.int_LanguageCode)

            BaseConn.cmd.Parameters.AddWithValue("@StatusCancel", obj.objSalRtnMain.int_StatusCancel)
            BaseConn.cmd.Parameters.AddWithValue("@UserComment", obj.objSalRtnMain.str_UserComment)
            BaseConn.cmd.Parameters.AddWithValue("@RtnVouType", obj.objSalRtnMain.str_RtnVouType)
            BaseConn.cmd.Parameters.AddWithValue("@RtnVouTypeNo", obj.objSalRtnMain.str_RtnVouTypeNo)


            BaseConn.cmd.Parameters.AddWithValue("@WHID", obj.objSalRtnMain.str_WHID)

            ''AM Specific
            BaseConn.cmd.Parameters.AddWithValue("@Desc1", obj.objSalRtnMain.str_Desc1)
            BaseConn.cmd.Parameters.AddWithValue("@Desc2", obj.objSalRtnMain.str_Desc2)
            BaseConn.cmd.Parameters.AddWithValue("@Desc3", obj.objSalRtnMain.str_Desc3)
            BaseConn.cmd.Parameters.AddWithValue("@Desc4", obj.objSalRtnMain.str_Desc4)
            BaseConn.cmd.Parameters.AddWithValue("@Desc5", obj.objSalRtnMain.str_Desc5)
            BaseConn.cmd.Parameters.AddWithValue("@Desc6", obj.objSalRtnMain.str_Desc6)
            BaseConn.cmd.Parameters.AddWithValue("@Desc7", obj.objSalRtnMain.str_Desc7)
            BaseConn.cmd.Parameters.AddWithValue("@Desc8", obj.objSalRtnMain.str_Desc8)
            'BaseConn.cmd.Parameters.AddWithValue("@TaxCode", obj.objSalRtnMain.str_TaxCode)
            BaseConn.cmd.Parameters.AddWithValue("@DiscText", obj.objSalRtnMain.str_DiscText)
            BaseConn.cmd.Parameters.AddWithValue("@SRTItemDetailsDT", obj.objSalRtnSub.dt_SalRtn)
            BaseConn.cmd.Parameters.AddWithValue("@MatchingDT", obj.objSalRtnSub.dt_SRTMatching)
            BaseConn.cmd.Parameters.AddWithValue("@InvAccDetDT", obj.objSalRtnMain.dt_InvoiceAccounts)
            BaseConn.cmd.Parameters.AddWithValue("@ItemBatchDT", obj.DTBatch)
            BaseConn.cmd.Parameters.AddWithValue("@InvTaxAmountDT", obj.objSalRtnMain.dt_TaxItemDetails)

            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.objSalRtnMain.str_Flag)
            BaseConn.cmd.Parameters.Add("@SRNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutSMSMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutEmailMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()

            SrtNo = BaseConn.cmd.Parameters("@SRNoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            _OutSMS = BaseConn.cmd.Parameters("@OutSMSMsgText").Value.ToString
            _OutEmail = BaseConn.cmd.Parameters("@OutEmailMsgText").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.objSalRtnMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "SalesRTN", Err.Number, "Error in " & obj.objSalRtnMain.str_Flag & " : " & obj.objSalRtnMain.str_SRNo & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_SalRtn = _ErrString
    End Function

    Public Sub ImportSRTfromExcel(ByVal _strPath As String, ByVal _strPwd As String, ByVal _SiteID As String, ByVal _BSID As Integer, _
                                       ByVal _JVLedgerID As Integer, ByVal _SRTMainDT As DataTable, _
                              ByVal _CreatedBy As String, ByRef ErrNo As Integer, ByRef _ErrDesc As String)
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("SP_ImportSRTfromExcel", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BSID)
            BaseConn.cmd.Parameters.AddWithValue("@JVLedgerID", _JVLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@SRTMainDT", _SRTMainDT)
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
