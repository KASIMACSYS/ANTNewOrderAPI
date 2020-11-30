'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Imports Classes

Public Class DAL_PurchaseRTN
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef Obj As csPurchaseRTN, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetPurchaseRTNDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@PRNo", Obj.objPurRtnMain.str_PRNo)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.objPurRtnMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.objPurRtnMain.str_Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            If Obj.objPurRtnMain.str_Flag = "PRT" Then
                Obj.objPurRtnMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
                Obj.objPurRtnMain.str_InvRef = ds.Tables(0).Rows(0)("InvRef").ToString()
                Obj.objPurRtnMain.str_MrvNo = ds.Tables(0).Rows(0)("MrvNo").ToString()
                Obj.objPurRtnMain.str_LpoNo = ds.Tables(0).Rows(0)("LpoNo").ToString()
                Obj.objPurRtnMain.dtp_RtnDate1 = ds.Tables(0).Rows(0)("RtnDate1").ToString()
                Obj.objPurRtnMain.dtp_RtnDate2 = ds.Tables(0).Rows(0)("RtnDate2").ToString()
                Obj.objPurRtnMain.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
                Obj.objPurRtnMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
                Obj.objPurRtnMain.int_Aging = ds.Tables(0).Rows(0)("Aging").ToString()
                Obj.objPurRtnMain.str_PayTerm = ds.Tables(0).Rows(0)("PayTerm").ToString()
                Obj.objPurRtnMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
                Obj.objPurRtnMain.str_TCCurrency = ds.Tables(0).Rows(0)("TCCurrency").ToString()
                Obj.objPurRtnMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()
                Obj.objPurRtnMain.int_StatusCancel = ds.Tables(0).Rows(0)("Statuscancel").ToString()
                Obj.objPurRtnMain.str_RtnVouType = ds.Tables(0).Rows(0)("RtnVouType").ToString()
                Obj.objPurRtnMain.str_RtnVouTypeNo = ds.Tables(0).Rows(0)("RtnVouTypeNo").ToString()

                Obj.objPurRtnMain.dbl_TCAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
                Obj.objPurRtnMain.dbl_TCItemTaxAmount = ds.Tables(0).Rows(0)("TCItemTaxAmount")
                Obj.objPurRtnMain.dbl_TCInvoiceTaxAmount = ds.Tables(0).Rows(0)("TCInvTaxAmount")
                Obj.objPurRtnMain.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
                Obj.objPurRtnMain.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
                Obj.objPurRtnMain.dbl_TCAdjAmount = ds.Tables(0).Rows(0)("TCAdjAmount").ToString()
                Obj.objPurRtnMain.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString()
                
                Obj.objPurRtnMain.str_WHID = ds.Tables(0).Rows(0)("WHID").ToString()

                Obj.objPurRtnMain.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
                Obj.objPurRtnMain.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
                Obj.objPurRtnMain.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
                Obj.objPurRtnMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
                Obj.objPurRtnMain.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
                Obj.objPurRtnMain.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
                Obj.objPurRtnMain.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
                Obj.objPurRtnMain.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()

                Obj.objPurRtnMain.str_ItemTaxCode = ds.Tables(0).Rows(0)("ItemTaxCode")
                Obj.objPurRtnMain.str_InvoiceTaxCode = ds.Tables(0).Rows(0)("InvoiceTaxCode")
                'Obj.objSalInvMain.dbl_ItemDiscPercentage = ds.Tables(0).Rows(0)("ItemDiscPercentage")
                Obj.objPurRtnMain.str_InvoiceTaxXML = ds.Tables(0).Rows(0)("InvoiceTaxDetails")


                Obj.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
                Obj.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
                Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
                Obj.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
                Obj.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
                Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus").ToString()
                Obj.objPurRtnMain.str_UserComment = ds.Tables(0).Rows(0)("UserComment").ToString()
                Obj.objPurRtnMain.bool_TaxFileReturn = ds.Tables(0).Rows(0)("TaxReturnFiled")
                Obj.objPurRtnMain.str_DiscText = ds.Tables(0).Rows(0)("DiscText").ToString()
                Obj.objPurRtnSub.dt_PurRtn = ds.Tables(1)

                If ds.Tables(2).Rows.Count > 0 Then
                    Obj.objproject.str_ProjectID = ds.Tables(2).Rows(0)("ProjectID").ToString()
                    Obj.objproject.str_ProjectLocation = ds.Tables(2).Rows(0)("ProjectLocation").ToString()
                    Obj.objproject.str_WorkOrderNo = ds.Tables(2).Rows(0)("WorkOrderNo").ToString()
                Else
                    Obj.objproject.str_ProjectID = ""
                    Obj.objproject.str_ProjectLocation = ""
                    Obj.objproject.str_WorkOrderNo = ""
                End If

                If ds.Tables(3).Rows.Count > 0 Then
                    Obj.objPurRtnMain.dt_InvoiceAccounts = ds.Tables(3)
                End If
                Obj.objPurRtnMain.dt_TaxItemDetails = ds.Tables(4)
                Obj.DTBatch = ds.Tables(5)
            ElseIf Obj.objPurRtnMain.str_Flag = "MRV" Then
                Obj.objPurRtnMain.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
                Obj.objPurRtnMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
                Obj.objPurRtnMain.dtp_RtnDate1 = Date.Now
                Obj.objPurRtnMain.dtp_RtnDate2 = Date.Now

                Obj.objPurRtnMain.int_Aging = ds.Tables(0).Rows(0)("Aging").ToString()

                Obj.objPurRtnMain.dbl_TCAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
                Obj.objPurRtnMain.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
                Obj.objPurRtnMain.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
                Obj.objPurRtnMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
                Obj.objPurRtnMain.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString()
                Obj.objPurRtnMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()

                Obj.objPurRtnMain.str_WHID = ds.Tables(0).Rows(0)("WHID").ToString()

                Obj.objPurRtnMain.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
                Obj.objPurRtnMain.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
                Obj.objPurRtnMain.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
                Obj.objPurRtnMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
                Obj.objPurRtnMain.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
                Obj.objPurRtnMain.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
                Obj.objPurRtnMain.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
                Obj.objPurRtnMain.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()
                Obj.objPurRtnMain.str_ItemTaxCode = ds.Tables(0).Rows(0)("TaxCode")
                Obj.objPurRtnMain.str_DiscText = ds.Tables(0).Rows(0)("DiscText").ToString()
                Obj.objPurRtnMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()

                Obj.objPurRtnMain.str_TCCurrency = ds.Tables(0).Rows(0)("TCCurrency").ToString()
                Obj.objPurRtnMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()


                If ds.Tables(1).Rows.Count > 0 Then
                    Obj.objPurRtnSub.dt_PurRtn = ds.Tables(1)
                End If

                If ds.Tables(2).Rows.Count > 0 Then
                    Obj.objproject.str_ProjectID = ds.Tables(2).Rows(0)("ProjectID").ToString()
                    Obj.objproject.str_ProjectLocation = ds.Tables(2).Rows(0)("ProjectLocation").ToString()
                    Obj.objproject.str_WorkOrderNo = ds.Tables(2).Rows(0)("WorkOrderNo").ToString()
                Else
                    Obj.objproject.str_ProjectID = ""
                    Obj.objproject.str_ProjectLocation = ""
                    Obj.objproject.str_WorkOrderNo = ""
                End If

                Obj.objPurRtnMain.dt_TaxItemDetails = ds.Tables(3)
                Obj.DTBatch = ds.Tables(4)

            ElseIf Obj.objPurRtnMain.str_Flag = "CP" Then
                Obj.objPurRtnMain.int_LedgerID = ds.Tables(0).Rows(0)("SrcLedgerID").ToString()
                Obj.objPurRtnMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
                Obj.objPurRtnMain.str_InvRef = ds.Tables(0).Rows(0)("InvRef").ToString()
                Obj.objPurRtnMain.dtp_RtnDate1 = ds.Tables(0).Rows(0)("InvDate").ToString()
                Obj.objPurRtnMain.dtp_RtnDate2 = ds.Tables(0).Rows(0)("DueDate").ToString()
                Obj.objPurRtnMain.int_Aging = ds.Tables(0).Rows(0)("Aging").ToString()
                Obj.objPurRtnMain.str_PayTerm = ds.Tables(0).Rows(0)("PaymentTerm").ToString()
                Obj.objPurRtnMain.dbl_TCAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
                Obj.objPurRtnMain.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
                Obj.objPurRtnMain.dbl_TCDiscount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
                Obj.objPurRtnMain.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
                Obj.objPurRtnMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
                Obj.objPurRtnMain.str_WHID = ds.Tables(0).Rows(0)("WHID").ToString()

                Obj.objPurRtnMain.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
                Obj.objPurRtnMain.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
                Obj.objPurRtnMain.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
                Obj.objPurRtnMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
                Obj.objPurRtnMain.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
                Obj.objPurRtnMain.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
                Obj.objPurRtnMain.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
                Obj.objPurRtnMain.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()
                Obj.objPurRtnMain.str_ItemTaxCode = ds.Tables(0).Rows(0)("TaxCode")
                Obj.objPurRtnMain.str_DiscText = ds.Tables(0).Rows(0)("DiscText").ToString()
                Obj.objPurRtnMain.str_TCCurrency = ds.Tables(0).Rows(0)("TCCurrency").ToString()
                Obj.objPurRtnMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()

                If ds.Tables(1).Rows.Count > 0 Then
                    Obj.objPurRtnSub.dt_PurRtn = ds.Tables(1)
                End If

                If ds.Tables(2).Rows.Count > 0 Then
                    Obj.objproject.str_ProjectID = ds.Tables(2).Rows(0)("ProjectID").ToString()
                    Obj.objproject.str_ProjectLocation = ds.Tables(2).Rows(0)("ProjectLocation").ToString()
                    Obj.objproject.str_WorkOrderNo = ds.Tables(2).Rows(0)("WorkOrderNo").ToString()
                Else
                    Obj.objproject.str_ProjectID = ""
                    Obj.objproject.str_ProjectLocation = ""
                    Obj.objproject.str_WorkOrderNo = ""
                End If

                Obj.objPurRtnMain.dt_TaxItemDetails = ds.Tables(3)
                Obj.DTBatch = ds.Tables(4)
            End If
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Update_PurRtn(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef PRNo As String, ByRef intRevNo As Integer, ByVal obj As csPurchaseRTN, ByRef _OutSMS As String, ByRef _OutEmail As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("PurchaseRTNUpdated", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.objPurRtnMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.objPurRtnMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@PRNo", obj.objPurRtnMain.str_PRNo)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.objPurRtnMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@InvRef", obj.objPurRtnMain.str_InvRef)
            BaseConn.cmd.Parameters.AddWithValue("@MrvNo", obj.objPurRtnMain.str_MrvNo)
            BaseConn.cmd.Parameters.AddWithValue("@LpoNo", obj.objPurRtnMain.str_LpoNo)
            BaseConn.cmd.Parameters.AddWithValue("@RtnDate1", obj.objPurRtnMain.dtp_RtnDate1)
            BaseConn.cmd.Parameters.AddWithValue("@RtnDate2", obj.objPurRtnMain.dtp_RtnDate1)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.objPurRtnMain.int_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Alias", obj.objPurRtnMain.str_Alias)
            BaseConn.cmd.Parameters.AddWithValue("@Aging", obj.objPurRtnMain.int_Aging)
            BaseConn.cmd.Parameters.AddWithValue("@PayTerm", obj.objPurRtnMain.str_PayTerm)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.objPurRtnMain.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@LanguageCode", obj.objPurRtnMain.int_LanguageCode)
            BaseConn.cmd.Parameters.AddWithValue("@RtnStatus", obj.objPurRtnMain.bool_RtnStatus)
            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", obj.objPurRtnMain.str_TCCurrency)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", obj.objPurRtnMain.dbl_ExchangeRate)

            BaseConn.cmd.Parameters.AddWithValue("@TCAmount", obj.objPurRtnMain.dbl_TCAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDisAmount", obj.objPurRtnMain.dbl_TCDisAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDiscount", obj.objPurRtnMain.dbl_TCDiscount)
            BaseConn.cmd.Parameters.AddWithValue("@InvoiceTaxXML", obj.objPurRtnMain.str_InvoiceTaxXML)
            BaseConn.cmd.Parameters.AddWithValue("@TCItemTaxAmount", obj.objPurRtnMain.dbl_TCItemTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCInvoiceTaxAmount", obj.objPurRtnMain.dbl_TCInvoiceTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCAdjAmount", obj.objPurRtnMain.dbl_TCAdjAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCNetAmount", obj.objPurRtnMain.dbl_TCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCBalAmount", obj.objPurRtnMain.dbl_TCBalAmount)
            BaseConn.cmd.Parameters.AddWithValue("@LCNetAmount", obj.objPurRtnMain.dbl_LCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@StatusCancel", obj.objPurRtnMain.int_StatusCancel)
            BaseConn.cmd.Parameters.AddWithValue("@RtnVouType", obj.objPurRtnMain.str_RtnVouType)
            BaseConn.cmd.Parameters.AddWithValue("@RtnVouTypeNo", obj.objPurRtnMain.str_RtnVouTypeNo)
            BaseConn.cmd.Parameters.AddWithValue("@WHID", obj.objPurRtnMain.str_WHID)
            'BaseConn.cmd.Parameters.AddWithValue("@TaxCode", obj.objPurRtnMain.str_ItemTaxCode)
            BaseConn.cmd.Parameters.AddWithValue("@Desc1", obj.objPurRtnMain.str_Desc1)
            BaseConn.cmd.Parameters.AddWithValue("@Desc2", obj.objPurRtnMain.str_Desc2)
            BaseConn.cmd.Parameters.AddWithValue("@Desc3", obj.objPurRtnMain.str_Desc3)
            BaseConn.cmd.Parameters.AddWithValue("@Desc4", obj.objPurRtnMain.str_Desc4)
            BaseConn.cmd.Parameters.AddWithValue("@Desc5", obj.objPurRtnMain.str_Desc5)
            BaseConn.cmd.Parameters.AddWithValue("@Desc6", obj.objPurRtnMain.str_Desc6)
            BaseConn.cmd.Parameters.AddWithValue("@Desc7", obj.objPurRtnMain.str_Desc7)
            BaseConn.cmd.Parameters.AddWithValue("@Desc8", obj.objPurRtnMain.str_Desc8)

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
            BaseConn.cmd.Parameters.AddWithValue("@UserComment", obj.objPurRtnMain.str_UserComment)

            BaseConn.cmd.Parameters.AddWithValue("@PurchaseRTNPrefix", obj.objPurRtnMain.str_PurchaseRTNPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@DiscText", obj.objPurRtnMain.str_DiscText)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectID", obj.objproject.str_ProjectID)
            BaseConn.cmd.Parameters.AddWithValue("@WorkOrderNo", obj.objproject.str_WorkOrderNo)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectLocation", obj.objproject.str_ProjectLocation)

            BaseConn.cmd.Parameters.AddWithValue("@ItemTaxCode", obj.objPurRtnMain.str_ItemTaxCode)
            BaseConn.cmd.Parameters.AddWithValue("@InvoiceTaxCode", obj.objPurRtnMain.str_InvoiceTaxCode)

            BaseConn.cmd.Parameters.AddWithValue("@PRTItemDetailsDT", obj.objPurRtnSub.dt_PurRtn)
            BaseConn.cmd.Parameters.AddWithValue("@MatchingDT", obj.objPurRtnSub.dt_PRTMatching)
            BaseConn.cmd.Parameters.AddWithValue("@InvAccDetDT", obj.objPurRtnMain.dt_InvoiceAccounts)
            BaseConn.cmd.Parameters.AddWithValue("@ItemBatchDT", obj.DTBatch)
            BaseConn.cmd.Parameters.AddWithValue("@InvTaxAmountDT", obj.objPurRtnMain.dt_TaxItemDetails)

            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.objPurRtnMain.str_Flag)
            BaseConn.cmd.Parameters.Add("@PRNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutSMSMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutEmailMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output

            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            PRNo = BaseConn.cmd.Parameters("@PRNoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            _OutSMS = BaseConn.cmd.Parameters("@OutSMSMsgText").Value.ToString
            _OutEmail = BaseConn.cmd.Parameters("@OutEmailMsgText").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.objPurRtnMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "PurRtn", Err.Number, "Error in " & obj.objPurRtnMain.str_Flag & " : " & obj.objPurRtnMain.str_PRNo & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_PurRtn = _ErrString
    End Function

    Public Sub ImportPRTfromExcel(ByVal _strPath As String, ByVal _strPwd As String, ByVal _SiteID As String, ByVal _BSID As Integer, _
                                      ByVal _JVLedgerID As Integer, ByVal _PRTMainDT As DataTable, _
                             ByVal _CreatedBy As String, ByRef ErrNo As Integer, ByRef _ErrDesc As String)
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("ImportPRTfromExcel", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BSID)
            BaseConn.cmd.Parameters.AddWithValue("@JVLedgerID", _JVLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@PRTMainDT", _PRTMainDT)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", _CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 2000
            BaseConn.cmd.ExecuteNonQuery()
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            _ErrDesc = _ErrString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(_SiteID)
            ObjDalGeneral.Elog_Insert(_SiteID, _strPath, _strPwd, _BSID, _CreatedBy, Date.Now, "", "SIS", Err.Number, "Error in Import from Excel :", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
    End Sub
End Class
