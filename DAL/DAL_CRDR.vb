
Imports Classes

Public Class DAL_CRDR
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByVal _DBPath As String, ByVal _DBPwd As String, ByRef Obj As csCRDR, ByRef ErrNo As Integer, ByVal ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetCrDrDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", Obj.objCRDRMain.str_VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", Obj.objCRDRMain.str_MenuID)
            BaseConn.cmd.Parameters.Add("@CashLedger", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@CashTendered", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            Obj.objCRDRMain.int_CashLedger = BaseConn.cmd.Parameters("@CashLedger").Value.ToString
            Obj.objCRDRMain.dbl_CashTendered = BaseConn.cmd.Parameters("@CashTendered").Value.ToString
            Obj.objCRDRMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()
            Obj.objCRDRMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
            Obj.objCRDRMain.str_Type = ds.Tables(0).Rows(0)("Type").ToString()
            Obj.objCRDRMain.str_SrcLedgerID = ds.Tables(0).Rows(0)("SrcLedgerID").ToString()
            Obj.objCRDRMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
            Obj.objCRDRMain.dtp_VouDate = ds.Tables(0).Rows(0)("VouDate").ToString()
            Obj.objCRDRMain.dbl_TCTotalAmount = ds.Tables(0).Rows(0)("TCAmount")
            Obj.objCRDRMain.dbl_TCItemTaxAmount = ds.Tables(0).Rows(0)("TCItemTaxAmount")
            Obj.objCRDRMain.dbl_TCInvoiceTaxAmount = ds.Tables(0).Rows(0)("TCInvTaxAmount")

            Obj.objCRDRMain.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount")
            Obj.objCRDRMain.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
            Obj.objCRDRMain.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString()
            'Obj.objCRDRMain.dbl_TCVatAmount = ds.Tables(0).Rows(0)("TCVatAmount").ToString()
            Obj.objCRDRMain.str_TCCurrency = ds.Tables(0).Rows(0)("TCCurrency").ToString()
            Obj.objCRDRMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()
            Obj.objCRDRMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.objCRDRMain.str_VouRefNo = ds.Tables(0).Rows(0)("VouRefNo").ToString()
            Obj.objCRDRMain.bool_TaxFileReturn = ds.Tables(0).Rows(0)("TaxReturnFiled")
            Obj.objCRDRMain.bool_PaymentType = ds.Tables(0).Rows(0)("PaymentType").ToString()

            Obj.objCRDRMain.str_ItemTaxCode = ds.Tables(0).Rows(0)("ItemTaxCode")
            Obj.objCRDRMain.str_InvoiceTaxCode = ds.Tables(0).Rows(0)("InvoiceTaxCode")
            Obj.objCRDRMain.str_InvoiceTaxXML = ds.Tables(0).Rows(0)("InvoiceTaxDetails")

            Obj.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
            Obj.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
            Obj.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
            Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
            Obj.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
            Obj.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
            Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus")

            If ds.Tables(1).Rows.Count > 0 Then
                Obj.objCRDRSub.dt_DstLedgerDetails = ds.Tables(1)
            End If
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub Update_CRDR(ByVal _strPath As String, ByVal _strPwd As String, ByVal obj As csCRDR, ByRef VouNo As String, ByRef intRevNo As Integer, ByRef ErrNo As Integer, ByRef _ErrString As String)
        _ErrString = ""
        ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("CrDrUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.objCRDRMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.objCRDRMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.objCRDRMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@Prefix", obj.objCRDRMain.str_Prefix)

            BaseConn.cmd.Parameters.AddWithValue("@VouNo", obj.objCRDRMain.str_VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.objCRDRMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@Type", obj.objCRDRMain.str_Type)
            BaseConn.cmd.Parameters.AddWithValue("@SrcLedgerID", obj.objCRDRMain.str_SrcLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Alias", obj.objCRDRMain.str_Alias)
            BaseConn.cmd.Parameters.AddWithValue("@VouDate", obj.objCRDRMain.dtp_VouDate)
            BaseConn.cmd.Parameters.AddWithValue("@VouRefNo", obj.objCRDRMain.str_VouRefNo)
            BaseConn.cmd.Parameters.AddWithValue("@TCAmount", obj.objCRDRMain.dbl_TCTotalAmount)
            'BaseConn.cmd.Parameters.AddWithValue("@TCDisAmount", obj.objPurInvMain.dbl_TCDisAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDiscountAmount", obj.objCRDRMain.dbl_TCDiscountAmount)

            BaseConn.cmd.Parameters.AddWithValue("@TCNetAmount", obj.objCRDRMain.dbl_TCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@LCNetAmount", obj.objCRDRMain.dbl_LCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@InvoiceTaxXML", obj.objCRDRMain.str_InvoiceTaxXML)
            BaseConn.cmd.Parameters.AddWithValue("@TCItemTaxAmount", obj.objCRDRMain.dbl_TCItemTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCInvoiceTaxAmount", obj.objCRDRMain.dbl_TCInvoiceTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@ItemTaxCode", obj.objCRDRMain.str_ItemTaxCode)
            BaseConn.cmd.Parameters.AddWithValue("@InvoiceTaxCode", obj.objCRDRMain.str_InvoiceTaxCode)

            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", obj.objCRDRMain.str_TCCurrency)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", obj.objCRDRMain.dbl_ExchangeRate)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.objCRDRMain.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@PaymentType", obj.objCRDRMain.bool_PaymentType)
            BaseConn.cmd.Parameters.AddWithValue("@CashorCredit", obj.objCRDRMain.str_CashorCredit)
            BaseConn.cmd.Parameters.AddWithValue("@CashLedger", obj.objCRDRMain.int_CashLedger)
            BaseConn.cmd.Parameters.AddWithValue("@CashTendered", obj.objCRDRMain.dbl_CashTendered)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.bool_ApprovedStatus)

            BaseConn.cmd.Parameters.AddWithValue("@ProjectID", obj.objProject.str_ProjectID)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectLocation", obj.objProject.str_ProjectLocation)
            BaseConn.cmd.Parameters.AddWithValue("@WorkOrderNo", obj.objProject.str_WorkOrderNo)

            BaseConn.cmd.Parameters.AddWithValue("@CrDrDetailsDT", obj.objCRDRSub.dt_DstLedgerDetails)
            BaseConn.cmd.Parameters.AddWithValue("@MatchingDT", obj.objCRDRSub.dt_InvMatching)
            BaseConn.cmd.Parameters.AddWithValue("@InvTaxAmountDT", obj.objCRDRMain.dt_TaxItemDetails)

            BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            VouNo = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.int_CID)
            ObjDalGeneral.Elog_Insert(obj.int_CID, _strPath, _strPwd, obj.objCRDRMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "GIP", Err.Number, "Error in " & obj.objCRDRMain.str_Flag & " : " & obj.objCRDRMain.str_VouNo & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub
End Class
