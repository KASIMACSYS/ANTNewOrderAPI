'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Imports Classes

Public Class DAL_Indent
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef Obj As csIndent, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetIndentDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@IndentNo", Obj.ObjIndentMain.str_IndentNo)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.int_CID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            Obj.ObjIndentMain.dtp_IndentDate1 = ds.Tables(0).Rows(0)("IndentDate1").ToString()
            Obj.ObjIndentMain.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
            Obj.ObjIndentMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()

            Obj.ObjIndentMain.int_Aging = ds.Tables(0).Rows(0)("Aging").ToString()
            Obj.ObjIndentMain.str_PayTerm = ds.Tables(0).Rows(0)("PayTerm").ToString()

            Obj.ObjIndentMain.str_IndentStatus = ds.Tables(0).Rows(0)("IndentStatus").ToString()
            Obj.ObjIndentMain.dtp_IndentDate2 = ds.Tables(0).Rows(0)("IndentDate2").ToString()
            Obj.ObjIndentMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
            Obj.ObjIndentMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
            Obj.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
            Obj.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
            Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
            Obj.ObjIndentMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()
            Obj.ObjIndentMain.int_StatusCancel = ds.Tables(0).Rows(0)("StatusCancel").ToString()

            Obj.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
            Obj.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
            Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus")
            Obj.ObjIndentMain.str_UserComment = ds.Tables(0).Rows(0)("UserComment").ToString()

            Obj.ObjIndentMain.dbl_TCAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
            Obj.ObjIndentMain.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
            Obj.ObjIndentMain.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
            Obj.ObjIndentMain.dbl_TCMiscPercentage = ds.Tables(0).Rows(0)("TCMiscPercentage").ToString()
            Obj.ObjIndentMain.dbl_TCMiscAmount = ds.Tables(0).Rows(0)("TCMiscAmount").ToString()
            Obj.ObjIndentMain.dbl_TCAdjAmount = ds.Tables(0).Rows(0)("TCAdjAmount").ToString()
            Obj.ObjIndentMain.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
            Obj.ObjIndentMain.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString()
            Obj.ObjIndentMain.dbl_TCItemTaxAmount = ds.Tables(0).Rows(0)("TCItemTaxAmount")
            Obj.ObjIndentMain.dbl_TCInvoiceTaxAmount = ds.Tables(0).Rows(0)("TCInvTaxAmount")

            Obj.ObjIndentMain.str_TCCurrency = ds.Tables(0).Rows(0)("TCCurrency").ToString()
            Obj.ObjIndentMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()

            Obj.ObjIndentMain.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
            Obj.ObjIndentMain.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
            Obj.ObjIndentMain.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
            Obj.ObjIndentMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
            Obj.ObjIndentMain.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
            Obj.ObjIndentMain.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
            Obj.ObjIndentMain.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
            Obj.ObjIndentMain.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()
            Obj.ObjIndentMain.str_MiscText = ds.Tables(0).Rows(0)("MiscText").ToString()
            Obj.ObjIndentMain.str_DiscText = ds.Tables(0).Rows(0)("DiscText").ToString()
            Obj.ObjIndentMain.str_ExpiryDays = ds.Tables(0).Rows(0)("ExpiryDays").ToString()
            Obj.ObjIndentMain.str_ItemTaxCode = ds.Tables(0).Rows(0)("ItemTaxCode")
            Obj.ObjIndentMain.str_InvoiceTaxCode = ds.Tables(0).Rows(0)("InvoiceTaxCode")
            'Obj.ObjIndentMain.str_PermitNo = ds.Tables(0).Rows(0)("PermitNo").ToString()
            'Obj.ObjIndentMain.dbl_ItemDiscPercentage = ds.Tables(0).Rows(0)("ItemDiscPercentage").ToString()
            Obj.ObjIndentMain.str_InvoiceTaxXML = ds.Tables(0).Rows(0)("InvoiceTaxDetails")
            Obj.ObjIndentSub.dt_IndentSub = ds.Tables(1) 'To Grid

            If ds.Tables(2).Rows.Count > 0 Then
                Obj.objProject.str_ProjectID = ds.Tables(2).Rows(0)("ProjectID").ToString()
                Obj.objProject.str_ProjectLocation = ds.Tables(2).Rows(0)("ProjectLocation").ToString()
                Obj.objProject.str_WorkOrderNo = ds.Tables(2).Rows(0)("WorkOrderNo").ToString()
            Else
                Obj.objProject.str_ProjectID = ""
                Obj.objProject.str_ProjectLocation = ""
                Obj.objProject.str_WorkOrderNo = ""
            End If

            Obj.ObjIndentMain.dt_TaxItemDetails = ds.Tables(3)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub


    Public Function Update_Indent(ByVal _StrDBPath As String, ByVal _strDBPwd As String, ByRef IndentNo As String, ByRef _IntRevNo As Integer, ByVal obj As csIndent, ByRef _OutSMS As String, ByRef _OutEmail As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _strDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("IndentUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.int_CID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.ObjIndentMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjIndentMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.ObjIndentMain.str_Prefix)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.ObjIndentMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@IndentNo", obj.ObjIndentMain.str_IndentNo)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.ObjIndentMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@IndentDate1", obj.ObjIndentMain.dtp_IndentDate1)
            BaseConn.cmd.Parameters.AddWithValue("@IndentDate2", obj.ObjIndentMain.dtp_IndentDate2)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.ObjIndentMain.int_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Alias", obj.ObjIndentMain.str_Alias)

            BaseConn.cmd.Parameters.AddWithValue("@Aging", obj.ObjIndentMain.int_Aging)
            BaseConn.cmd.Parameters.AddWithValue("@PayTerm", obj.ObjIndentMain.str_PayTerm)

            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.ObjIndentMain.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@IndentStatus", obj.ObjIndentMain.str_IndentStatus)

            BaseConn.cmd.Parameters.AddWithValue("@TCAmount", obj.ObjIndentMain.dbl_TCAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDisAmount", obj.ObjIndentMain.dbl_TCDisAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDiscountAmount", obj.ObjIndentMain.dbl_TCDiscountAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCMiscPercentage", obj.ObjIndentMain.dbl_TCMiscPercentage)
            BaseConn.cmd.Parameters.AddWithValue("@TCMiscAmount", obj.ObjIndentMain.dbl_TCMiscAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCNetAmount", obj.ObjIndentMain.dbl_TCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@LCNetAmount", obj.ObjIndentMain.dbl_LCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", obj.ObjIndentMain.str_TCCurrency)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", obj.ObjIndentMain.dbl_ExchangeRate)
            BaseConn.cmd.Parameters.AddWithValue("@TCAdjAmount", obj.ObjIndentMain.dbl_TCAdjAmount)

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
            BaseConn.cmd.Parameters.AddWithValue("@LanguageCode", obj.ObjIndentMain.int_LanguageCode)

            BaseConn.cmd.Parameters.AddWithValue("@StatusCancel", obj.ObjIndentMain.int_StatusCancel)

            BaseConn.cmd.Parameters.AddWithValue("@ProjectID", obj.objProject.str_ProjectID)
            BaseConn.cmd.Parameters.AddWithValue("@WorkOrderNo", obj.objProject.str_WorkOrderNo)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectLocation", obj.objProject.str_ProjectLocation)
            BaseConn.cmd.Parameters.AddWithValue("@UserComment", obj.ObjIndentMain.str_UserComment)
            BaseConn.cmd.Parameters.AddWithValue("@DiscText", obj.ObjIndentMain.str_DiscText)
            BaseConn.cmd.Parameters.AddWithValue("@MiscText", obj.ObjIndentMain.str_MiscText)
            ''AM Specific
            BaseConn.cmd.Parameters.AddWithValue("@Desc1", obj.ObjIndentMain.str_Desc1)
            BaseConn.cmd.Parameters.AddWithValue("@Desc2", obj.ObjIndentMain.str_Desc2)
            BaseConn.cmd.Parameters.AddWithValue("@Desc3", obj.ObjIndentMain.str_Desc3)
            BaseConn.cmd.Parameters.AddWithValue("@Desc4", obj.ObjIndentMain.str_Desc4)
            BaseConn.cmd.Parameters.AddWithValue("@Desc5", obj.ObjIndentMain.str_Desc5)
            BaseConn.cmd.Parameters.AddWithValue("@Desc6", obj.ObjIndentMain.str_Desc6)
            BaseConn.cmd.Parameters.AddWithValue("@Desc7", obj.ObjIndentMain.str_Desc7)
            BaseConn.cmd.Parameters.AddWithValue("@Desc8", obj.ObjIndentMain.str_Desc8)
            BaseConn.cmd.Parameters.AddWithValue("@ExpiryDays", obj.ObjIndentMain.str_ExpiryDays)


            BaseConn.cmd.Parameters.AddWithValue("@ItemTaxCode", obj.ObjIndentMain.str_ItemTaxCode)
            BaseConn.cmd.Parameters.AddWithValue("@InvoiceTaxCode", obj.ObjIndentMain.str_InvoiceTaxCode)
            BaseConn.cmd.Parameters.AddWithValue("@TCItemTaxAmount", obj.ObjIndentMain.dbl_TCItemTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCInvoiceTaxAmount", obj.ObjIndentMain.dbl_TCInvoiceTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@InvoiceTaxXML", obj.ObjIndentMain.str_InvoiceTaxXML)

            BaseConn.cmd.Parameters.AddWithValue("@IndentItemDetailsDT", obj.ObjIndentSub.dt_IndentSub)
            BaseConn.cmd.Parameters.AddWithValue("@InvTaxAmountDT", obj.ObjIndentMain.dt_TaxItemDetails)


            BaseConn.cmd.Parameters.Add("@IndentNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output

            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutSMSMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutEmailMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output

            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()

            IndentNo = BaseConn.cmd.Parameters("@IndentNoOut").Value.ToString
            _IntRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString

            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString

            _OutSMS = BaseConn.cmd.Parameters("@OutSMSMsgText").Value.ToString
            _OutEmail = BaseConn.cmd.Parameters("@OutEmailMsgText").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.int_CID)
            ObjDalGeneral.Elog_Insert(obj.int_CID, _StrDBPath, _strDBPwd, obj.ObjIndentMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "Indent", Err.Number, "Error in '" & obj.ObjIndentMain.str_Flag & "'ED '" & obj.ObjIndentMain.str_IndentNo & "' ", ex.Message, 5, 3, 1, 0)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Update_Indent = _ErrString
    End Function
End Class
