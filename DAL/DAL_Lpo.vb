'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes

Public Class DAL_Lpo
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef Obj As csLpo, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef iRC As Integer, ByRef ErrStr As String)
        iRC = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetLpoDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_CID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", Obj.objLpoMain.str_LpoNo)
            'BaseConn.cmd.Parameters.AddWithValue("@IndentNo", Obj.objLpoMain.str_IndentNo)
            'BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.objLpoMain.int_BusinessPeriodID)

            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.objLpoMain.str_Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            If Obj.objLpoMain.str_Flag = "LPO" Then
                Obj.objLpoMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()
                Obj.objLpoMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
                Obj.objLpoMain.dtp_LpoDate1 = ds.Tables(0).Rows(0)("LpoDate1").ToString()
                Obj.objLpoMain.dtp_LpoDate2 = ds.Tables(0).Rows(0)("LpoDate2").ToString()
                Obj.objLpoMain.str_ConvertFrom = ds.Tables(0).Rows(0)("ConvertFrom").ToString()
                'Obj.objLpoMain.str_EnqNo = ds.Tables(0).Rows(0)("EnqNo").ToString()
                Obj.objLpoMain.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
                Obj.objLpoMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
                Obj.objLpoMain.int_Aging = ds.Tables(0).Rows(0)("Aging").ToString()
                Obj.objLpoMain.str_PayTerm = ds.Tables(0).Rows(0)("PayTerm").ToString()
                Obj.objLpoMain.str_IndRef = ds.Tables(0).Rows(0)("IndRef").ToString()
                Obj.objLpoMain.str_DelivAddress = ds.Tables(0).Rows(0)("DelivAddress").ToString()
                Obj.objLpoMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
                Obj.objLpoMain.str_Contact = ds.Tables(0).Rows(0)("Contact").ToString()
                Obj.objLpoMain.str_RefNo = ds.Tables(0).Rows(0)("RefNo").ToString()
                Obj.objLpoMain.str_LpoStatus = ds.Tables(0).Rows(0)("LpoStatus").ToString()
                Obj.objLpoMain.int_StatusCancel = ds.Tables(0).Rows(0)("StatusCancel").ToString

                Obj.objLpoMain.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
                Obj.objLpoMain.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
                Obj.objLpoMain.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
                Obj.objLpoMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
                Obj.objLpoMain.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
                Obj.objLpoMain.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
                Obj.objLpoMain.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
                Obj.objLpoMain.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()
                'Obj.objLpoMain.str_TaxCode = ds.Tables(0).Rows(0)("TaxCode")
                Obj.objLpoMain.dbl_ItemDiscPercentage = ds.Tables(0).Rows(0)("ItemDiscPercentage").ToString()
                Obj.objLpoMain.str_ItemTaxCode = ds.Tables(0).Rows(0)("ItemTaxCode")
                Obj.objLpoMain.str_InvoiceTaxCode = ds.Tables(0).Rows(0)("InvoiceTaxCode")
                Obj.objLpoMain.str_InvoiceTaxXML = ds.Tables(0).Rows(0)("InvoiceTaxDetails").ToString()

                Obj.objLpoMain.dbl_TCAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
                Obj.objLpoMain.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
                Obj.objLpoMain.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
                Obj.objLpoMain.dbl_TCMiscPercentage = ds.Tables(0).Rows(0)("TCMiscPercentage").ToString()
                Obj.objLpoMain.dbl_TCMiscAmount = ds.Tables(0).Rows(0)("TCMiscAmount").ToString()
                Obj.objLpoMain.dbl_TCItemTaxAmount = ds.Tables(0).Rows(0)("TCItemTaxAmount")
                Obj.objLpoMain.dbl_TCInvoiceTaxAmount = ds.Tables(0).Rows(0)("TCInvTaxAmount")
                Obj.objLpoMain.dbl_TCAdjAmount = ds.Tables(0).Rows(0)("TCAdjAmount").ToString()
                Obj.objLpoMain.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
                Obj.objLpoMain.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString()
                Obj.objLpoMain.str_MiscText = ds.Tables(0).Rows(0)("MiscText").ToString()
                Obj.objLpoMain.str_DiscText = ds.Tables(0).Rows(0)("DiscText").ToString()
                Obj.objLpoMain.str_TCCurrency = ds.Tables(0).Rows(0)("TCCurrency").ToString()
                Obj.objLpoMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()
                Obj.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
                Obj.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
                Obj.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
                Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
                Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus")
                Obj.str_UserComment = ds.Tables(0).Rows(0)("UserComment").ToString()
            Else
                Obj.objLpoMain.dbl_TCAmount = 0
                Obj.objLpoMain.dbl_TCDisAmount = 0
                Obj.objLpoMain.dbl_TCDiscountAmount = 0
                Obj.objLpoMain.dbl_TCMiscAmount = 0
                Obj.objLpoMain.dbl_TCMiscPercentage = 0
                Obj.objLpoMain.dbl_TCAdjAmount = 0
                Obj.objLpoMain.dbl_TCNetAmount = 0
                Obj.objLpoMain.dbl_LCNetAmount = 0

                Obj.objLpoMain.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
                Obj.objLpoMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
                Obj.objLpoMain.int_Aging = ds.Tables(0).Rows(0)("Aging").ToString()
                Obj.objLpoMain.str_PayTerm = ds.Tables(0).Rows(0)("PayTerm").ToString()
                Obj.objLpoMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
                If Obj.objLpoMain.str_Flag = "ENQUIRY" Then
                    Obj.objLpoMain.str_DiscText = "Discount"
                    Obj.objLpoMain.str_MiscText = "Misc"
                Else
                    Obj.objLpoMain.str_DiscText = ds.Tables(0).Rows(0)("DiscText").ToString()
                End If
                'Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus").ToString()
                Obj.objLpoMain.dtp_LpoDate1 = Date.Now 'ds.Tables(0).Rows(0)("IndentDate").ToString()
                Obj.objLpoMain.dtp_LpoDate2 = Date.Now
                Obj.dtp_CreatedDate = Date.Now
                Obj.dtp_LastUpdatedDate = Date.Now
                Obj.dtp_ApprovedDate = Date.Now
                If Obj.objLpoMain.str_Flag = "INDENT" Then
                    Obj.objLpoMain.str_IndentNo = ds.Tables(0).Rows(0)("IndentNo").ToString()
                    Obj.objLpoMain.str_EnqNo = "N/A"
                    Obj.objLpoMain.dtp_IndentDate = ds.Tables(0).Rows(0)("IndentDate").ToString()
                    Obj.objLpoMain.str_ExpiryDays = ds.Tables(0).Rows(0)("ExpiryDays").ToString()
                ElseIf Obj.objLpoMain.str_Flag = "ENQUIRY" Then
                    Obj.objLpoMain.str_IndentNo = "N/A"
                    Obj.objLpoMain.str_EnqNo = ds.Tables(0).Rows(0)("EnqNo").ToString()
                End If
               
            End If

            If ds.Tables(1).Rows.Count > 0 Then
                Obj.objLpoSub.dt_Lpo = ds.Tables(1)
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

        Catch ex As Exception
            iRC = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Update_Lpo(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef LpoNo As String, ByRef intRevNo As Integer, ByVal obj As csLpo, ByRef _OutSMS As String, ByRef _OutEmail As String, ByRef ErrNo As Integer) As String
        ObjDalGeneral = New DAL_General(obj.str_CID)
        Dim JsonString As String = ObjDalGeneral.ClassToJSon(obj.objLpoMain)
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("LpoUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure

            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_CID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.objLpoMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Json", JsonString)
            BaseConn.cmd.Parameters.AddWithValue("@LpoDate1", obj.objLpoMain.dtp_LpoDate1)
            BaseConn.cmd.Parameters.AddWithValue("@LpoDate2", obj.objLpoMain.dtp_LpoDate2)
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



            BaseConn.cmd.Parameters.AddWithValue("@LpoItemDetailsDT", obj.objLpoSub.dt_Lpo)
            BaseConn.cmd.Parameters.AddWithValue("@InvTaxAmountDT", obj.objLpoMain.dt_TaxItemDetails)
            BaseConn.cmd.Parameters.Add("@LpoNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutSMSMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutEmailMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output


            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            LpoNo = BaseConn.cmd.Parameters("@LpoNoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            BaseConn.cmd.Parameters.Add("@OutSMSMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutEmailMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_CID)
            ObjDalGeneral.Elog_Insert(obj.str_CID, _StrDBPath, _StrDBPwd, obj.objLpoMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "LPO", Err.Number, "Error in " & obj.objLpoMain.str_Flag & " : " & obj.objLpoMain.str_LpoNo & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Update_Lpo = _ErrString
    End Function

End Class
