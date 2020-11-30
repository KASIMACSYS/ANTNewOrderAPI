'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Imports Classes

Public Class DAL_Mrv
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Print(ByRef Obj As csMrv, ByVal _strPath As String, ByVal _strPwd As String, ByRef iRC As Integer, ByRef ErrStr As String)
        iRC = 0
        ErrStr = ""
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetMrvDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@Mrv", Obj.objMrvMain.str_MrvNo)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.objMrvMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Obj.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.objMrvMain.str_Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            'Obj.objMrvSub.dt_MrvSub = ds.Tables(1)
            Dim dv As New DataView(ds.Tables(1))
            dv.Sort = "Slno ASC"
            Obj.objMrvSub.dt_MrvSub = dv.ToTable
            Obj.objMrvSub.dt_MrvSub.AcceptChanges()
        Catch ex As Exception
            iRC = 1
            ErrStr = ex.Message
        End Try
    End Sub


    Public Sub Get_Structure(ByVal _strPath As String, ByVal _strPwd As String, ByRef Obj As csMrv, ByRef IsInvCP As Boolean, ByRef iRC As Integer, ByRef ErrStr As String)
        iRC = 0
        ErrStr = ""
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetMrvDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@Mrv", Obj.objMrvMain.str_MrvNo)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.objMrvMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.objMrvMain.str_Flag)
            BaseConn.cmd.Parameters.Add("@IsInvCP", SqlDbType.Bit).Direction = ParameterDirection.Output
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            IsInvCP = BaseConn.cmd.Parameters("@IsInvCP").Value.ToString
            If Obj.objMrvMain.str_Flag = "Invoice" Then
                Obj.objMrvMain.int_LedgerID = ds.Tables(0).Rows(0)("SrcLedgerID").ToString()
                Obj.objMrvMain.str_PayTerm = ds.Tables(0).Rows(0)("PaymentTerm").ToString()

            Else
                Obj.objMrvMain.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
                Obj.objMrvMain.str_PayTerm = ds.Tables(0).Rows(0)("PayTerm").ToString()
            End If

            Obj.objMrvMain.str_Pin = "" ' MMF Pin NO is DONO only in Mrv table else assign value is emplty
            Obj.objMrvMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
            Obj.objMrvMain.int_Aging = ds.Tables(0).Rows(0)("Aging").ToString()

            Obj.objMrvMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.objMrvMain.str_TCCurrency = ds.Tables(0).Rows(0)("TCCurrency").ToString()
            Obj.objMrvMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()

            Obj.objMrvMain.dbl_TCAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()

            If Obj.objMrvMain.str_Flag = "Invoice" Then
                Obj.objMrvMain.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
                Obj.objMrvMain.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
                Obj.objMrvMain.str_MiscText = ds.Tables(0).Rows(0)("MiscText").ToString()
                Obj.objMrvMain.str_DiscText = ds.Tables(0).Rows(0)("DiscText").ToString()
                Obj.objMrvMain.dbl_TCItemTaxAmount = ds.Tables(0).Rows(0)("TCItemTaxAmount")
            Else
                Obj.objMrvMain.dbl_TCDisAmount = ds.Tables(0).Rows(0)("TCDisAmount").ToString()
                Obj.objMrvMain.dbl_TCDiscountAmount = ds.Tables(0).Rows(0)("TCDiscountAmount").ToString()
            End If

            Obj.objMrvMain.dbl_TCInvoiceTaxAmount = ds.Tables(0).Rows(0)("TCInvTaxAmount")
            Obj.objMrvMain.dbl_TCMiscAmount = ds.Tables(0).Rows(0)("TCMiscAmount").ToString()
            Obj.objMrvMain.dbl_TCMiscPercentage = ds.Tables(0).Rows(0)("TCMiscPercentage").ToString()
            Obj.objMrvMain.dbl_TCAdjAmount = ds.Tables(0).Rows(0)("TCAdjAmount").ToString()
            Obj.objMrvMain.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
            Obj.objMrvMain.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString() 'TODO
            Obj.objMrvMain.str_UserComment = ds.Tables(0).Rows(0)("UserComment").ToString()

            Obj.objMrvMain.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
            Obj.objMrvMain.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
            Obj.objMrvMain.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
            Obj.objMrvMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
            Obj.objMrvMain.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
            Obj.objMrvMain.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
            Obj.objMrvMain.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
            Obj.objMrvMain.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()
            Obj.objMrvMain.str_ItemTaxCode = ds.Tables(0).Rows(0)("ItemTaxCode")
            'Obj.objMrvMain.str_TaxCode = ds.Tables(0).Rows(0)("TaxCode")
            Obj.objMrvMain.str_InvoiceTaxCode = ds.Tables(0).Rows(0)("InvoiceTaxCode")
            Obj.objMrvMain.dbl_ItemDiscPercentage = ds.Tables(0).Rows(0)("ItemDiscPercentage").ToString()
            Obj.objMrvMain.dbl_TCItemTaxAmount = ds.Tables(0).Rows(0)("TCItemTaxAmount")
            Obj.objMrvMain.dbl_TCInvoiceTaxAmount = ds.Tables(0).Rows(0)("TCInvTaxAmount")
            Obj.objMrvMain.int_LanguageCode = ds.Tables(0).Rows(0)("LanguageCode")
            If Obj.objMrvMain.str_Flag = "MRV" Then
                Obj.objMrvMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()
                Obj.objMrvMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
                Obj.objMrvMain.str_LpoNo = ds.Tables(0).Rows(0)("Lpo").ToString()
                Obj.objMrvMain.dtp_MrvDate1 = ds.Tables(0).Rows(0)("MrvDate1").ToString()
                Obj.objMrvMain.dtp_MrvDate2 = ds.Tables(0).Rows(0)("MrvDate2").ToString()
                Obj.objMrvMain.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()

                Obj.objMrvMain.int_StatusCancel = ds.Tables(0).Rows(0)("StatusCancel").ToString()
                Obj.objMrvMain.str_Pin = ds.Tables(0).Rows(0)("Pin").ToString()
                Obj.objMrvMain.str_PIP = ds.Tables(0).Rows(0)("Pip").ToString()
                Obj.objMrvMain.str_PayCertComment = ds.Tables(0).Rows(0)("PaycertComment").ToString()
                Obj.objMrvMain.bool_ConvertLpo = ds.Tables(0).Rows(0)("ConvertLPO").ToString()
                Obj.objMrvMain.bool_ConvertInv = ds.Tables(0).Rows(0)("ConvertINV").ToString()
                Obj.objMrvMain.str_DeliveryAddress = ds.Tables(0).Rows(0)("DeliveryAddress").ToString()
                Obj.objMrvMain.str_ContactPerson = ds.Tables(0).Rows(0)("ContactPerson").ToString()

                Obj.objMrvMain.dbl_TCItemTaxAmount = ds.Tables(0).Rows(0)("TCItemTaxAmount").ToString()
                Obj.objMrvMain.dbl_LCLandingCost = ds.Tables(0).Rows(0)("LCLandingCost").ToString()
                Obj.objMrvMain.dbl_PIPAmount = ds.Tables(0).Rows(0)("PipAmount").ToString()
                Obj.objMrvMain.str_MiscText = ds.Tables(0).Rows(0)("MiscText").ToString()
                Obj.objMrvMain.str_DiscText = ds.Tables(0).Rows(0)("DiscText").ToString()
                Obj.objMrvMain.str_WHID = ds.Tables(0).Rows(0)("WHID").ToString()

                Obj.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
                Obj.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
                Obj.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
                Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
                Obj.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
                Obj.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
                Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus")


            ElseIf Obj.objMrvMain.str_Flag.ToUpper = "LPO" Then
                Obj.objMrvMain.str_DoNo = "N/A"
                Obj.objMrvMain.str_LpoNo = ds.Tables(0).Rows(0)("LpoNo").ToString()
                'Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus").ToString()
                Obj.objMrvMain.dtp_MrvDate1 = Date.Now
                Obj.objMrvMain.dtp_MrvDate2 = Date.Now
                Obj.objMrvMain.str_MiscText = ds.Tables(0).Rows(0)("MiscText").ToString()
                Obj.objMrvMain.str_DiscText = ds.Tables(0).Rows(0)("DiscText").ToString()
                Obj.objMrvMain.str_Pin = ds.Tables(0).Rows(0)("IndRef").ToString()
                Obj.objMrvMain.str_DeliveryAddress = ds.Tables(0).Rows(0)("DelivAddress").ToString()
                Obj.objMrvMain.str_ContactPerson = ds.Tables(0).Rows(0)("Contact").ToString()
            ElseIf Obj.objMrvMain.str_Flag.ToUpper = "INVOICE" Then
                Obj.objMrvMain.str_PIP = ds.Tables(0).Rows(0)("PIPNO").ToString()
                'Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus").ToString()
                Obj.objMrvMain.str_LpoNo = "N/A"
                Obj.objMrvMain.str_DoNo = "N/A"
                Obj.objMrvMain.dtp_MrvDate1 = Date.Now
                Obj.objMrvMain.dtp_MrvDate2 = Date.Now
                Obj.objMrvMain.str_WHID = ds.Tables(0).Rows(0)("WHID").ToString()
            End If


            Obj.objMrvSub.dt_MrvSub = ds.Tables(1)

            If ds.Tables.Count >= 3 Then

                If ds.Tables(2).Rows.Count > 0 Then
                    Obj.objproject.str_ProjectID = ds.Tables(2).Rows(0)("ProjectID").ToString()
                    Obj.objproject.str_ProjectLocation = ds.Tables(2).Rows(0)("ProjectLocation").ToString()
                    Obj.objproject.str_WorkOrderNo = ds.Tables(2).Rows(0)("WorkOrderNo").ToString()
                Else
                    Obj.objproject.str_ProjectID = ""
                    Obj.objproject.str_ProjectLocation = ""
                    Obj.objproject.str_WorkOrderNo = ""
                End If
            End If
            If ds.Tables.Count >= 4 Then
                Obj.DTBatch = ds.Tables(3)
            End If

        Catch ex As Exception
            iRC = 1
            ErrStr = ex.Message
        End Try
    End Sub

    Public Function Update_Mrv(ByVal _strPath As String, ByVal _strPwd As String, ByRef MrvNo As String, ByRef intRevNo As Integer, ByVal obj As csMrv, ByRef _OutSMS As String, ByRef _OutEmail As String, ByRef ErrNo As Integer, Optional ByVal _AffectInventory As Boolean = True) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("MrvUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.int_CID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.objMrvMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.objMrvMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.objMrvMain.str_FormPrefix)

            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.objMrvMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.objMrvMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@Mrv", obj.objMrvMain.str_MrvNo)
            BaseConn.cmd.Parameters.AddWithValue("@Lpo", obj.objMrvMain.str_LpoNo)
            BaseConn.cmd.Parameters.AddWithValue("@MrvDate1", obj.objMrvMain.dtp_MrvDate1)
            BaseConn.cmd.Parameters.AddWithValue("@MrvDate2", obj.objMrvMain.dtp_MrvDate2)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.objMrvMain.int_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Alias", obj.objMrvMain.str_Alias)

            BaseConn.cmd.Parameters.AddWithValue("@Aging", obj.objMrvMain.int_Aging)
            BaseConn.cmd.Parameters.AddWithValue("@PayTerm", obj.objMrvMain.str_PayTerm)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.objMrvMain.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@StatusInvMatched", obj.objMrvMain.bool_StatusInvMatched)
            BaseConn.cmd.Parameters.AddWithValue("@StatusCancel", obj.objMrvMain.int_StatusCancel)
            BaseConn.cmd.Parameters.AddWithValue("@Pip", obj.objMrvMain.str_PIP)
            BaseConn.cmd.Parameters.AddWithValue("@Pin", obj.objMrvMain.str_Pin)
            BaseConn.cmd.Parameters.AddWithValue("@PaycertComment", obj.objMrvMain.str_PayCertComment)
            BaseConn.cmd.Parameters.AddWithValue("@ConvertLPO", obj.objMrvMain.bool_ConvertLpo)
            BaseConn.cmd.Parameters.AddWithValue("@ConvertINV", obj.objMrvMain.bool_ConvertInv)
            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", obj.objMrvMain.str_TCCurrency)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", obj.objMrvMain.dbl_ExchangeRate)
            BaseConn.cmd.Parameters.AddWithValue("@MiscText", obj.objMrvMain.str_MiscText)
            BaseConn.cmd.Parameters.AddWithValue("@DiscText", obj.objMrvMain.str_DiscText)
            BaseConn.cmd.Parameters.AddWithValue("@TCAmount", obj.objMrvMain.dbl_TCAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDisAmount", obj.objMrvMain.dbl_TCDisAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCDiscountAmount", obj.objMrvMain.dbl_TCDiscountAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCMiscPercentage", obj.objMrvMain.dbl_TCMiscPercentage)
            BaseConn.cmd.Parameters.AddWithValue("@TCMiscAmount", obj.objMrvMain.dbl_TCMiscAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCAdjAmount", obj.objMrvMain.dbl_TCAdjAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCNetAmount", obj.objMrvMain.dbl_TCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@LCNetAmount", obj.objMrvMain.dbl_LCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@LCLandingCost", obj.objMrvMain.dbl_LCLandingCost)
            BaseConn.cmd.Parameters.AddWithValue("@PipAmount", obj.objMrvMain.dbl_PIPAmount)

            BaseConn.cmd.Parameters.AddWithValue("@WHID", obj.objMrvMain.str_WHID)
            BaseConn.cmd.Parameters.AddWithValue("@UserComment", obj.objMrvMain.str_UserComment)
            BaseConn.cmd.Parameters.AddWithValue("@DeliveryAddress", obj.objMrvMain.str_DeliveryAddress)
            BaseConn.cmd.Parameters.AddWithValue("@ContactPerson", obj.objMrvMain.str_ContactPerson)
            BaseConn.cmd.Parameters.AddWithValue("@Desc1", obj.objMrvMain.str_Desc1)
            BaseConn.cmd.Parameters.AddWithValue("@Desc2", obj.objMrvMain.str_Desc2)
            BaseConn.cmd.Parameters.AddWithValue("@Desc3", obj.objMrvMain.str_Desc3)
            BaseConn.cmd.Parameters.AddWithValue("@Desc4", obj.objMrvMain.str_Desc4)
            BaseConn.cmd.Parameters.AddWithValue("@Desc5", obj.objMrvMain.str_Desc5)
            BaseConn.cmd.Parameters.AddWithValue("@Desc6", obj.objMrvMain.str_Desc6)
            BaseConn.cmd.Parameters.AddWithValue("@Desc7", obj.objMrvMain.str_Desc7)
            BaseConn.cmd.Parameters.AddWithValue("@Desc8", obj.objMrvMain.str_Desc8)
            BaseConn.cmd.Parameters.AddWithValue("@ItemDiscPercentage", obj.objMrvMain.dbl_ItemDiscPercentage)
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
            BaseConn.cmd.Parameters.AddWithValue("@AffectInventory", _AffectInventory)
            BaseConn.cmd.Parameters.AddWithValue("@ItemTaxCode", obj.objMrvMain.str_ItemTaxCode)
            BaseConn.cmd.Parameters.AddWithValue("@InvoiceTaxCode", obj.objMrvMain.str_InvoiceTaxCode)
            BaseConn.cmd.Parameters.AddWithValue("@TCItemTaxAmount", obj.objMrvMain.dbl_TCItemTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCInvoiceTaxAmount", obj.objMrvMain.dbl_TCInvoiceTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@InvoiceTaxXML", obj.objMrvMain.str_InvoiceTaxXML)
            BaseConn.cmd.Parameters.AddWithValue("@LanguageCode", obj.objMrvMain.int_LanguageCode)

            BaseConn.cmd.Parameters.AddWithValue("@MrvItemDetailsDT", obj.objMrvSub.dt_MrvSub)
            BaseConn.cmd.Parameters.AddWithValue("@InvTaxAmountDT", obj.objMrvMain.dt_TaxItemDetails)
            BaseConn.cmd.Parameters.AddWithValue("@ItemBatchDT", obj.DTBatch)

            BaseConn.cmd.Parameters.Add("@MrvNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutSMSMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutEmailMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output

            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            MrvNo = BaseConn.cmd.Parameters("@MrvNoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            _OutSMS = BaseConn.cmd.Parameters("@OutSMSMsgText").Value.ToString
            _OutEmail = BaseConn.cmd.Parameters("@OutEmailMsgText").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.int_CID)
            ObjDalGeneral.Elog_Insert(obj.int_CID, _strPath, _strPwd, obj.objMrvMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "MRV", ErrNo, "Error in " & obj.objMrvMain.str_Flag & " : " & obj.objMrvMain.str_MrvNo & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_Mrv = _ErrString
    End Function

End Class
