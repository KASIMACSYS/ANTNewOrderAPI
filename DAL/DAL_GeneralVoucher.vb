'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes
Public Class DAL_GeneralVoucher
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef Obj As csGeneralVoucher, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetGeneralVoucherDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", Obj.objGVMain.str_VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.objGVMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.objGVMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", Obj.objGVMain.str_MenuID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            Obj.objGVMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
            Obj.objGVMain.dtp_VouDate = ds.Tables(0).Rows(0)("VouDate").ToString()
            Obj.objGVMain.int_SrcLedgerID = ds.Tables(0).Rows(0)("SrcLedgerID")
            Obj.objGVMain.str_Type = ds.Tables(0).Rows(0)("Type").ToString()
            Obj.objGVMain.str_VouRef = ds.Tables(0).Rows(0)("VouRef").ToString()
            Obj.objGVMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.objGVMain.str_TCCurrency = ds.Tables(0).Rows(0)("TCCurrency").ToString()
            Obj.objGVMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()
            Obj.objGVMain.dbl_TCAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()
            Obj.objGVMain.dbl_LCAmount = ds.Tables(0).Rows(0)("LCAmount").ToString()
            Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus")
            Obj.objGVMain.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
            Obj.objGVMain.str_LedgerDepartment = ds.Tables(0).Rows(0)("LedgerDepartment").ToString()
            Obj.objGVMain.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
            Obj.objGVMain.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
            Obj.objGVMain.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
            Obj.objGVMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
            Obj.objGVMain.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
            Obj.objGVMain.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
            Obj.objGVMain.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
            Obj.objGVMain.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()
            Obj.objGVMain.int_StatusCancel = ds.Tables(0).Rows(0)("StatusCancel").ToString()
            Obj.objGVMain.dbl_TCTaxAmount = ds.Tables(0).Rows(0)("TCTaxAmount").ToString()
            Obj.DT_GVDetails = ds.Tables(1)

            If ds.Tables(2).Rows.Count > 0 Then
                Obj.objProject.str_ProjectID = ds.Tables(2).Rows(0)("ProjectID").ToString()
                Obj.objProject.str_ProjectLocation = ds.Tables(2).Rows(0)("ProjectLocation").ToString()
                Obj.objProject.str_WorkOrderNo = ds.Tables(2).Rows(0)("WorkOrderNo").ToString()
            Else
                Obj.objProject.str_ProjectID = ""
                Obj.objProject.str_ProjectLocation = ""
                Obj.objProject.str_WorkOrderNo = ""
            End If
            Obj.dt_VouMatching = ds.Tables(3)

            If ds.Tables(4).Rows.Count > 0 Then
                Obj.dt_TaxItemDetails = ds.Tables(4)
            End If
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Function Update_GenVou(ByVal obj As csGeneralVoucher, ByRef GenVouNo As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef int_RevNo As Integer, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GeneralVoucherUpdate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.objGVMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.objGVMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.objGVMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", obj.objGVMain.str_VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.objGVMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@Prefix", obj.objGVMain.str_Prefix)
            BaseConn.cmd.Parameters.AddWithValue("@Type", obj.objGVMain.str_Type)
            BaseConn.cmd.Parameters.AddWithValue("@FormType", obj.objGVMain.str_FormType)
            BaseConn.cmd.Parameters.AddWithValue("@VouRef", obj.objGVMain.str_VouRef)
            BaseConn.cmd.Parameters.AddWithValue("@SrcLedgerID", obj.objGVMain.int_SrcLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@VouDate", obj.objGVMain.dtp_VouDate)
            BaseConn.cmd.Parameters.AddWithValue("@TCAmount", obj.objGVMain.dbl_TCAmount)
            BaseConn.cmd.Parameters.AddWithValue("@LCAmount", obj.objGVMain.dbl_LCAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", obj.objGVMain.str_TCCurrency)
            BaseConn.cmd.Parameters.AddWithValue("@TCTaxAmount", obj.objGVMain.dbl_TCTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", obj.objGVMain.dbl_ExchangeRate)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.objGVMain.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerDepartment", obj.objGVMain.str_LedgerDepartment)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.bool_ApprovedStatus)

            BaseConn.cmd.Parameters.AddWithValue("@ProjectID", obj.objProject.str_ProjectID)
            BaseConn.cmd.Parameters.AddWithValue("@WorkOrderNo", obj.objProject.str_WorkOrderNo)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectLocation", obj.objProject.str_ProjectLocation)
            BaseConn.cmd.Parameters.AddWithValue("@Desc1", obj.objGVMain.str_Desc1)
            BaseConn.cmd.Parameters.AddWithValue("@Desc2", obj.objGVMain.str_Desc2)
            BaseConn.cmd.Parameters.AddWithValue("@Desc3", obj.objGVMain.str_Desc3)
            BaseConn.cmd.Parameters.AddWithValue("@Desc4", obj.objGVMain.str_Desc4)
            BaseConn.cmd.Parameters.AddWithValue("@Desc5", obj.objGVMain.str_Desc5)
            BaseConn.cmd.Parameters.AddWithValue("@Desc6", obj.objGVMain.str_Desc6)
            BaseConn.cmd.Parameters.AddWithValue("@Desc7", obj.objGVMain.str_Desc7)
            BaseConn.cmd.Parameters.AddWithValue("@Desc8", obj.objGVMain.str_Desc8)
            BaseConn.cmd.Parameters.AddWithValue("@StatusCancel", obj.objGVMain.int_StatusCancel)

            BaseConn.cmd.Parameters.AddWithValue("@GVDetailsDT", obj.DT_GVDetails)
            BaseConn.cmd.Parameters.AddWithValue("@MatchingDT", obj.dt_VouMatching)
            BaseConn.cmd.Parameters.AddWithValue("@InvTaxAmountDT", obj.dt_TaxItemDetails)

            BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            GenVouNo = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            int_RevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.int_CID)
            ObjDalGeneral.Elog_Insert(obj.int_CID, _StrDBPath, _StrDBPwd, obj.objGVMain.int_BusinessPeriodID, obj.str_CreatedBy,
                                      obj.CreatedDate, "", obj.objGVMain.str_FormType, Err.Number,
                                      "Error in " & obj.objGVMain.str_Flag & " : " & obj.objGVMain.str_VouNo & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_GenVou = _ErrString
    End Function
End Class