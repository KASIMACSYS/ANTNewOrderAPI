'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes
Public Class DAL_JV
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef Obj As csJV, ByRef iRC As Integer, ByRef ErrStr As String)
        iRC = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetJVDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.ObjJVMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@JVNo", Obj.ObjJVMain.str_JVNo)


            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            Obj.ObjJVMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
            Obj.ObjJVMain.dtp_JVDate = ds.Tables(0).Rows(0)("JVDate").ToString()

            Obj.ObjJVMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.ObjJVMain.str_TCCurrency = ds.Tables(0).Rows(0)("TCCurrency").ToString()
            Obj.ObjJVMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()

            Obj.ObjJVMain.dbl_TCNetAmount = ds.Tables(0).Rows(0)("TCNetAmount").ToString()
            Obj.ObjJVMain.dbl_LCNetAmount = ds.Tables(0).Rows(0)("LCNetAmount").ToString()
            Obj.ObjJVMain.dbl_TCTaxAmount = ds.Tables(0).Rows(0)("TCTaxAmount").ToString

            Obj.ObjJVMain.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus")
            Obj.ObjJVMain.str_UserComment = ds.Tables(0).Rows(0)("UserComment").ToString()
            Obj.ObjJVMain.int_StatusCancel = ds.Tables(0).Rows(0)("StatusCancel")
            Obj.ObjJVMain.str_BatchID = ds.Tables(0).Rows(0)("BatchID")
            Obj.ObjJVMain.str_RefNo = ds.Tables(0).Rows(0)("RefNo").ToString

            Obj.ObjJVSub.dt_JVSub = ds.Tables(1)
            Obj.ObjJVSub.dt_JVMatching = ds.Tables(2)

            If ds.Tables(3).Rows.Count > 0 Then
                Obj.objproject.str_ProjectID = ds.Tables(3).Rows(0)("ProjectID").ToString()
                Obj.objproject.str_ProjectLocation = ds.Tables(3).Rows(0)("ProjectLocation").ToString()
                Obj.objproject.str_WorkOrderNo = ds.Tables(3).Rows(0)("WorkOrderNo").ToString()
            Else
                Obj.objproject.str_ProjectID = ""
                Obj.objproject.str_ProjectLocation = ""
                Obj.objproject.str_WorkOrderNo = ""
            End If

            If ds.Tables(4).Rows.Count > 0 Then
                Obj.ObjJVSub.dt_TaxItemDetails = ds.Tables(4)
            End If

        Catch ex As Exception
            iRC = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Function Update_JV(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef _JVNo As String, ByRef _RevNo As Integer, ByVal _obj As csJV, ByRef _OutSMS As String, ByRef _OutEmail As String, ByRef _ErrNo As Integer,
                              ByRef _ErrString As String) As String
        _ErrString = ""
        _ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("JVUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _obj.ObjJVMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", _obj.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", _obj.ObjJVMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _obj.ObjJVMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", _obj.ObjJVMain.str_Prefix)

            BaseConn.cmd.Parameters.AddWithValue("@JVNo", _obj.ObjJVMain.str_JVNo)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", _obj.ObjJVMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@JVDate", _obj.ObjJVMain.dtp_JVDate)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", _obj.ObjJVMain.str_Comment)

            BaseConn.cmd.Parameters.AddWithValue("@TCNetAmount", _obj.ObjJVMain.dbl_TCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@LCNetAmount", _obj.ObjJVMain.dbl_LCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCTaxAmount", _obj.ObjJVMain.dbl_TCTaxAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", _obj.ObjJVMain.str_TCCurrency)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", _obj.ObjJVMain.dbl_ExchangeRate)


            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", _obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", _obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", _obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", _obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", _obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", _obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", _obj.bool_ApprovedStatus)
            BaseConn.cmd.Parameters.AddWithValue("@UserComment", _obj.ObjJVMain.str_UserComment)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedLevel", _obj.ApprovedLevel)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedComment", _obj.ApprovedComment)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedHigherLevel", _obj.ApprovedHigherLevel)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectID", _obj.objproject.str_ProjectID)
            BaseConn.cmd.Parameters.AddWithValue("@WorkOrderNo", _obj.objproject.str_WorkOrderNo)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectLocation", _obj.objproject.str_ProjectLocation)
            BaseConn.cmd.Parameters.AddWithValue("@StatusCancel", _obj.ObjJVMain.int_StatusCancel)
            BaseConn.cmd.Parameters.AddWithValue("@BatchID", _obj.ObjJVMain.str_BatchID)
            BaseConn.cmd.Parameters.AddWithValue("@LanguageCode", _obj.ObjJVMain.int_LanguageCode)
            BaseConn.cmd.Parameters.AddWithValue("@RefNo", _obj.ObjJVMain.str_RefNo)


            BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output

            BaseConn.cmd.Parameters.AddWithValue("@JVSubDT", _obj.ObjJVSub.dt_JVSub)

            BaseConn.cmd.Parameters.AddWithValue("@AdvanceDT", _obj.ObjJVSub.dt_Wages)
            BaseConn.cmd.Parameters.AddWithValue("@InvTaxAmountDT", _obj.ObjJVSub.dt_TaxItemDetails)

            BaseConn.cmd.Parameters.AddWithValue("@MatchingDT", _obj.ObjJVSub.dt_JVMatching)

            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutSMSMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutEmailMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output

            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            _JVNo = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
            _ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            _RevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            _OutSMS = BaseConn.cmd.Parameters("@OutSMSMsgText").Value.ToString
            _OutEmail = BaseConn.cmd.Parameters("@OutEmailMsgText").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(_obj.int_CID)
            ObjDalGeneral.Elog_Insert(_obj.int_CID, _StrDBPath, _StrDBPwd, _obj.ObjJVMain.int_BusinessPeriodID, _obj.str_CreatedBy, _obj.dtp_CreatedDate, "",
                                      "JV", Err.Number, "Error in " & _obj.ObjJVMain.str_Flag & " : " & _obj.ObjJVMain.str_JVNo & "", ex.Message, 5, 3, 1, _ErrNo)
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_JV = _ErrString
    End Function


End Class
