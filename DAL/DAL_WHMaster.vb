'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes
Public Class DAL_WHMaster
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef Obj As csWHMaster, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            dt = New DataTable
            BaseConn.cmd = New SqlClient.SqlCommand("[GetWareHouseDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@WHID", Obj.str_WHID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.str_Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
            Obj.str_WHDesc = dt.Rows(0)("WHDesc").ToString()
            Obj.bool_DefaultWH = dt.Rows(0)("DefaultWH").ToString()
            Obj.str_Address = dt.Rows(0)("Address").ToString()
            Obj.str_Comment = dt.Rows(0)("Comment").ToString()
            Obj.str_CreatedBy = dt.Rows(0)("CreatedBy").ToString()
            Obj.dtp_CreatedDate = dt.Rows(0)("CreatedDate").ToString()
            Obj.str_LastUpdatedBy = dt.Rows(0)("LastUpdatedBy").ToString()
            Obj.dtp_LastUpdatedDate = dt.Rows(0)("LastUpdatedDate").ToString()
            Obj.int_BusinessPeriodID = dt.Rows(0)("BusinessPeriodID").ToString()
            Obj.str_ApprovedBy = dt.Rows(0)("ApprovedBy").ToString()
            Obj.dtp_ApprovedDate = dt.Rows(0)("ApprovedDate").ToString()
            Obj.bool_ApprovedStatus = dt.Rows(0)("ApprovedStatus")
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Function Update_WHMaster(ByVal obj As csWHMaster, ByRef str_DocumentNo As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("WHMasterUpdated", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@WHID", obj.str_WHID)
            BaseConn.cmd.Parameters.AddWithValue("@WHDesc", obj.str_WHDesc)
            BaseConn.cmd.Parameters.AddWithValue("@DefaultWH", obj.bool_DefaultWH)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@Address", obj.str_Address)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.bool_ApprovedStatus)
            BaseConn.cmd.Parameters.AddWithValue("@DTItems", obj.DTItems)

            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 5000
            BaseConn.cmd.ExecuteNonQuery()
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "WareHouseMaster", Err.Number, "Error in " & obj.str_Flag & " : " & obj.str_WHID & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
        Update_WHMaster = _ErrString
    End Function

    Public Function GetWHID(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByRef _ErrNo As Integer) As Integer
        _ErrNo = 0
        GetWHID = 101

        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("GetWHID", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@WHID", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            GetWHID = BaseConn.cmd.Parameters("@WHID").Value.ToString
        Catch ex As Exception
            _ErrNo = 1
        End Try

        Return GetWHID
    End Function
End Class
