'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes

Public Class DAL_Task
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef Obj As csTask, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetTaskDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@TaskID", Obj.str_TaskID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", 101)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", Obj.str_MenuID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            Obj.str_AssetID = ds.Tables(0).Rows(0)("ID").ToString()
            Obj.str_TaskCategory = ds.Tables(0).Rows(0)("TaskCategory").ToString()
            Obj.str_Type = ds.Tables(0).Rows(0)("Type").ToString()
            Obj.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
            Obj.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
            Obj.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
            Obj.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
            Obj.dtp_StartDate = ds.Tables(0).Rows(0)("StartDate").ToString()
            Obj.dtp_DueDate = ds.Tables(0).Rows(0)("DueDate").ToString()
            Obj.dtp_PostingDate = ds.Tables(0).Rows(0)("PostDate").ToString()
            Obj.str_Status = ds.Tables(0).Rows(0)("Status").ToString()
            Obj.int_PopUpdays = ds.Tables(0).Rows(0)("PopUpDays").ToString()
            Obj.bool_NotifyFlag = ds.Tables(0).Rows(0)("NotifyFlag").ToString()
            Obj.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.dbl_Amount1 = ds.Tables(0).Rows(0)("Amount1").ToString()
            Obj.dbl_Amount2 = ds.Tables(0).Rows(0)("Amount2").ToString()
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Update_AssetTask(ByVal obj As csTask, ByRef _AssetNo As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("TaskUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.int_CID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID)

            BaseConn.cmd.Parameters.AddWithValue("@TaskID", obj.str_TaskID)
            BaseConn.cmd.Parameters.AddWithValue("@AssetID", obj.str_AssetID)
            BaseConn.cmd.Parameters.AddWithValue("@TaskCategory", obj.str_TaskCategory)
            BaseConn.cmd.Parameters.AddWithValue("@Type", obj.str_Type)
            BaseConn.cmd.Parameters.AddWithValue("@Desc1", obj.str_Desc1)
            BaseConn.cmd.Parameters.AddWithValue("@Desc2", obj.str_Desc2)
            BaseConn.cmd.Parameters.AddWithValue("@Desc3", obj.str_Desc3)
            BaseConn.cmd.Parameters.AddWithValue("@Desc4", obj.str_Desc4)
            BaseConn.cmd.Parameters.AddWithValue("@StartDate", obj.dtp_StartDate)
            BaseConn.cmd.Parameters.AddWithValue("@DueDate", obj.dtp_DueDate)
            BaseConn.cmd.Parameters.AddWithValue("@PostDate", obj.dtp_PostingDate)
            BaseConn.cmd.Parameters.AddWithValue("@Status", obj.str_Status)
            BaseConn.cmd.Parameters.AddWithValue("@PopUpDays", obj.int_PopUpdays)
            BaseConn.cmd.Parameters.AddWithValue("@NotifyFlag", obj.bool_NotifyFlag)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.str_Comment)

            BaseConn.cmd.Parameters.AddWithValue("@Amount1", obj.dbl_Amount1)
            BaseConn.cmd.Parameters.AddWithValue("@Amount2", obj.dbl_Amount2)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)

            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.str_Flag)
            BaseConn.cmd.Parameters.Add("@TaskIDOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            _AssetNo = BaseConn.cmd.Parameters("@TaskIDOut").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.int_CID)
            ObjDalGeneral.Elog_Insert(obj.int_CID, _StrDBPath, _StrDBPwd, obj.int_BusinessPeriodID, "", DateTime.Now, "", "AssetTaskMgt", Err.Number, "Error no " & obj.str_Flag & " : " & obj.str_TaskID & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        Update_AssetTask = _ErrString
    End Function
    Public Sub Get_TaskMain(ByRef Obj As csTask, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetTaskMain]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", 101)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", Obj.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Option", Obj.bool_Option)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            Obj.dt_Task = ds.Tables(0)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
End Class
