'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Imports Classes
Public Class DAL_LedgerMaster
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef Obj As csLedgerMaster, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef _ErrNo As Integer, ByRef _ErrString As String)
        _ErrNo = 0
        _ErrString = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            dt = New DataTable
            BaseConn.cmd = New SqlClient.SqlCommand("[GetLedgerMasterDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", Obj.int_LedgerID)
            'BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            'BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.str_Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
            Obj.int_LedgerID = dt.Rows(0)("LedgerID").ToString()
            Obj.str_Class = dt.Rows(0)("LedgerType").ToString()
            Obj.str_Classification = dt.Rows(0)("SubClass").ToString()
            'Obj.str_Catagory = dt.Rows(0)("Catagory").ToString()
            'Obj.str_COA = dt.Rows(0)("COA").ToString()
            Obj.str_Description = dt.Rows(0)("Description").ToString()
            Obj.str_LedgerType = dt.Rows(0)("LedgerType").ToString()
            'Obj.dbl_Amount = dt.Rows(0)("Amount").ToString()
            'Obj.dbl_Advance = dt.Rows(0)("Advance").ToString()
            Obj.str_StartRange = dt.Rows(0)("StartRange").ToString()
            Obj.str_EndRange = dt.Rows(0)("EndRange").ToString()
            Obj.str_AccountNo1 = dt.Rows(0)("AccountCode1").ToString()
            Obj.str_AccountNo2 = dt.Rows(0)("AccountCode2").ToString()
            Obj.bool_InActive = dt.Rows(0)("Status").ToString()
            Obj.bool_CostCentre = dt.Rows(0)("CostCentre").ToString()

            Obj.str_Comment = dt.Rows(0)("Comment").ToString()
            Obj.str_ParentAccount = dt.Rows(0)("ParentAccount").ToString()
            Obj.bool_Readonly = dt.Rows(0)("ReadOnly").ToString()
        Catch ex As Exception
            _ErrNo = 1
            _ErrString = ex.Message.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Function Update_LedgerMaster(ByVal obj As csLedgerMaster, ByRef str_COA As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef LedgerDetails As String) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        LedgerDetails = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("LedgerMasterUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID)

            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.str_Flag)

            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.int_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Class", obj.str_Class)
            BaseConn.cmd.Parameters.AddWithValue("@Classification", obj.str_Classification)
            BaseConn.cmd.Parameters.AddWithValue("@StartRange", obj.str_StartRange)
            BaseConn.cmd.Parameters.AddWithValue("@EndRange", obj.str_EndRange)
            BaseConn.cmd.Parameters.AddWithValue("@AccountCode1", obj.str_AccountNo1)
            BaseConn.cmd.Parameters.AddWithValue("@AccountCode2", obj.str_AccountNo2)

            ''BaseConn.cmd.Parameters.AddWithValue("@Catagory", obj.str_Catagory)
            ''BaseConn.cmd.Parameters.AddWithValue("@COA", obj.str_COA)
            BaseConn.cmd.Parameters.AddWithValue("@Description", obj.str_Description)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerType", obj.str_LedgerType)
            ''BaseConn.cmd.Parameters.AddWithValue("@Amount", obj.dbl_Amount)
            ''BaseConn.cmd.Parameters.AddWithValue("@Advance", obj.dbl_Advance)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@ParentAccount", obj.str_ParentAccount)
            BaseConn.cmd.Parameters.AddWithValue("@InActive", obj.bool_InActive)
            BaseConn.cmd.Parameters.AddWithValue("@CostCentre", obj.bool_CostCentre)
            BaseConn.cmd.Parameters.AddWithValue("@ReadOnly", obj.bool_Readonly)
            BaseConn.cmd.Parameters.AddWithValue("@Category", obj.str_Category)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)

            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@LedgerIDOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@LedgerDetails", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            str_COA = BaseConn.cmd.Parameters("@LedgerIDOut").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            LedgerDetails = BaseConn.cmd.Parameters("@LedgerDetails").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ErrNo = 1
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "LedgerMaster", ErrNo, obj.str_Flag, ex.Message, 5, 3, 1, ErrNo)
        Finally
            BaseConn.Close()
        End Try
        Update_LedgerMaster = _ErrString
    End Function
    Public Function Import_LedgerMaster(ByVal str_SiteID As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal dt As DataTable, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("LedgerMasterImport", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@DT", dt)
            BaseConn.cmd.ExecuteNonQuery()
        Catch ex As Exception
            _ErrString = ex.Message
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Import_LedgerMaster = _ErrString
    End Function

    Public Sub GetParentIDDetails(ByVal str_SiteID As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _Category As String, ByRef _dt As DataTable, ByRef ErrNo As Integer)
        Dim _ErrString As String = ""
        ErrNo = 0
        _dt = New DataTable
        Try
            Dim ds As New DataSet
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("GetTopParentID", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Category", _Category)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(ds)
            _dt = ds.Tables(0)
        Catch ex As Exception
            _ErrString = ex.Message
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub

End Class
