Imports Classes

Public Class DAL_TaxGroup
    Private ObjDalGeneral As DAL_General
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Public Sub Get_Structure(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef Obj As csTaxGroup, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetTaxGroup]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@TaxGroupName", Obj.ObjTaxGroupMain.str_TaxGroupName)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.ObjTaxGroupMain.str_Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            Obj.ObjTaxGroupMain.int_TaxGroupID = ds.Tables(0).Rows(0)("TaxGroupID").ToString()
            Obj.ObjTaxGroupMain.str_Description = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.ObjTaxGroupSub.dt_TaxGroupSub = ds.Tables(1)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Function GetTaxCodeDetails(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _Flag As String, ByVal dt_TaxLedgerID As DataTable, ByRef ErrNo As Integer) As DataTable
        Dim _ErrString As String = ""
        ErrNo = 0
        dt = New DataTable
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetTaxGroup]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@TaxGroupName", 0)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = ds.Tables(0)
        Catch ex As Exception
            _ErrString = ex.Message
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
    Public Function Update_TaxGroup(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal obj As csTaxGroup, ByRef Int_TaxID As Integer, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[TaxGroupUpdate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjTaxGroupMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@TaxGroupID", obj.ObjTaxGroupMain.int_TaxGroupID)
            BaseConn.cmd.Parameters.AddWithValue("@TaxGroupName", obj.ObjTaxGroupMain.str_TaxGroupName)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.ObjTaxGroupMain.str_Description)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.ObjTaxGroupMain.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.ObjTaxGroupMain.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.ObjTaxGroupMain.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.ObjTaxGroupMain.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@TaxGroupDetailsDT", obj.ObjTaxGroupSub.dt_TaxGroupSub)
            BaseConn.cmd.Parameters.Add("@TaxGroupIDOut", SqlDbType.Int, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            Int_TaxID = BaseConn.cmd.Parameters("@TaxGroupIDOut").Value
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.int_BusinessPeriodID, obj.ObjTaxGroupMain.str_CreatedBy, obj.ObjTaxGroupMain.dtp_CreatedDate, "", "TaxGroup", Err.Number, "Error in " & obj.ObjTaxGroupMain.str_Flag & " : " & obj.ObjTaxGroupMain.int_TaxGroupID & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        Update_TaxGroup = _ErrString
    End Function
End Class
