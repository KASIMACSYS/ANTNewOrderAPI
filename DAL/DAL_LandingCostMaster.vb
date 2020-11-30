Imports Classes
Public Class DAL_LandingCostMaster
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General
    Public Function Put_Structure(ByVal obj As csLandingCostMaster, ByRef LCID As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("LandingCostMasterUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID) 'obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@LCID", obj.str_LCID)
            BaseConn.cmd.Parameters.AddWithValue("@VendorLedgerID", obj.int_VendorLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@DefDistributionType", obj.str_DefDistributionType)
            BaseConn.cmd.Parameters.AddWithValue("@GenExpLedgerID", obj.int_GenExpLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@Active", obj.bool_Active)


            BaseConn.cmd.Parameters.Add("@LCIDOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            LCID = BaseConn.cmd.Parameters("@LCIDOut").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.int_BusinessPeriodID, "", "", "", "LandingCostMaster", Err.Number, "Error in " & obj.str_Flag & " : " & obj.str_LCID & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
        Put_Structure = _ErrString
    End Function
    Public Sub Get_Structure(ByRef Obj As Classes.csLandingCostMaster, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal LngCode As Integer, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetLandingCostMasterDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@LCID", Obj.str_LCID)
            BaseConn.cmd.Parameters.AddWithValue("@LngCode", LngCode)

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            'If ds.Tables(0).Rows.Count > 0 Then
            Obj.dt = ds.Tables(0)
            If ds.Tables.Count = 2 Then
                Obj.str_LCID = ds.Tables(1).Rows(0)("LCID").ToString()

                Obj.int_VendorLedgerID = ds.Tables(1).Rows(0)("VendorLedgerID").ToString()
                Obj.str_DefDistributionType = ds.Tables(1).Rows(0)("DefDistributionType").ToString()
                Obj.int_GenExpLedgerID = ds.Tables(1).Rows(0)("GenExpLedgerID").ToString()
                Obj.str_Comment = ds.Tables(1).Rows(0)("Comment").ToString()
                Obj.bool_Active = ds.Tables(1).Rows(0)("Active").ToString()
            End If


            'Obj.dt = ds.Tables(1)
            ' End If


        Catch ex As Exception

            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try

    End Sub
End Class
