Imports Classes
Public Class DAL_DefaultLedger
    Dim dt_ As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General



    Public Function Update_DefaultLedger(ByRef obj As csDefaultLedger, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[DropdownList]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPerionID)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", "DefaultLedger")
            BaseConn.cmd.Parameters.AddWithValue("@Condition", obj.str_Types)
            BaseConn.cmd.Parameters.AddWithValue("@DefaultLedgerDT", obj.dt_DefaultLedger)
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
        Catch ex As Exception
            ErrNo = 1
            _ErrString = ex.Message

        Finally
            BaseConn.Close()
        End Try
        Update_DefaultLedger = _ErrString
    End Function
    Public Sub Get_Structure(ByRef Obj As csDefaultLedger, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetDefaultLedger]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            'BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPerionID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", Obj.str_Types)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            Obj.dt_DefaultLedger = ds.Tables(0)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
End Class
