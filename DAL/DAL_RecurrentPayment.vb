Imports Classes


Public Class DAL_RecurrentPayment
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _CID As Integer, ByRef Obj As csRecurrentPayment, ByRef _DTSub As DataTable)
        _DTSub = New DataTable

        BaseConn.Open(_DBPath, _DBPwd)
        BaseConn.cmd = New SqlClient.SqlCommand("[GetRecurrentPayment]", BaseConn.cnn)
        BaseConn.cmd.CommandType = CommandType.StoredProcedure
        BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
        BaseConn.cmd.Parameters.AddWithValue("@ID", Obj.ID)
        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
        Dim ds As New DataSet
        BaseConn.da.Fill(ds)

        Obj.Name = ""
        'Obj.Prefix = ds.Tables(0).Rows(0)("RevNo")
        Obj.VouPrefix = ds.Tables(0).Rows(0)("VouPrefix")
        Obj.Period = ds.Tables(0).Rows(0)("Period")
        Obj.NextDate = ds.Tables(0).Rows(0)("NextDate")
        Obj.RemainingPeriod = ds.Tables(0).Rows(0)("RemainingPeriod")
        Obj.RefNo = ds.Tables(0).Rows(0)("RefNo")
        Obj.Comment = ds.Tables(0).Rows(0)("Comment")

        If ds.Tables(1).Rows.Count > 0 Then
            _DTSub = ds.Tables(1)
        End If
    End Sub


    Public Sub UpdateRecurrentTemplate(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _CID As Integer, ByVal _BSPID As Integer,
           ByVal _MenuID As String, _ObjRP As csRecurrentPayment, ByVal _RevNo As Integer, ByVal _DTSub As DataTable, ByVal _UpdateBy As String, ByVal _UpdateDate As Date,
                                       ByRef _ErrNo As Integer, ByRef _ErrString As String)
        Dim RCP As String = String.Empty
        _ErrString = ""
        _ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("UpdateRecurrentTemplate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BSPID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _ObjRP.Flag)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", _MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Name", _ObjRP.Name)
            BaseConn.cmd.Parameters.AddWithValue("@Prefix", _ObjRP.Prefix)
            BaseConn.cmd.Parameters.AddWithValue("@VouPrefix", _ObjRP.VouPrefix)

            BaseConn.cmd.Parameters.AddWithValue("@ID", _ObjRP.ID)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", _RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@Period", _ObjRP.Period)
            BaseConn.cmd.Parameters.AddWithValue("@NextDate", _ObjRP.NextDate)
            BaseConn.cmd.Parameters.AddWithValue("@RemainingPeriod", _ObjRP.RemainingPeriod)
            BaseConn.cmd.Parameters.AddWithValue("@RefNo", _ObjRP.RefNo)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", _ObjRP.Comment)

            BaseConn.cmd.Parameters.AddWithValue("@UpdatedBy", _UpdateBy)
            BaseConn.cmd.Parameters.AddWithValue("@UpdatedDate", _UpdateDate)
            BaseConn.cmd.Parameters.AddWithValue("@RPVoucherDT", _DTSub)

            BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output

            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output

            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            RCP = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
            _RevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            _ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString

        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(_CID)
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

    End Sub

    Public Sub UpdateRecurrentPayment(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _CID As Integer, ByVal _ID As String, ByVal _UpdateBy As String, ByVal _UpdateDate As Date,
                                           ByRef _ErrNo As Integer, ByRef _ErrString As String)
        Dim RCP As String = String.Empty
        _ErrString = ""
        _ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("UpdateRecurrentPayment", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@ID", _ID)
            BaseConn.cmd.Parameters.AddWithValue("@UpdateBy", _UpdateBy)
            'BaseConn.cmd.Parameters.AddWithValue("@UpdatedDate", _UpdateDate)

            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output

            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()

            _ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString

        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(_CID)
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

    End Sub
End Class
