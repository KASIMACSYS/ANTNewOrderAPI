Imports Classes
Public Class DAL_EmpSalHike

    Dim dt As DataTable
    Dim Basecon As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Function put_structure(ByVal obj As csEmpSalHike, ByRef str_VouNo As String, ByRef intRevNo As Integer, ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByVal ErrStr As String) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            Basecon.Open(_strDBPath, _StrDBPwd)
            Basecon.cmd = New SqlClient.SqlCommand("EmpSalHikeUpdate", Basecon.cnn)
            Basecon.cmd.CommandType = CommandType.StoredProcedure
            Basecon.cmd.Parameters.AddWithValue("@CID", obj.int_CID)
            Basecon.cmd.Parameters.AddWithValue("@VouNo", obj.str_VouNo)
            Basecon.cmd.Parameters.AddWithValue("@FormPrefix", obj.str_FormPrefix)
            Basecon.cmd.Parameters.AddWithValue("@EmpHikeDT", obj.dt_EmpDocumnet)
            Basecon.cmd.Parameters.AddWithValue("@HikeDate", obj.dtp_VoucherDate)
            Basecon.cmd.Parameters.AddWithValue("@Flag", obj.str_Flag)
            Basecon.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            Basecon.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            Basecon.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            Basecon.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            Basecon.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            Basecon.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            Basecon.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.bool_ApprovedStatus)

            Basecon.cmd.Parameters.Add("@VouNoOut", SqlDbType.NVarChar, 50).Direction = ParameterDirection.Output
            Basecon.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            Basecon.cmd.Parameters.Add("@ERRORDESC", SqlDbType.NVarChar, 50).Direction = ParameterDirection.Output
            Basecon.cmd.ExecuteNonQuery()

            str_VouNo = Basecon.cmd.Parameters("@VouNoOut").Value.ToString
            ErrNo = Basecon.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = Basecon.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            ErrNo = 1
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.int_CID)
            ObjDalGeneral.Elog_Insert(obj.int_CID, _strDBPath, _StrDBPwd, "", obj.str_CreatedBy, obj.dtp_CreatedDate, "", "EMPSALARYHIKE", Err.Number, "", ex.Message, 5, 3, 1, ErrNo)
        Finally
            Basecon.Close()
        End Try
        put_structure = _ErrString
    End Function

    Public Function get_structure(ByVal obj As csEmpSalHike, ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByVal ErrNo As Integer, ByVal ErrStr As String) As csEmpSalHike
        ErrNo = 0
        Try
            Basecon.Open(_strDBPath, _StrDBPwd)
            Basecon.cmd = New SqlClient.SqlCommand("[GetEmpSalHike]", Basecon.cnn)
            Basecon.cmd.CommandType = CommandType.StoredProcedure
            Basecon.cmd.Parameters.AddWithValue("@CID", obj.int_CID)
            Basecon.cmd.Parameters.AddWithValue("@VouNo", obj.str_VouNo)
            Basecon.cmd.Parameters.AddWithValue("@LedgerID", obj.str_LedgerID)
            Basecon.cmd.Parameters.AddWithValue("@Flag", obj.str_Flag)
            Basecon.cmd.Parameters.AddWithValue("@Category", obj.str_Category)
            Basecon.cmd.Parameters.AddWithValue("@HikeDate", obj.dtp_VoucherDate)
            Basecon.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID)
            Basecon.cmd.Parameters.AddWithValue("@Tag", obj.str_Tag)
            Basecon.cmd.Parameters.AddWithValue("@Percentage", obj.dbl_Percentage)
            Basecon.da = New SqlClient.SqlDataAdapter(Basecon.cmd)
            Dim ds As New DataSet
            Basecon.da.Fill(ds)
            obj.dt_Main = ds.Tables(0)
            'If ds.Tables.Count >= 2 Then
            obj.dt_SalaryHike.Clear()
            obj.dt_SalaryHike = ds.Tables(1)
            'If ds.Tables.Count >= 3 Then
            obj.dt_SalaryArrears = ds.Tables(2)
            If ds.Tables.Count >= 4 Then
                obj.dt_Main = ds.Tables(3)
            End If
        Catch ex As Exception
            ErrNo = 1
            MsgBox(ex.Message)
        Finally
            Basecon.Close()
        End Try
        get_structure = obj
        Return get_structure
    End Function

End Class
