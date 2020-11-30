Imports Classes

Public Class DAL_EmpPaySlipPayment

    Dim dt As DataTable
    Dim Basecon As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub get_structure(ByRef obj As csEmpPaySlipPayment, ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            Basecon.Open(_strDBPath, _StrDBPwd)
            Basecon.cmd = New SqlClient.SqlCommand("[GetEmpPaySlipPayment]", Basecon.cnn)
            Basecon.cmd.CommandType = CommandType.StoredProcedure
            Basecon.cmd.Parameters.AddWithValue("@CID", obj.int_CID)
            Basecon.cmd.Parameters.AddWithValue("@Flag", obj.str_Flag)
            Basecon.cmd.Parameters.AddWithValue("@VouNo", obj.str_VoucherNo)
            Basecon.cmd.Parameters.AddWithValue("@LedgerID", obj.str_EmpID)
            Basecon.cmd.Parameters.AddWithValue("@Category", obj.str_Category)
            Basecon.da = New SqlClient.SqlDataAdapter(Basecon.cmd)
            Dim ds As New DataSet
            Basecon.da.Fill(ds)

            obj.dt_Main = ds.Tables(0)

            If ds.Tables.Count = 2 Then
                obj.str_LedgerID = ds.Tables(1).Rows(0)("LedgerID").ToString()
                obj.str_Type = ds.Tables(1).Rows(0)("Type_").ToString().Remove(0, 4)
                obj.str_Comment = ds.Tables(1).Rows(0)("Comment").ToString()
                obj.dtp_VoucherDate = ds.Tables(1).Rows(0)("Date").ToString
            End If
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message.ToLower
        Finally
            Basecon.Close()
        End Try
    End Sub

    Public Function Put_structure(ByVal obj As csEmpPaySlipPayment, ByRef _VouNo As String, ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            Basecon.Open(_strDBPath, _StrDBPwd)
            Basecon.cmd = New SqlClient.SqlCommand("EmpPaySlipPaymentUpdate", Basecon.cnn)
            Basecon.cmd.CommandType = CommandType.StoredProcedure
            Basecon.cmd.Parameters.AddWithValue("@CID", obj.int_CID) 'obj.str_SiteID
            Basecon.cmd.Parameters.AddWithValue("@DT", obj.dt_Main)
            Basecon.cmd.Parameters.AddWithValue("@Type_", obj.str_Type)
            Basecon.cmd.Parameters.AddWithValue("@VouNo", obj.str_VoucherNo)
            Basecon.cmd.Parameters.AddWithValue("@FormPrefix", obj.str_FormPrefix)
            Basecon.cmd.Parameters.AddWithValue("@MenuID", obj.str_MenuID)
            Basecon.cmd.Parameters.AddWithValue("@LedgerID", obj.str_LedgerID)
            Basecon.cmd.Parameters.AddWithValue("@RevNo", obj.int_RevNo)
            Basecon.cmd.Parameters.AddWithValue("@Flag", obj.str_Flag)
            Basecon.cmd.Parameters.AddWithValue("@VoucherDate", obj.dtp_VoucherDate)
            Basecon.cmd.Parameters.AddWithValue("@Comment", obj.str_Comment)

            Basecon.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            Basecon.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            Basecon.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            Basecon.cmd.ExecuteNonQuery()
            _VouNo = Basecon.cmd.Parameters("@VouNoOut").Value.ToString
            ErrNo = Basecon.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = Basecon.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.int_CID)
            ''ObjDalGeneral.Elog_Insert(obj.str_SiteID, _strDBPath, _StrDBPwd, obj.int_BusinessPeriodID, "", "", "EMPPAYSLIPPAYMENT", Err.Number, "", ex.Message, 5, 3, 1, ErrNo)
            ObjDalGeneral.Elog_Insert(obj.int_CID, _strDBPath, _StrDBPwd, obj.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "EmpPaySlipPayment", Err.Number, "Error in " & obj.str_Flag & " : " & obj.str_VoucherNo & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            Basecon.Close()
        End Try

        Put_structure = _ErrString

    End Function
End Class
