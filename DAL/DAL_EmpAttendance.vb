Imports Classes

Public Class DAL_EmpAttendance
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub GetEmployeeForAttendance(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal int_CID As Integer, ByVal _Flag As String, ByVal _LedgerID As String, ByVal _Category As String,
                                        ByVal _FromDate As Date, ByRef _DTEmployee As DataTable, ByRef _DTDBINOUT As DataTable, ByRef iRC As Integer, ByRef ErrStr As String)
        iRC = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetEmpForAtt]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Category", _Category)
            BaseConn.cmd.Parameters.AddWithValue("@GivenDate", _FromDate)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _DTEmployee = ds.Tables(0)
            _DTDBINOUT = ds.Tables(1)
        Catch ex As Exception
            iRC = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub EmployeeAttendanceUpdate(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal int_CID As Integer, ByVal _Flag As String,
                                      ByVal _AttDate As Date, ByVal _Category As String, ByVal _DTAttMain As DataTable, ByVal _DTAttSub As DataTable,
                                      ByVal _UserID As String, ByRef iRC As Integer, ByRef _ErrString As String)
        iRC = 0
        'Dim _ErrString As String = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[EmpForAttUpdate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@AttDate", _AttDate)
            BaseConn.cmd.Parameters.AddWithValue("@Category", _Category)
            BaseConn.cmd.Parameters.AddWithValue("@AttMainDT", _DTAttMain)
            BaseConn.cmd.Parameters.AddWithValue("@AttSubDT", _DTAttSub)
            BaseConn.cmd.Parameters.AddWithValue("@UserID", _UserID)

            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output

            BaseConn.cmd.ExecuteNonQuery()

            iRC = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            iRC = 1
            _ErrString = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

End Class
