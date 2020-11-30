Imports Classes

Public Class DAL_EmpAtt
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General


    Public Sub SyncFingerPrintDB(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _BSPID As String, _
                                      ByVal _GivenDate As Date, ByVal _DeviceID As String, ByVal _CreatedBy As String, _
                                     ByRef iRC As Integer, ByRef _ErrString As String)
        iRC = 0
        'Dim _ErrString As String = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_SyncFingerPrintDB]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BSPID)
            BaseConn.cmd.Parameters.AddWithValue("@GivenDate", _GivenDate)
            BaseConn.cmd.Parameters.AddWithValue("@DeviceID", _DeviceID)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", _CreatedBy)

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

    Public Sub GetEmployeeAttendance(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _LedgerID As String, _
                                     ByVal _Category As String, ByVal _FromDate As Date, ByVal _ToDate As Date, ByRef _DTOT As DataTable, ByRef _DTPresent As DataTable, _
                                     ByRef iRC As Integer, ByRef ErrStr As String)
        iRC = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetEmployeeAttendance]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Category", _Category)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _DTOT = ds.Tables(0)
            _DTPresent = ds.Tables(1)
        Catch ex As Exception
            iRC = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

   

    Public Sub GetEmployeeForAttendanceInOut(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _LedgerID As String, _
                                        ByVal _FromDate As Date, ByRef _DTDBINOUT As DataTable, ByRef iRC As Integer, ByRef ErrStr As String)
        iRC = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetEmployeeForAttendanceInOut]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@GivenDate", _FromDate)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            '_DTEmployee = ds.Tables(0)
            _DTDBINOUT = ds.Tables(0)
        Catch ex As Exception
            iRC = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub


    Public Sub GetEmployeeForAttendance(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _Flag As String, ByVal _LedgerID As String, ByVal _Category As String, _
                                        ByVal _FromDate As Date, ByRef _DTEmployee As DataTable, ByRef _DTDBINOUT As DataTable, ByRef iRC As Integer, ByRef ErrStr As String)
        iRC = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetEmployeeForAttendance]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
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

    Public Sub EmployeeAttendanceUpdate(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _Flag As String, _
                                      ByVal _AttDate As Date, ByVal _Category As String, ByVal _DTAttMain As DataTable, ByVal _DTAttSub As DataTable, _
                                      ByVal _CreatedBy As String, ByRef iRC As Integer, ByRef _ErrString As String)
        iRC = 0
        'Dim _ErrString As String = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_EmployeeAttendanceUpdate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@AttDate", _AttDate)
            BaseConn.cmd.Parameters.AddWithValue("@Category", _Category)
            BaseConn.cmd.Parameters.AddWithValue("@AttMainDT", _DTAttMain)
            BaseConn.cmd.Parameters.AddWithValue("@AttSubDT", _DTAttSub)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", _CreatedBy)

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

    Public Sub GetEmployeeDeatailsAttendance(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _DTEmpID As DataTable, _
                                             ByRef _DTEmpIDWithLedger As DataTable, ByRef iRC As Integer, ByRef ErrStr As String)
        iRC = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetEmployeeDeatailsAttendance]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@DTEmpID", _DTEmpID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _DTEmpIDWithLedger = ds.Tables(0)
        Catch ex As Exception
            iRC = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub AttExcelImport(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _AttDT As DataTable, ByRef _ErrNo As Integer, ByRef ErrStr As String)
        _ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_AttExcelImport]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@AttDT", _AttDT)
            BaseConn.cmd.ExecuteNonQuery()
        Catch ex As Exception
            _ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
End Class
