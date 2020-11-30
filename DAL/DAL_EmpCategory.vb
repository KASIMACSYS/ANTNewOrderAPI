Imports Classes
Public Class DAL_EmpCategory
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Function Update_EmpCategory(ByVal Obj As csEmpCategory, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[EmployeeCategory]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.str_flag)
            BaseConn.cmd.Parameters.AddWithValue("@EmpCategoryDT", Obj.dt_empcategory)
            BaseConn.cmd.Parameters.AddWithValue("@EmpCategoryDetailsDT", Obj.dt_empcategorysub)
            BaseConn.cmd.Parameters.AddWithValue("@GratuityCalculationDT", Obj.dt_Gratuity)
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(Obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(Obj.str_SiteID, _StrDBPath, _StrDBPwd, Obj.int_BusinessPeriodID, "", DateTime.Now, "", "EmpCategory", Err.Number, "Error in : " & Obj.str_LeaveTypes & "  ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        Update_EmpCategory = _ErrString
    End Function

    Public Sub Get_Structure(ByRef Obj As csEmpCategory, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[EmployeeCategory]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.str_flag)
            BaseConn.cmd.Parameters.AddWithValue("@ErrorNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            Obj.dt_empcategory = ds.Tables(0)
            Obj.dt_empcategorysub = ds.Tables(1)
            Obj.dt_Gratuity = ds.Tables(2)
            Obj.dt_LeaveType = ds.Tables(3)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
   
End Class
