'======================================================================================
'$Author: Kasim $
'$Rev:  $
'$Date: 2014-07-03  $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes
Public Class DAL_LeaveSalary
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Function CalcLeaveSalary(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef _ObjLS As csLeaveSalary, _
                                    ByRef _DaysInMonth As Integer, ByRef _DaysInYear As Integer, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0

        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("CalcLeaveSalary", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _ObjLS._SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _ObjLS._LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ObjLS._PostDate)
            BaseConn.cmd.Parameters.AddWithValue("@ExcludeEmergency", _ObjLS._ExcludeEmergency)

            'BaseConn.cmd.Parameters.Add("@PaySlipAmt", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@RejoinDate", SqlDbType.Date).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@DaysInMonth", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@DaysInYear", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@NoofDays", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@LeaveDays", SqlDbType.Int).Direction = ParameterDirection.Output

            BaseConn.cmd.Parameters.Add("@PerDay", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@TotalLeaveDays", SqlDbType.Float).Direction = ParameterDirection.Output

            BaseConn.cmd.Parameters.Add("@MonthlyLeaveAmount", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@PerDayLeaveAmount", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@LeaveSalary", SqlDbType.Float).Direction = ParameterDirection.Output

            BaseConn.cmd.Parameters.Add("@PassageAmt", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@PerDayPassageAmt", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@NetPassageAmt", SqlDbType.Float).Direction = ParameterDirection.Output

            BaseConn.cmd.Parameters.Add("@BasicSalaryforGratuity", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@JoiningDate", SqlDbType.Date).Direction = ParameterDirection.Output

            'BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            'BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()

            '_ObjLS._PSAmt = BaseConn.cmd.Parameters("@PaySlipAmt").Value.ToString
            _ObjLS._FromDate = BaseConn.cmd.Parameters("@RejoinDate").Value.ToString
            _ObjLS._NoofDays = BaseConn.cmd.Parameters("@NoofDays").Value.ToString
            _ObjLS._LeaveDays = BaseConn.cmd.Parameters("@LeaveDays").Value.ToString
            _ObjLS._PerDay = BaseConn.cmd.Parameters("@PerDay").Value.ToString
            _ObjLS._TotalLeaveDays = BaseConn.cmd.Parameters("@TotalLeaveDays").Value.ToString
            _ObjLS._MonthlyLeaveAmount = BaseConn.cmd.Parameters("@MonthlyLeaveAmount").Value.ToString
            _ObjLS._PerDayLeaveAmount = BaseConn.cmd.Parameters("@PerDayLeaveAmount").Value.ToString
            _ObjLS._CalcLeaveAmt = BaseConn.cmd.Parameters("@LeaveSalary").Value.ToString
            _ObjLS._ChagLeaveAmt = _ObjLS._CalcLeaveAmt
            _ObjLS._PassageAmt = BaseConn.cmd.Parameters("@PassageAmt").Value.ToString
            _ObjLS._PerDayPassageAmt = BaseConn.cmd.Parameters("@PerDayPassageAmt").Value.ToString
            _ObjLS._CalcPassageAmt = BaseConn.cmd.Parameters("@NetPassageAmt").Value.ToString
            _ObjLS._ChagPassageAmt = _ObjLS._CalcPassageAmt

            _ObjLS._BasicSalaryforGratuity = BaseConn.cmd.Parameters("@BasicSalaryforGratuity").Value.ToString
            _ObjLS._JoiningDate = BaseConn.cmd.Parameters("@JoiningDate").Value.ToString
            _DaysInMonth = BaseConn.cmd.Parameters("@DaysInMonth").Value.ToString
            _DaysInYear = BaseConn.cmd.Parameters("@DaysInYear").Value.ToString
            'ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            '_ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString

        Catch ex As Exception
            _ErrString = ex.Message
            'ObjDalGeneral = New DAL_General(obj.str_SiteID)
            'ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.objDOMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "DO", Err.Number, "Error in " & obj.objDOMain.str_Flag & " : " & obj.objDOMain.str_DoNo & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return _ErrString
    End Function

    Public Function GetLeaveSalary(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef _ObjLS As csLeaveSalary, _
                                     ByRef _DTInvAccDetails As DataTable, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Dim _DTPayslip As DataTable, _DTLeaveSalary As DataTable

        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("GetLeaveSalaryDetails", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _ObjLS._SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@PSRef", _ObjLS._RefNo)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _DTPayslip = ds.Tables(0)
            _DTLeaveSalary = ds.Tables(1)
            _DTInvAccDetails = ds.Tables(2)

            _ObjLS._LedgerID = _DTPayslip.Rows(0)("LedgerID")
            _ObjLS._PostDate = _DTPayslip.Rows(0)("PostDate")
            _ObjLS._PSAmt = 0 '_DTPayslip.Rows(0)("NetAmount")
            _ObjLS._Comment = _DTPayslip.Rows(0)("Comment")
            _ObjLS._RevNo = _DTPayslip.Rows(0)("RevNo")

            _ObjLS._FromDate = _DTLeaveSalary.Rows(0)("FromDate")
            _ObjLS._ToDate = _DTLeaveSalary.Rows(0)("ToDate")
            _ObjLS._NoofDays = _DTLeaveSalary.Rows(0)("NoofDays")
            _ObjLS._LeaveDays = _DTLeaveSalary.Rows(0)("LeaveDays")
            _ObjLS._PerDay = _DTLeaveSalary.Rows(0)("PerDay")
            _ObjLS._TotalLeaveDays = _DTLeaveSalary.Rows(0)("TotalDays")
            _ObjLS._MonthlyLeaveAmount = _DTLeaveSalary.Rows(0)("MonthlyLeaveAmt")
            _ObjLS._PerDayLeaveAmount = _DTLeaveSalary.Rows(0)("PerDayLeaveAmt")
            _ObjLS._CalcLeaveAmt = _DTLeaveSalary.Rows(0)("CalcLeaveAmt")
            _ObjLS._ChagLeaveAmt = _DTLeaveSalary.Rows(0)("ChagLeaveAmt")
            _ObjLS._PassageAmt = _DTLeaveSalary.Rows(0)("PassageAmt")
            _ObjLS._PerDayPassageAmt = _DTLeaveSalary.Rows(0)("PerDayPassageAmt")
            '_ObjLS._CalcLeaveAmt = _DTLeaveSalary.Rows(0)("CalcLeaveAmt")
            _ObjLS._ChagPassageAmt = _DTLeaveSalary.Rows(0)("ChagPassageAmt")
            _ObjLS._CalcPassageAmt = _DTLeaveSalary.Rows(0)("CalcPassageAmt")
            _ObjLS._PensionAmt = _DTLeaveSalary.Rows(0)("PensionAmt")
            _ObjLS._TotalAmt = _DTLeaveSalary.Rows(0)("TotalAmt")
            _ObjLS._LessAmt = _DTLeaveSalary.Rows(0)("LessAmt")
            _ObjLS._NetAmt = _DTLeaveSalary.Rows(0)("NetAmt")
            _ObjLS._ExcludeEmergency = _DTLeaveSalary.Rows(0)("ExcludeEmergency")

            _ObjLS._CalcLastMonthSalary = _DTLeaveSalary.Rows(0)("CalcLastMonthSalary")
            _ObjLS._ChangeLastMonthSalary = _DTLeaveSalary.Rows(0)("ChangeLastMonthSalary")
            _ObjLS._CalcGratuity = _DTLeaveSalary.Rows(0)("CalcGratuity")
            _ObjLS._ChangeGratuity = _DTLeaveSalary.Rows(0)("ChangeGratuity")
            _ObjLS._PSMonth = _DTPayslip.Rows(0)("PSMonth")

            _ObjLS._BasicSalary = _DTLeaveSalary.Rows(0)("BasicSalary")
            _ObjLS._PerDayGratuity = _DTLeaveSalary.Rows(0)("PerDayGratuity")
            _ObjLS._TotalDaysGratuity = _DTLeaveSalary.Rows(0)("TotalDaysGratuity")

        Catch ex As Exception
            _ErrString = ex.Message
            'ObjDalGeneral = New DAL_General(obj.str_SiteID)
            'ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.objDOMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "DO", Err.Number, "Error in " & obj.objDOMain.str_Flag & " : " & obj.objDOMain.str_DoNo & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return _ErrString
    End Function

    Public Function Update_LeaveSalary _
                        (ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _ObjLS As csLeaveSalary, ByRef _VouNo As String, _
                         ByVal _Alias As String, ByVal _DTAccountLedger As DataTable, ByRef _RevNo As Integer, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0

        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("UpdateLeaveSalary", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _ObjLS._SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _ObjLS._Flag)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", _ObjLS._MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", _ObjLS._FormPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _ObjLS._BusinessPeriodID)

            BaseConn.cmd.Parameters.AddWithValue("@VouNo", _ObjLS._RefNo)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _ObjLS._LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Alias", _Alias)
            BaseConn.cmd.Parameters.AddWithValue("@PostDate", _ObjLS._PostDate)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _ObjLS._FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ObjLS._PostDate)

            BaseConn.cmd.Parameters.AddWithValue("@NoofDays", _ObjLS._NoofDays)
            BaseConn.cmd.Parameters.AddWithValue("@LeaveDays", _ObjLS._LeaveDays)
            BaseConn.cmd.Parameters.AddWithValue("@WorkingDays", _ObjLS._NoofDays - _ObjLS._LeaveDays)
            BaseConn.cmd.Parameters.AddWithValue("@PerDay", _ObjLS._PerDay)
            BaseConn.cmd.Parameters.AddWithValue("@TotalDays", _ObjLS._TotalLeaveDays)

            BaseConn.cmd.Parameters.AddWithValue("@PSAmount", _ObjLS._PSAmt)

            BaseConn.cmd.Parameters.AddWithValue("@MonthlyLeaveAmt", _ObjLS._MonthlyLeaveAmount)
            BaseConn.cmd.Parameters.AddWithValue("@PerDayLeaveAmt", _ObjLS._PerDayLeaveAmount)
            BaseConn.cmd.Parameters.AddWithValue("@CalcLeaveAmt", _ObjLS._CalcLeaveAmt)
            BaseConn.cmd.Parameters.AddWithValue("@ChagLeaveAmt", _ObjLS._ChagLeaveAmt)
            BaseConn.cmd.Parameters.AddWithValue("@PassageAmt", _ObjLS._PassageAmt)
            BaseConn.cmd.Parameters.AddWithValue("@PerDayPassageAmt", _ObjLS._PerDayPassageAmt)
            BaseConn.cmd.Parameters.AddWithValue("@CalcPassageAmt", _ObjLS._CalcPassageAmt)
            BaseConn.cmd.Parameters.AddWithValue("@ChagPassageAmt", _ObjLS._ChagPassageAmt)
            BaseConn.cmd.Parameters.AddWithValue("@PensionAmt", _ObjLS._PensionAmt)

            BaseConn.cmd.Parameters.AddWithValue("@CalcLastMonthSalary", _ObjLS._CalcLastMonthSalary)
            BaseConn.cmd.Parameters.AddWithValue("@ChangeLastMonthSalary", _ObjLS._ChangeLastMonthSalary)
            BaseConn.cmd.Parameters.AddWithValue("@CalcGratuity", _ObjLS._CalcGratuity)
            BaseConn.cmd.Parameters.AddWithValue("@ChangeGratuity", _ObjLS._ChangeGratuity)

            BaseConn.cmd.Parameters.AddWithValue("@BasicSalary", _ObjLS._BasicSalary)
            BaseConn.cmd.Parameters.AddWithValue("@PerDayGratuity", _ObjLS._PerDayGratuity)
            BaseConn.cmd.Parameters.AddWithValue("@TotalDaysGratuity", _ObjLS._TotalDaysGratuity)

            BaseConn.cmd.Parameters.AddWithValue("@TotalAmt", _ObjLS._TotalAmt)
            BaseConn.cmd.Parameters.AddWithValue("@LessAmt", _ObjLS._LessAmt)
            BaseConn.cmd.Parameters.AddWithValue("@NetAmt", _ObjLS._NetAmt)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", _ObjLS._Comment)
            BaseConn.cmd.Parameters.AddWithValue("@ExcludeEmergency", _ObjLS._ExcludeEmergency)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", _ObjLS.CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", _ObjLS.CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", _ObjLS.LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", _ObjLS.LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", _ObjLS.ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", _ObjLS.ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", _ObjLS.ApprovedStatus)

            BaseConn.cmd.Parameters.AddWithValue("@InvAccDetDT", _DTAccountLedger)

            BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@RevNoOut", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            _VouNo = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
            _RevNo = BaseConn.cmd.Parameters("@RevNoOut").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString

        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(_ObjLS._SiteID)
            ObjDalGeneral.Elog_Insert(_ObjLS._SiteID, _StrDBPath, _StrDBPwd, _ObjLS._BusinessPeriodID, _ObjLS.CreatedBy, _ObjLS.CreatedDate, "", "LeaveSalary", Err.Number, "Error in " & _ObjLS._Flag & " : " & _VouNo & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_LeaveSalary = _ErrString
    End Function
End Class
