Imports Classes

Public Class DAL_EmpPaySlip
    Private BaseConn As New SQLConn()
    Private dt As DataTable
    Private ObjDalGeneral As DAL_General

    Public Sub GetPaySlipBasicDetails(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _SiteID As Integer, ByVal _BusinessPeriod As Integer,
                                  ByRef dt_SalParticulars As DataTable, ByRef dt_SalPartWithOriginalValues As DataTable, ByRef _ErrNo As Integer, ByVal payslipmonth As Date)
        _ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetPaySlipBasicDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@PayslipMonth", payslipmonth)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriod)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt_SalParticulars = ds.Tables(0)
            dt_SalPartWithOriginalValues = ds.Tables(1)
        Catch ex As Exception
            _ErrNo = 1
        End Try
    End Sub

    'Public Sub GetPaySlipDetails(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _SiteID As Integer, ByVal _BusinessPeriod As Integer,
    '                             ByVal _PSRef As String, ByRef _DTPaySlipMain As DataTable, ByRef _ErrNo As Integer)
    '    _ErrNo = 0
    '    Try
    '        BaseConn.Open(_DBPath, _DBPwd)
    '        BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetPaySlipDetails]", BaseConn.cnn)
    '        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    '        BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
    '        BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriod)
    '        BaseConn.cmd.Parameters.AddWithValue("@PSRef", _PSRef)
    '        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    '        Dim ds As New DataSet
    '        BaseConn.da.Fill(ds)
    '        _DTPaySlipMain = ds.Tables(0)
    '    Catch ex As Exception
    '        _ErrNo = 1
    '    End Try
    'End Sub

    Public Sub GetPaySlipDetails(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _SiteID As Integer, ByVal _BusinessPeriod As Integer,
                                 ByVal _PSRef As String, ByRef _DTPaySlipMain As DataTable, ByRef _DTPayslipSub As DataTable,
                                ByRef _DTEmployees As DataTable, ByRef _DTSalPartWithGrossAmount As DataTable, ByRef _DTAccountLedger As DataTable,
                                ByRef _DTPSParameter As DataTable, ByVal _Category As String, ByVal _LedgerID As String, ByRef _TempID As String, ByRef _ErrNo As Integer)
        _ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetPaySlipDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriod)
            BaseConn.cmd.Parameters.AddWithValue("@PSRef", _PSRef)
            BaseConn.cmd.Parameters.AddWithValue("@Category", _Category)
            BaseConn.cmd.Parameters.AddWithValue("@Ledger", _LedgerID)
            BaseConn.cmd.Parameters.Add("@TemplateID", SqlDbType.NVarChar, 50).Direction = ParameterDirection.Output
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _DTPaySlipMain = ds.Tables(0)
            _DTPayslipSub = ds.Tables(1)
            _DTEmployees = ds.Tables(2)
            _DTSalPartWithGrossAmount = ds.Tables(3)
            _DTAccountLedger = ds.Tables(4)
            _DTPSParameter = ds.Tables(5)
            _TempID = BaseConn.cmd.Parameters("@TemplateID").Value.ToString
        Catch ex As Exception
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub Record4HRPayroll(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _SiteID As Integer, ByVal _BusinessPeriod As Integer,
                                  ByVal _Flag As String, ByVal _PSRef As String, ByVal _GivenDate As Date, ByVal _Category As String, ByVal _Ledger As String,
                                  ByRef dt_PaySlip As DataTable, ByRef dt_LedgerAccounts As DataTable, ByRef _PaidPSCnt As Integer, ByVal _MenuID As String, ByVal _Type As String, ByVal _dtAttendance As DataTable, ByRef dt_CostCentre As DataTable, ByRef _ErrNo As Integer)
        _ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetHRPayroll]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriod)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@PSRef", _PSRef)
            BaseConn.cmd.Parameters.AddWithValue("@GivenDate", _GivenDate)
            BaseConn.cmd.Parameters.AddWithValue("@Category", _Category)
            BaseConn.cmd.Parameters.AddWithValue("@Ledger", _Ledger)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", _MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Type", _Type)
            BaseConn.cmd.Parameters.AddWithValue("@AttenDT", _dtAttendance)
            BaseConn.cmd.Parameters.Add("@PaidPSCnt", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt_PaySlip = ds.Tables(0)
            dt_LedgerAccounts = ds.Tables(1)
            If ds.Tables(2) IsNot Nothing Then
                dt_CostCentre = ds.Tables(2)
            End If
            _PaidPSCnt = BaseConn.cmd.Parameters("@PaidPSCnt").Value.ToString
        Catch ex As Exception
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub UpdateMultiPayroll(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _ObjcsPayRoll As csEmpPayslip, ByRef _PSRef As String,
                                  ByRef intRevNo As Integer, ByVal _TemplateID As String, ByRef _OutSMS As String, ByRef _OutEmail As String, ByRef _ErrNo As Integer, ByRef _ErrString As String)

        _ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[PayrollUpdatedMulti]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _ObjcsPayRoll.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _ObjcsPayRoll.objPSMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", _ObjcsPayRoll.objPSMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@PSRef", _ObjcsPayRoll.objPSMain.str_PSRef)
            BaseConn.cmd.Parameters.AddWithValue("@TemplateID", _TemplateID)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", _ObjcsPayRoll.objPSMain.str_FormPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _ObjcsPayRoll.objPSMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@PostDate", _ObjcsPayRoll.objPSMain.date_PostDate)
            BaseConn.cmd.Parameters.AddWithValue("@PSMonth", _ObjcsPayRoll.objPSMain.date_PSMonth)
            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", _ObjcsPayRoll.str_CurrencyCode)

            BaseConn.cmd.Parameters.AddWithValue("@PaySlipMainDT", _ObjcsPayRoll.objPSMain.DT_PaySlipMain)
            BaseConn.cmd.Parameters.AddWithValue("@PaySlipSubDT", _ObjcsPayRoll.objPSMain.DT_PaySlipSub)
            BaseConn.cmd.Parameters.AddWithValue("@PaySlipSubParam", _ObjcsPayRoll.objPSMain.DT_PayslipParam)
            BaseConn.cmd.Parameters.AddWithValue("@InvAccDetDT", _ObjcsPayRoll.objPSMain.dt_InvoiceAccounts)
            BaseConn.cmd.Parameters.AddWithValue("@AdvanceDT", _ObjcsPayRoll.objPSMain.dt_Wages)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", _ObjcsPayRoll.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", _ObjcsPayRoll.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", _ObjcsPayRoll.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", _ObjcsPayRoll.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", _ObjcsPayRoll.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", _ObjcsPayRoll.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", _ObjcsPayRoll.bool_ApprovedStatus)

            BaseConn.cmd.Parameters.Add("@PSRefOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@RevNoOut", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutSMSMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutEmailMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output

            BaseConn.cmd.ExecuteNonQuery()
            _PSRef = BaseConn.cmd.Parameters("@PSRefOut").Value.ToString
            _ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@RevNoOut").Value.ToString
            _OutSMS = BaseConn.cmd.Parameters("@OutSMSMsgText").Value.ToString
            _OutEmail = BaseConn.cmd.Parameters("@OutEmailMsgText").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(_ObjcsPayRoll.str_SiteID)
            ObjDalGeneral.Elog_Insert(_ObjcsPayRoll.str_SiteID, _DBPath, _DBPwd, 0, _ObjcsPayRoll.str_CreatedBy, _ObjcsPayRoll.dtp_CreatedDate, "", "PSV", Err.Number, "Error in " & _ObjcsPayRoll.objPSMain.str_Flag & " : " & "" & " ", ex.Message, 5, 3, 1, _ErrNo)
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub PayrollMonthDaysDetails(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _SiteID As Integer,
                                  ByVal _GivenDate As Date, ByRef _NoofHoliday As Double, ByRef _dt_WeekEnd As DataTable, ByRef _DaysInMonth As Double, ByRef _ErrNo As Integer)
        _ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[PayrollMonthDaysDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@GivenDate", _GivenDate)
            BaseConn.cmd.Parameters.Add("@NoofHoliday", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@MonthInDays", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _NoofHoliday = BaseConn.cmd.Parameters("@NoofHoliday").Value
            _DaysInMonth = BaseConn.cmd.Parameters("@MonthInDays").Value
            _dt_WeekEnd = ds.Tables(0)
        Catch ex As Exception
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

    End Sub
    Public Sub Get_Structure(ByRef Obj As csEmpPaySlipMain, ByVal _SiteID As String, ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByVal _Flag As String, ByVal ErrNo As Integer, ByVal ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_strDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetEmpPaySlipMain]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", Obj.str_EmpName)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", Obj.str_VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@Category", Obj.str_Categotry)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", Obj.dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", Obj.dtp_ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Date1", Obj.dtp_Date)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            dt = New DataTable
            BaseConn.da.Fill(dt)
            Obj.dt_Main = dt
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ""
        Finally
            BaseConn.Close()
        End Try

    End Sub
    Public Sub Get_Incentive(ByRef Obj As csEmpPaySlipMain, ByVal _SiteID As String, ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByVal _Flag As String, ByVal _Condition As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_strDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetIncentiveMain]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", Obj.str_EmpName)
            BaseConn.cmd.Parameters.AddWithValue("@Category", Obj.str_Categotry)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", Obj.dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", Obj.dtp_ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Date1", Obj.dtp_Date)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", _Condition)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            dt = New DataTable
            BaseConn.da.Fill(dt)
            Obj.dt_Main = dt
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try

    End Sub
    ''' <summary>
    ''' Used to retrive data(Datatable) based on the mode(Payslip/Wages) which we pass
    ''' </summary>
    ''' <param name="_strDBPath"></param>
    ''' <param name="_StrDBPwd"></param>
    ''' <param name="_StrSiteID"></param>
    ''' <param name="_BusinessPeriodID"></param>
    ''' <param name="_LedgerID"></param>
    ''' <param name="_Mode"></param>
    ''' <param name="_VouNo"></param>
    ''' <param name="ErrNo"></param>
    ''' <param name="ErrStr"></param>
    ''' <remarks></remarks>
    Public Sub GetHRDetialsForVou(ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByVal _StrSiteID As String, ByVal _BusinessPeriodID As Integer,
                                ByRef DT4VouPS As DataTable, ByRef DT4VouWages As DataTable, ByVal _LedgerID As Integer, ByVal _Mode As String, ByVal _VouNo As String,
                                  ByVal _VouDate As Date, ByVal _FormType As String, _AdvType As String, ByVal ErrNo As Integer, ByVal ErrStr As String, ByRef _AdvOut As Double)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_strDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetEmpAdvanceMatching]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _StrSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Mode", _Mode)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", _VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@VouDate", _VouDate)
            'BaseConn.cmd.Parameters.AddWithValue("@Condition", _Condition)
            BaseConn.cmd.Parameters.AddWithValue("@FormType", _FormType)
            BaseConn.cmd.Parameters.AddWithValue("@AdvType", _AdvType)
            BaseConn.cmd.Parameters.Add("@AdvOut", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _AdvOut = BaseConn.cmd.Parameters("@AdvOut").Value
            DT4VouPS = ds.Tables(0)
            DT4VouWages = ds.Tables(1)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ""
        Finally
            BaseConn.Close()
        End Try
    End Sub

    'Public Sub GetHRDetialsForVou(ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByVal _StrSiteID As String, ByVal _BusinessPeriodID As Integer, _
    '                             ByRef DT4VouPS As DataTable, ByRef DT4VouWages As DataTable, ByVal _LedgerID As Integer, ByVal _Mode As String, ByVal _VouNo As String, _
    '                             ByVal _VouDate As Date, ByVal ErrNo As Integer, ByVal ErrStr As String, ByVal _Condition As String, Optional ByRef _AdvOut As Double = 0)
    '    ErrNo = 0
    '    ErrStr = ""
    '    Try
    '        BaseConn.Open(_strDBPath, _StrDBPwd)
    '        BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetHRDetailsForVou]", BaseConn.cnn)
    '        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    '        BaseConn.cmd.Parameters.AddWithValue("@SiteID", _StrSiteID)
    '        BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
    '        BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
    '        BaseConn.cmd.Parameters.AddWithValue("@Mode", _Mode)
    '        BaseConn.cmd.Parameters.AddWithValue("@VouNo", _VouNo)
    '        BaseConn.cmd.Parameters.AddWithValue("@VouDate", _VouDate)
    '        BaseConn.cmd.Parameters.AddWithValue("@Condition", _Condition)
    '        BaseConn.cmd.Parameters.Add("@AdvOut", SqlDbType.Decimal).Direction = ParameterDirection.Output
    '        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    '        Dim ds As New DataSet
    '        BaseConn.da.Fill(ds)
    '        _AdvOut = BaseConn.cmd.Parameters("@AdvOut").Value
    '        DT4VouPS = ds.Tables(0)
    '        DT4VouWages = ds.Tables(1)
    '    Catch ex As Exception
    '        ErrNo = 1
    '        ErrStr = ""
    '    Finally
    '        BaseConn.Close()
    '    End Try
    'End Sub

    Public Sub GetPensionReport(ByVal _SiteID As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _BusinessPeriodID As Integer, ByVal _Flag As String, ByVal _LedgerID As Integer, ByVal _Fromdate As Date, ByVal _ToDate As Date, ByRef _dtPension As DataTable, ByRef _HeaderText As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetPensionReport]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _Fromdate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.Add("@HeaderText", SqlDbType.NVarChar, 50).Direction = ParameterDirection.Output
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _HeaderText = BaseConn.cmd.Parameters("@HeaderText").Value
            _dtPension = ds.Tables(0)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Sub GetJVsReport(ByVal _SiteID As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _BusinessPeriodID As Integer, ByVal _Flag As String, ByVal _LedgerID As Integer, ByVal _Fromdate As Date, ByVal _ToDate As Date, ByRef _dtJVs As DataTable, ByRef _dtHead As DataTable, ByRef _HeaderText As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetJVsReport]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _Fromdate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.Add("@HeaderText", SqlDbType.NVarChar, 50).Direction = ParameterDirection.Output
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _HeaderText = BaseConn.cmd.Parameters("@HeaderText").Value
            _dtJVs = ds.Tables(0)
            _dtHead = ds.Tables(1)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetDeductionDetails(ByVal _SiteID As String, ByVal _StrDBPath As String, ByVal _StrDBPWD As String, ByVal _PSMonth As Date, ByRef _Flag As String, ByRef _dt As DataTable, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        _ErrNo = 0
        _ErrStr = ""
        _dt = New DataTable
        Try
            BaseConn.Open(_StrDBPath, _StrDBPWD)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_EmployeeDeductionDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@PSMonth", _PSMonth)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _dt = ds.Tables(0)
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function GetEmpAtt(ByVal _SiteID As String, ByVal _StrDBPath As String, ByVal _StrDBPWD As String, ByVal _PSMonth As Date, ByRef _MonthDays As Integer, ByVal _CompanyName As String, ByRef _dt As DataTable, ByRef _ErrNo As Integer, ByRef _ErrStr As String, ByVal _MenuID As String) As DataTable
        _ErrNo = 0
        _ErrStr = ""
        _dt = New DataTable
        Try
            BaseConn.Open(_StrDBPath, _StrDBPWD)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_ERP_EmpAtt]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", "")
            BaseConn.cmd.Parameters.AddWithValue("@PSMonth", _PSMonth)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", _MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@CompanyName", _CompanyName)
            BaseConn.cmd.Parameters.Add("@WorkDays", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _dt = ds.Tables(0)
            _MonthDays = BaseConn.cmd.Parameters("@WorkDays").Value.ToString
        Catch ex As Exception
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Return _dt
    End Function
    Public Function GetLeaveDetails(ByVal _SiteID As String, ByVal _StrDBPath As String, ByVal _StrDBPWD As String, ByVal _PSMonth As Date, ByRef _dt As DataTable, ByRef _ErrNo As Integer, ByRef _ErrStr As String) As DataTable
        _ErrNo = 0
        _ErrStr = ""
        _dt = New DataTable
        Try
            BaseConn.Open(_StrDBPath, _StrDBPWD)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_ERP_GetLeaveDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@PSMonth", _PSMonth)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _dt = ds.Tables(0)
        Catch ex As Exception
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Return _dt
    End Function

    Public Sub GetPaySlipTemplate(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _SiteID As Integer,
                                ByVal _TempID As String, ByRef dt_PayslipTemplate As DataTable, ByRef _ErrNo As Integer)
        _ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetPaySlipTemplate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@TemplateID", _TempID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt_PayslipTemplate = ds.Tables(0)
            'dt_PayslipTemplateParam = ds.Tables(1)
        Catch ex As Exception
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetPaySlipGrossSalaryDetails(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _SiteID As Integer, ByVal payslipmonth As Date,
                                 ByVal _Category As String, ByVal _Ledger As String, ByVal _TemplateID As String, ByRef dt_Employees As DataTable,
                                ByRef dt_SalPartWithGrossAmount As DataTable, ByRef dt_Absent As DataTable, ByRef dt_OTHours As DataTable, ByRef _ErrNo As Integer)
        _ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetPaySlipGrossSalaryDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@PayslipMonth", payslipmonth)
            BaseConn.cmd.Parameters.AddWithValue("@Category", _Category)
            BaseConn.cmd.Parameters.AddWithValue("@Ledger", _Ledger)
            BaseConn.cmd.Parameters.AddWithValue("@TemplateID", _TemplateID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt_Employees = ds.Tables(0)
            dt_SalPartWithGrossAmount = ds.Tables(1)
            dt_Absent = ds.Tables(2)
            dt_OTHours = ds.Tables(3)
        Catch ex As Exception
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub

End Class
