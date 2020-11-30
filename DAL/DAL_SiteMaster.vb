'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Imports Classes

Public Class DAL_SiteMaster
    Dim dt As DataTable
    Dim ds As DataSet
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef Obj As csSiteMaster, ByVal _strDBPath As String, ByVal _strDBPwd As String, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        _ErrNo = 0
        _ErrStr = ""
        Try
            BaseConn.Open(_strDBPath, _strDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetSiteMasterDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            ds = New DataSet
            BaseConn.da.Fill(ds)

            Obj.str_SiteMasterID = ds.Tables(0).Rows(0)("CID").ToString 'changes
            Obj.str_SiteName = ds.Tables(0).Rows(0)("CompanyName").ToString
            Obj.str_Alias = ds.Tables(0).Rows(0)("CompanyName").ToString 'ds.Tables(0).Rows(0)("Alias").ToString
            Obj.str_Address = ds.Tables(0).Rows(0)("Address").ToString
            Obj.str_Contact = ds.Tables(0).Rows(0)("Contact").ToString
            Obj.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString
            Obj.str_DBBackUpPath = ds.Tables(0).Rows(0)("DBBackUpPath").ToString
            Obj.int_DaysAnnualLeave = ds.Tables(0).Rows(0)("DaysAnnualLeave").ToString
            Obj.int_DaysSickLeave = ds.Tables(0).Rows(0)("DaysSickLeave").ToString
            Obj.int_DaysCarryForward = ds.Tables(0).Rows(0)("DaysCarryFrwd").ToString
            Obj.str_BankLedgerID = ds.Tables(0).Rows(0)("BankLedgerID").ToString
            Obj.int_CashLedgerID = ds.Tables(0).Rows(0)("CashLedgerID").ToString
            Obj.str_TRN = ds.Tables(0).Rows(0)("TRN").ToString
            Obj.str_TaxablePersonNameEng = ds.Tables(0).Rows(0)("TaxablePersonNameEn").ToString
            Obj.str_TaxablePersonNameArab = ds.Tables(0).Rows(0)("TaxablePersonNameAr").ToString
            Obj.str_ProductVersion = ds.Tables(0).Rows(0)("ProductVersion").ToString
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Update_SiteMaster(ByVal obj As csSiteMaster, ByRef SiteID As String, ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_strDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[SiteMasterUpdate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@SiteMasterID", obj.str_SiteMasterID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@SiteName", obj.str_SiteName)
            BaseConn.cmd.Parameters.AddWithValue("@Alias", obj.str_Alias)
            BaseConn.cmd.Parameters.AddWithValue("@Contact", obj.str_Contact)
            BaseConn.cmd.Parameters.AddWithValue("@Address", obj.str_Address)
            BaseConn.cmd.Parameters.AddWithValue("@DBBackUpPath", obj.str_DBBackUpPath)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.str_Comment)
            'BaseConn.cmd.Parameters.AddWithValue("@DaysSickLeave", obj.int_DaysSickLeave)
            'BaseConn.cmd.Parameters.AddWithValue("@DaysAnnualLeave", obj.int_DaysAnnualLeave)
            'BaseConn.cmd.Parameters.AddWithValue("@DaysCarryFrwd", obj.int_DaysCarryForward)
            BaseConn.cmd.Parameters.AddWithValue("@BankLedgerID", obj.str_BankLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@CashLedgerID", obj.int_CashLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@TRN", obj.str_TRN)
            BaseConn.cmd.Parameters.AddWithValue("@TaxablePersonNameEng", obj.str_TaxablePersonNameEng)
            BaseConn.cmd.Parameters.AddWithValue("@TaxablePersonNameArab", obj.str_TaxablePersonNameArab)
            BaseConn.cmd.Parameters.AddWithValue("@ProductVersion", obj.str_ProductVersion)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdateBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdateDate)
			BaseConn.cmd.Parameters.Add("@SiteIDOut", SqlDbType.NVarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
 			SiteID = BaseConn.cmd.Parameters("@SiteIDOut").Value.ToString           
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _strDBPath, _StrDBPwd, 0, obj.str_CreatedBy, obj.dtp_LastUpdateDate, obj.str_CreatedBy, "SiteMaster", Err.Number, "Error in " & obj.str_Flag & " : " & obj.str_SiteName & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        Update_SiteMaster = _ErrString
    End Function

    'Public Function GetAllSites(ByRef ErrNo As Integer) As DataTable
    '    GetAllSites = Nothing
    '    ErrNo = 0
    '    Try
    '        BaseConn.Open()
    '        BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetAllSite]", BaseConn.cnn)
    '        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    '        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    '        ds = New DataSet
    '        BaseConn.da.Fill(ds)
    '        GetAllSites = ds.Tables(0)
    '    Catch ex As Exception
    '        ErrNo = 1
    '    End Try
    '    Return GetAllSites
    'End Function

    'Public Sub GETSiteDefaultValues(ByRef ObjSiteDefault As csSiteDefaults, ByVal SiteID As String, ByVal LoggedUserName As String, ByRef dt_ConfigParam As DataTable, ByRef dt_menu As DataTable, ByRef dt_menuoptions As DataTable, ByRef ErrNo As Integer)
    '    ErrNo = 0
    '    Try

    '        BaseConn.Open()
    '        BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetSiteValues]", BaseConn.cnn)
    '        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    '        BaseConn.cmd.Parameters.AddWithValue("@SiteID", SiteID)
    '        BaseConn.cmd.Parameters.AddWithValue("@UserName", LoggedUserName)
    '        BaseConn.cmd.Parameters.Add("@DecimalPlace", SqlDbType.Int).Direction = ParameterDirection.Output
    '        BaseConn.cmd.Parameters.Add("@BusinessPeriodId", SqlDbType.Int).Direction = ParameterDirection.Output
    '        BaseConn.cmd.Parameters.Add("@BusinessStartDate", SqlDbType.Date).Direction = ParameterDirection.Output
    '        'BaseConn.cmd.ExecuteNonQuery()
    '        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    '        Dim ds As New DataSet
    '        BaseConn.da.Fill(ds)
    '        dt_ConfigParam = ds.Tables(0)
    '        dt_menu = ds.Tables(1)
    '        dt_menuoptions = ds.Tables(2)

    '        'Adding Variable Values
    '        ObjSiteDefault.DecimalPlace = BaseConn.cmd.Parameters("@DecimalPlace").Value
    '        ObjSiteDefault.BusinessPeriodID = BaseConn.cmd.Parameters("@BusinessPeriodId").Value
    '        ObjSiteDefault.BusinessStartDate = BaseConn.cmd.Parameters("BusinessStartDate").Value

    '    Catch ex As Exception
    '        ErrNo = 1
    '    End Try

    'End Sub

    Public Sub GETSiteDefaultValues(ByVal _CID As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _AppType As Integer, ByVal _GroupID As String,
                                    ByRef _dt_ConfigParam As DataTable, ByRef _GrpGenParam As DataTable, ByRef _LicenseDetails As DataTable, ByRef _dt_menu As DataTable,
                                    ByRef _dt_menuoptions As DataTable, ByRef _DTOrgMenu As DataTable, ByRef _DTGrpSalesMan As DataTable,
                                    ByRef _BusinessStartDate As Date, ByRef _DTBSPeriod As DataTable, ByRef _DTLanguageToken As DataTable, ByRef _DTErrorMessage As DataTable, _UserName As String, ByRef _ErrNo As Integer)
        _ErrNo = 0
        Try

            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetSiteValues]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@GroupID", _GroupID)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", _UserName)
            BaseConn.cmd.Parameters.AddWithValue("@ApplicationType", _AppType)
            BaseConn.cmd.Parameters.Add("@BusinessPeriodID", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@BusinessStartDate", SqlDbType.Date).Direction = ParameterDirection.Output
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            _dt_ConfigParam = ds.Tables(0)
            _dt_menu = ds.Tables(1)
            _dt_menuoptions = ds.Tables(2)
            _DTBSPeriod = ds.Tables(3)
            _DTGrpSalesMan = ds.Tables(4)
            _GrpGenParam = ds.Tables(5)
            _LicenseDetails = ds.Tables(6)
            _DTLanguageToken = ds.Tables(7)
            _DTErrorMessage = ds.Tables(8)
            _BusinessStartDate = BaseConn.cmd.Parameters("@BusinessStartDate").Value
        Catch ex As Exception
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub

    'Public Sub getAllSite(ByVal _SiteID As String, ByVal _DBPath As String, ByVal _DBPwd As String, ByRef _RemoteSiteGroup As DataTable, _
    '                         ByRef _LoggedUserGrpID As String, ByRef _LoggedGrpName As String, _
    '                         ByVal strUserName As String, ByVal strUserPwd As String, ByRef _ErrNo As Integer, ByRef _ErrString As String)

    '    _ErrNo = 0
    '    Try
    '        BaseConn.Open(_DBPath, _DBPwd)
    '        BaseConn.cmd = New SqlClient.SqlCommand("[sp_UserValidate]", BaseConn.cnn)
    '        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    '        BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
    '        BaseConn.cmd.Parameters.AddWithValue("@UserName", strUserName)
    '        BaseConn.cmd.Parameters.AddWithValue("@Password", strUserPwd)
    '        BaseConn.cmd.Parameters.Add("@GroupID", SqlDbType.Int).Direction = ParameterDirection.Output
    '        BaseConn.cmd.Parameters.Add("@GroupName", SqlDbType.Int).Direction = ParameterDirection.Output
    '        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    '        Dim ds As New DataSet
    '        BaseConn.da.Fill(ds)
    '        _LoggedUserGrpID = BaseConn.cmd.Parameters("@GroupID").Value
    '        _LoggedGrpName = BaseConn.cmd.Parameters("@GroupName").Value

    '        'DTRemoteGroup = ds.Tables(0)
    '        _RemoteSiteGroup = ds.Tables(0)

    '        'For Each drow1 In dt.Rows
    '        '    Dim strAloowSite As String = drow1("SiteID").ToString
    '        '    For Each drow In _DTConfigSite.Rows
    '        '        If drow("SiteID").ToString = strSiteID Then
    '        '            drow("LoginSiteID") = True
    '        '            drow("IsAllowSite") = True
    '        '            _DTConfigSite.AcceptChanges()
    '        '        End If
    '        '        If drow("SiteID").ToString = strAloowSite Then
    '        '            drow("IsAllowSite") = True
    '        '            _DTConfigSite.AcceptChanges()
    '        '        End If
    '        '    Next
    '        'Next


    '    Catch ex As Exception
    '        _ErrNo = 1
    '        Dim SPErrString As String = ex.Message.ToString
    '        If SPErrString = "2" Then
    '            _ErrString = "Invalid UserID"
    '        ElseIf SPErrString = "3" Then
    '            _ErrString = "Invalid Password"
    '        ElseIf SPErrString = "4" Then
    '            _ErrString = "User Locked, Please contact the Admin User"
    '        End If
    '    Finally
    '        BaseConn.Close()
    '    End Try
    'End Sub
End Class
