'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes

Public Class DAL_UserMgt
    Private ObjDalGeneral As DAL_General
    Dim dt As DataTable
    'Dim objcsIndent As New csIndent
    Dim BaseConn As New SQLConn()

    Public Sub Get_Structure(ByRef Obj As csUserMgt, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal ErrNo As String, ByVal ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetUserMgtDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@UserID", Obj.int_UserID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.int_SiteID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            If Obj.int_UserID <> 0 Then
                If ds.Tables(1).Rows.Count > 0 Then
                    Obj.dt_UserMgt = ds.Tables(0)
                    'Obj.dt_UserDynSettings = ds.Tables(2)
                End If
                Obj.str_Password = ds.Tables(1).Rows(0)("Password").ToString()
                Obj.str_GroupID = ds.Tables(1).Rows(0)("GroupID").ToString()
                Obj.bool_InActive = ds.Tables(1).Rows(0)("InActive").ToString()
                Obj.str_DefaultSiteID = ds.Tables(1).Rows(0)("DefaultSite").ToString()
                Obj.bool_ShowPopUp = ds.Tables(1).Rows(0)("ShowPopUp").ToString()
                Obj.int_EmployeeLedgerID = ds.Tables(1).Rows(0)("LedgerID").ToString()
                Obj.int_LanguageCode = ds.Tables(1).Rows(0)("DefaultLngCode").ToString()
                Obj.str_HeaderandButtonBackColor = ds.Tables(1).Rows(0)("ERPMainColor").ToString()
                Obj.str_FormbackColor = ds.Tables(1).Rows(0)("ERPSecondaryColor").ToString()
                Obj.str_ActiveDirectoryPath = ds.Tables(1).Rows(0)("ActiveDirectoryPath").ToString()
                Obj.str_ActiveDirectoryDomain = ds.Tables(1).Rows(0)("ActiveDirectoryDomain").ToString()
                Obj.str_ActiveDirectoryUserID = ds.Tables(1).Rows(0)("ActiveDirectoryUserID").ToString()
            Else
                Obj.dt_UserMain = ds.Tables(0)
            End If
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub


    Public Function Update_UserMgt(ByVal obj As csUserMgt, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("UserMgtUpdated", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.int_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@UserID", obj.int_UserID)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", obj.str_UserName)
            BaseConn.cmd.Parameters.AddWithValue("@UserIDOld", obj.int_UserID_Old)
            BaseConn.cmd.Parameters.AddWithValue("@Password", obj.str_Password)
            BaseConn.cmd.Parameters.AddWithValue("@GroupID", obj.str_GroupID)
            BaseConn.cmd.Parameters.AddWithValue("@InActive", obj.bool_InActive)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.int_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@DefaultSite", obj.str_DefaultSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.int_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ShowPopUp", obj.bool_ShowPopUp)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.int_EmployeeLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@MainColor", obj.str_HeaderandButtonBackColor)
            BaseConn.cmd.Parameters.AddWithValue("@SecondaryColor", obj.str_FormbackColor)
            'BaseConn.cmd.Parameters.AddWithValue("@SalesManID", obj.str_SalesManID)
            BaseConn.cmd.Parameters.AddWithValue("@LngCode", obj.int_LanguageCode)
            BaseConn.cmd.Parameters.AddWithValue("@ActiveDirectoryPath", obj.str_ActiveDirectoryPath)
            BaseConn.cmd.Parameters.AddWithValue("@ActiveDirectoryDomain", obj.str_ActiveDirectoryDomain)
            BaseConn.cmd.Parameters.AddWithValue("@ActiveDirectoryUserID", obj.str_ActiveDirectoryUserID)

            BaseConn.cmd.Parameters.AddWithValue("@SiteAccessDT", obj.dt_UserMgt)
            'BaseConn.cmd.Parameters.AddWithValue("@UserDynSettings", obj.dt_UserDynSettings)

            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.int_SiteID)
            ObjDalGeneral.Elog_Insert(obj.int_SiteID, _StrDBPath, _StrDBPwd, 0, obj.str_UserName, obj.dtp_LastUpdatedDate, "", "UserMgt", Err.Number, "Error in " & obj.str_Flag & " : " & obj.str_UserName & "  ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_UserMgt = _ErrString
    End Function


    Public Sub Get_FormsForDynamicSettings(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _TagID As String, ByRef _DTForms As DataTable, ByRef ErrNo As String, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetFormsForDynamicSettings]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@TagID", _TagID)

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            _DTForms = ds.Tables(0)

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function UserMgtDynSettingsUpdate(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal obj As csUserMgtDynSettings, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("sp_UserMgtDynSettingsUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", obj.str_UserName)
            BaseConn.cmd.Parameters.AddWithValue("@Functionality", obj.str_Functionality)
            BaseConn.cmd.Parameters.AddWithValue("@Module", obj.str_Module)
            BaseConn.cmd.Parameters.AddWithValue("@Enable", obj.bool_Enable)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.str_Comment)

            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, 0, obj.str_UserName, obj.dtp_ApprovedDate, "", "UserMgt", Err.Number, "Error in " & obj.str_Flag & " : " & obj.str_UserName & "  ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        UserMgtDynSettingsUpdate = _ErrString
    End Function

    Public Sub GetDynamicSettingsForUser(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _UserName As String, _
                                         ByVal _Functionality As String, ByVal _Module As String, ByRef _DTForms As DataTable, ByRef ErrNo As String, _
                                         ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetDynamicSettingsForUser]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", _UserName)
            BaseConn.cmd.Parameters.AddWithValue("@Functionality", _Functionality)
            BaseConn.cmd.Parameters.AddWithValue("@Module", _Module)

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            _DTForms = ds.Tables(0)

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetUserDynamicDetails(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _UserName As String,
                                        ByVal _Functionality As String, ByVal _Module As String, ByVal _Enable As Boolean, ByVal _DateType As String,
                                        ByVal _FromDate As Date, ByVal _ToDate As Date, ByRef _DTForms As DataTable, ByRef ErrNo As Integer,
                                        ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetUserDynamicDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", _UserName)
            BaseConn.cmd.Parameters.AddWithValue("@Functionality", _Functionality)
            BaseConn.cmd.Parameters.AddWithValue("@Module", _Module)
            BaseConn.cmd.Parameters.AddWithValue("@Enable", _Enable)
            BaseConn.cmd.Parameters.AddWithValue("@Date1", _DateType)
            BaseConn.cmd.Parameters.AddWithValue("@Fromdate", _FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@Todate", _ToDate)

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            _DTForms = ds.Tables(0)

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetUserDetails(_StrDBPath As String, _StrDBPwd As String, cid As Integer, username As String, pwd As String, ADDomain As String, ADLogin As Boolean, ByRef errno As Integer,
                              ByRef errdesc As String, ByRef dtUserdetails As DataTable)

        errno = 0
        errdesc = String.Empty
        dtUserdetails = New DataTable
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetUserDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", cid)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", username)
            BaseConn.cmd.Parameters.AddWithValue("@ADDomain", ADDomain)
            BaseConn.cmd.Parameters.AddWithValue("@ActiveDirectoryLogin", ADLogin)
            BaseConn.cmd.Parameters.Add("@ElogDate", SqlDbType.Date).Direction = ParameterDirection.Output

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            'Dim ds As New DataSet
            BaseConn.da.Fill(dtUserdetails)
            Dim _MaxElogDate As DateTime
            _MaxElogDate = Convert.ToDateTime(BaseConn.cmd.Parameters("@Elogdate").Value)
        Catch ex As Exception
            errno = 1
            errdesc = ex.ToString()
        Finally
            BaseConn.Close()
        End Try

    End Sub

    Public Function GetConfigParam(_StrDBPath As String, _StrDBPwd As String, ByVal cid As Integer) As DataTable
        Dim dtConfigParam As New DataTable

        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[MA_GetConfigParam]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", cid)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            'Dim ds As New DataSet
            BaseConn.da.Fill(dtConfigParam)
        Catch ex As Exception

        Finally
            BaseConn.Close()
        End Try

        Return dtConfigParam
    End Function

    Public Function GetSalesmanIDByLedgerID(_StrDBPath As String, _StrDBPwd As String, ByVal cid As Integer, ByVal ledgerid As Integer) As Integer
        Dim salesmanid As Integer

        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[MA_GetSalesmanIDByLedgerID]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", cid)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", ledgerid)
            BaseConn.cmd.Parameters.Add("@SalesmanID", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            salesmanid = BaseConn.cmd.Parameters("@SalesmanID").Value
        Catch ex As Exception

        Finally
            BaseConn.Close()
        End Try

        Return salesmanid
    End Function
End Class
