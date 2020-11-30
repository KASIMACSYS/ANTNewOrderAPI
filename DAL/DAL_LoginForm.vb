'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Imports Classes
Imports System.Data.OleDb

Public Class DAL_LoginForm
    Dim BaseConn As New SQLConn()

    'Public Function GetBusinessPeriodID(ByVal SiteID As String, ByRef intBusinessPeriodID As Integer) As Integer
    '    'BaseConn.Open()
    '    'BaseConn.cmd = New SqlClient.SqlCommand("select max(BusinessPeriodID) from [" + SiteID + "_BusinessPeriodMaster]", BaseConn.cnn)
    '    'BaseConn.cmd.CommandType = CommandType.Text
    '    'BaseConn.dr = BaseConn.cmd.ExecuteReader()
    '    'If BaseConn.dr.HasRows Then
    '    '    BaseConn.dr.Read()
    '    '    intBusinessPeriodID = BaseConn.dr(0).ToString
    '    'End If
    'End Function

    Public Function ValidateUser(ByVal SiteID As String, ByVal StrPath As String, ByVal strPwd As String, ByVal UserID As String, ByVal Password As String, ByRef errNo As Integer) As String
        Dim ErrStr As String = "Login Successful"
        errNo = 0
        Try
            BaseConn.Open(StrPath, strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("select Password, Active from [" + SiteID + "_UserMgt] where UserName=@user", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.Text
            BaseConn.cmd.Parameters.AddWithValue("@User", UserID)
            BaseConn.dr = BaseConn.cmd.ExecuteReader()
            If BaseConn.dr.HasRows Then
                BaseConn.dr.Read()
                If BaseConn.dr("Active") = True Then
                    If Password <> BaseConn.dr("Password").ToString Then
                        errNo = 1
                        ErrStr = "Invalid Password"
                    End If
                Else
                    errNo = 2
                    ErrStr = "User Locked, Please contact the Admin User"
                End If
            Else
                errNo = 3
                ErrStr = "Invalid User ID"
            End If
        Catch ex As Exception
            ErrStr = ex.Message.ToString
        End Try
        Return ErrStr
    End Function

    Public Sub GetUserDetails(ByRef _DBPath As String, ByRef _DBPwd As String, ByRef _CID As Integer, ByRef _User As String, ByRef _ADDomain As String, ByRef ActiveDirectoryLogin As Boolean,
                              ByRef _dtUserDetails As DataTable, ByRef _MaxElogDate As Date, ByRef _ErrNo As Integer, ByRef _ErrString As String)

        _ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetUserDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", _User)
            BaseConn.cmd.Parameters.AddWithValue("@ADDomain", _ADDomain)
            BaseConn.cmd.Parameters.AddWithValue("@ActiveDirectoryLogin", ActiveDirectoryLogin)
            BaseConn.cmd.Parameters.Add("@Elogdate", SqlDbType.Date).Direction = ParameterDirection.Output
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            _dtUserDetails = ds.Tables(0)
            _MaxElogDate = BaseConn.cmd.Parameters("@Elogdate").Value
        Catch ex As Exception
            _ErrNo = 1
            Dim SPErrString As String = ex.Message.ToString
            If SPErrString = "2" Then
                _ErrString = "Invalid UserID"
                _ErrNo = 2
            End If
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub getAllSite(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _CID As String, ByRef _RemoteSiteByUser As DataTable, ByRef _RemoteSiteWithGroup As DataTable,
                               ByRef _MenuMgt As DataTable, ByVal _UserID As String, ByRef _ErrNo As Integer, ByRef _ErrString As String)

        _ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[UserValidate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)

            BaseConn.cmd.Parameters.AddWithValue("@UserID", _UserID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            _RemoteSiteByUser = ds.Tables(0)
            _RemoteSiteWithGroup = ds.Tables(1)
            _MenuMgt = ds.Tables(2)


        Catch ex As Exception
            _ErrNo = 1
            Dim SPErrString As String = ex.Message.ToString
            If SPErrString = "2" Then
                _ErrString = "Invalid UserID"
                _ErrNo = 2
            ElseIf SPErrString = "3" Then
                _ErrString = "Invalid Password"
                _ErrNo = 3
            ElseIf SPErrString = "4" Then
                _ErrString = "User Locked, Please contact the Admin User"
                _ErrNo = 4
            End If
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetDomainName(ByVal DBPath As String, ByVal DBPwd As String, ByVal CID As Integer, ByRef DomainName As String, ByRef ErrString As String)
        ErrString = String.Empty
        Try
            BaseConn.Open(DBPath, DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetDomainName]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.Add("@Domain", SqlDbType.VarChar, 20).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()

            DomainName = BaseConn.cmd.Parameters("@Domain").Value
        Catch ex As Exception
            ErrString = ex.ToString
        End Try

    End Sub

    'Friend Function DecryptText(ByRef strText As String) As Object
    '    Dim i, c As Short
    '    Dim strBuff As String = ""
    '    Dim strpwd As String = "acsysit"
    '    'Decrypt string
    '    If Len(strpwd) Then
    '        For i = 1 To Len(strText)
    '            c = Asc(Mid(strText, i, 1))
    '            c = c - Asc(Mid(strpwd, (i Mod Len(strpwd)) + 1, 1))
    '            strBuff = strBuff & Chr(c And &HFFS)
    '        Next i
    '    Else
    '        strBuff = strText
    '    End If
    '    'UPGRADE_WARNING: Couldn't resolve default property of object DecryptText. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '    DecryptText = strBuff
    'End Function


    ' ''Public Sub GETSiteDefaultValues(ByRef ObjSiteDefault As csSiteDefaults, ByVal SiteID As String, ByVal strDBPath As String, ByVal strPwd As String, ByVal _LoginSiteID As String, ByVal GroupName As String, ByRef dt_ConfigParam As DataTable, ByRef dt_menu As DataTable, ByRef dt_menuoptions As DataTable, ByRef ErrNo As Integer)
    ' ''    ErrNo = 0
    ' ''    Try
    ' ''        BaseConn.Open(strDBPath, strPwd)
    ' ''        BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetSiteValues]", BaseConn.cnn)
    ' ''        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    ' ''        BaseConn.cmd.Parameters.AddWithValue("@SiteID", SiteID)
    ' ''        BaseConn.cmd.Parameters.AddWithValue("@LoginSiteID", _LoginSiteID)
    ' ''        BaseConn.cmd.Parameters.AddWithValue("@GroupName", GroupName)
    ' ''        BaseConn.cmd.Parameters.AddWithValue("@Flag", "SiteDefault")
    ' ''        BaseConn.cmd.Parameters.Add("@DecimalPlace", SqlDbType.Int).Direction = ParameterDirection.Output
    ' ''        BaseConn.cmd.Parameters.Add("@BusinessPeriodId", SqlDbType.Int).Direction = ParameterDirection.Output
    ' ''        'BaseConn.cmd.ExecuteNonQuery()
    ' ''        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    ' ''        Dim ds As New DataSet
    ' ''        BaseConn.da.Fill(ds)
    ' ''        dt_ConfigParam = ds.Tables(0)
    ' ''        dt_menu = ds.Tables(1)
    ' ''        dt_menuoptions = ds.Tables(2)
    ' ''        'Adding Variable Values
    ' ''        ObjSiteDefault.DecimalPlace = BaseConn.cmd.Parameters("@DecimalPlace").Value
    ' ''        ObjSiteDefault.BusinessPeriodID = BaseConn.cmd.Parameters("@BusinessPeriodId").Value
    ' ''    Catch ex As Exception
    ' ''        ErrNo = 1
    ' ''    End Try

    ' ''End Sub

    'Public Sub GETSiteDefaultValues(ByRef ObjSiteDefault As csSiteDefaults, ByVal SiteID As String, ByVal strDBPath As String, ByVal strPwd As String, ByVal _LoginSiteID As String, ByVal _RemoteSiteID As String, ByVal GroupID As String, ByRef dt_ConfigParam As DataTable, ByRef dt_menu As DataTable, ByRef dt_menuoptions As DataTable, ByRef ErrNo As Integer)
    '    ErrNo = 0
    '    Try
    '        BaseConn.Open(strDBPath, strPwd)
    '        BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetSiteValues]", BaseConn.cnn)
    '        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    '        BaseConn.cmd.Parameters.AddWithValue("@SiteID", SiteID)
    '        BaseConn.cmd.Parameters.AddWithValue("@LoginSiteID", _LoginSiteID)
    '        BaseConn.cmd.Parameters.AddWithValue("@RemoteSiteID", _RemoteSiteID)
    '        BaseConn.cmd.Parameters.AddWithValue("@GroupID", GroupID)
    '        BaseConn.cmd.Parameters.Add("@DecimalPlace", SqlDbType.Int).Direction = ParameterDirection.Output
    '        BaseConn.cmd.Parameters.Add("@BusinessPeriodID", SqlDbType.Int).Direction = ParameterDirection.Output
    '        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    '        Dim ds As New DataSet
    '        BaseConn.da.Fill(ds)
    '        dt_ConfigParam = ds.Tables(0)
    '        If _RemoteSiteID <> "" Then
    '            dt_menu = ds.Tables(1)
    '            dt_menuoptions = ds.Tables(2)
    '        End If
    '        ObjSiteDefault.DecimalPlace = BaseConn.cmd.Parameters("@DecimalPlace").Value
    '        ObjSiteDefault.BusinessPeriodID = BaseConn.cmd.Parameters("@BusinessPeriodID").Value
    '    Catch ex As Exception
    '        ErrNo = 1
    '    Finally
    '        BaseConn.Close()
    '    End Try
    'End Sub
End Class
