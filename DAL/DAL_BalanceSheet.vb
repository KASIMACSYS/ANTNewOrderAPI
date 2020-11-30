Imports Classes
Imports System.Data.SqlClient
Public Class DAL_BalanceSheet
    Dim dt, dt1 As DataTable
    Dim BaseConn As New SQLConn()
    Dim SiteID As String

    Public Sub New(ByVal siteid As String)
        Me.SiteID = siteid
    End Sub
    Public Function getBalanceSheet(ByVal Str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal int_BusinessPeroidID As Integer, ByVal dtp_from As Date, ByVal dtp_to As Date) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_BalanceSheet]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", int_BusinessPeroidID)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", dtp_from)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", dtp_to)
            BaseConn.cmd.CommandTimeout = 1000
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = ds.Tables(0)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
    Public Function getBalanceSheetDetails(ByVal Str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal int_BusinessPeroidID As Integer, ByVal _Flag As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetBalanceSheetDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", int_BusinessPeroidID) 'TO DO  Enable 
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.CommandTimeout = 1000
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = ds.Tables(0)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function getProfitandLoss(ByVal Str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal int_BusinessPeroidID As Integer, ByVal dtp_from As Date, ByVal dtp_to As Date) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_ProfitandLoss]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", int_BusinessPeroidID) 'TO DO  Enable 
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", dtp_from)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", dtp_to)
            BaseConn.cmd.CommandTimeout = 1000
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = ds.Tables(0)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
    'Public Function getTrailBalance(ByVal _strPath As String, ByVal _strPwd As String, ByVal _CID As Integer,
    '                                ByVal dtp_from As Date, ByVal dtp_to As Date, ByVal _ReportLevel As String, ByVal _ZeroSuppress As Boolean,
    '                                ByVal _ShowInActive As Boolean, ByVal _Type As Integer, ByVal _MenuID As String, ByVal _Ledgers As DataTable) As DataTable
    '    Try
    '        dt = New DataTable
    '        BaseConn.Open(_strPath, _strPwd)
    '        BaseConn.cmd = New SqlClient.SqlCommand("[GetTrialBalance]", BaseConn.cnn)
    '        'BaseConn.cmd = New SqlClient.SqlCommand("[sp_TrailBalance_Detailed]", BaseConn.cnn)
    '        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    '        BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
    '        BaseConn.cmd.Parameters.AddWithValue("@FromDate", dtp_from)
    '        BaseConn.cmd.Parameters.AddWithValue("@ToDate", dtp_to)
    '        BaseConn.cmd.Parameters.AddWithValue("@ReportLevel", _ReportLevel)
    '        BaseConn.cmd.Parameters.AddWithValue("@ZeroSuppress", _ZeroSuppress)
    '        BaseConn.cmd.Parameters.AddWithValue("@ShowInActive", _ShowInActive)
    '        BaseConn.cmd.Parameters.AddWithValue("@Type", _Type)
    '        BaseConn.cmd.Parameters.AddWithValue("@MenuID", _MenuID)
    '        BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _Ledgers)
    '        BaseConn.cmd.CommandTimeout = 1000
    '        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    '        Dim ds As New DataSet
    '        BaseConn.da.Fill(ds)
    '        dt = ds.Tables(0)
    '    Catch ex As Exception
    '        MsgBox("Error" & ex.Message)
    '    Finally
    '        BaseConn.Close()
    '    End Try
    '    Return dt
    'End Function

    Public Function getTrailBalance(ByVal Str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String,
                                    ByVal dtp_from As Date, ByVal dtp_to As Date, ByVal _ReportLevel As String, ByVal _ZeroSuppress As Boolean,
                                    ByVal _ShowInActive As Boolean, ByVal _Type As Integer, ByVal _FormName As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[TrialBalance_CS]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", dtp_from)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", dtp_to)
            BaseConn.cmd.Parameters.AddWithValue("@ReportLevel", _ReportLevel)
            BaseConn.cmd.Parameters.AddWithValue("@ZeroSuppress", _ZeroSuppress)
            BaseConn.cmd.Parameters.AddWithValue("@ShowInActive", _ShowInActive)
            BaseConn.cmd.Parameters.AddWithValue("@Type", _Type)
            BaseConn.cmd.Parameters.AddWithValue("@Form", _FormName)
            BaseConn.cmd.CommandTimeout = 1000
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = ds.Tables(0)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function getBalanceSheet(ByVal Str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, _
                                   ByVal dtp_from As Date, ByVal dtp_to As Date, ByVal _ReportLevel As String, ByVal _ZeroSuppress As Boolean, _
                                   ByVal _ShowInActive As Boolean) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_BalanceSheet_Detailed]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", dtp_from)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", dtp_to)
            BaseConn.cmd.Parameters.AddWithValue("@ReportLevel", _ReportLevel)
            BaseConn.cmd.Parameters.AddWithValue("@ZeroSuppress", _ZeroSuppress)
            BaseConn.cmd.Parameters.AddWithValue("@ShowInActive", _ShowInActive)
            BaseConn.cmd.CommandTimeout = 1000
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = ds.Tables(0)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

End Class
