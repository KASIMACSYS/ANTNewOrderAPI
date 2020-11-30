'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes
Public Class DAL_LedgerGrouping

    Dim dt As DataTable
    Dim BaseConn As New SQLConn()

    Public Function GetGrouping(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef _CID As Integer, ByRef _LedgerDT As Integer, ByRef _FormType As String,
                                ByRef _Flag As String, ByRef _ErrNo As String, ByRef _ErrStr As String) As DataTable
        GetGrouping = New DataTable
        _ErrNo = 0
        _ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            dt = New DataTable
            BaseConn.cmd = New SqlClient.SqlCommand("[GetGrouping]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerDT)
            BaseConn.cmd.Parameters.AddWithValue("@FormType", _FormType)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            GetGrouping = ds.Tables(0)
            Return GetGrouping
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Function

    Public Sub Get_Structure(ByVal obj As csLedgerGrouping, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef _ErrNo As String, ByRef _ErrStr As String)
        _ErrNo = 0
        _ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            dt = New DataTable
            BaseConn.cmd = New SqlClient.SqlCommand("[GetLedgerGrouping]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.int_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@ParentID", obj.int_ParentID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerDescription", obj.str_Description)
            BaseConn.cmd.Parameters.AddWithValue("@OnlyControlAC", obj.bool_OnlyControlAC)
            BaseConn.cmd.Parameters.AddWithValue("@OnlyClassAC", obj.bool_OnlyClassAC)
            BaseConn.cmd.Parameters.AddWithValue("@OnlyNominalAC", obj.bool_OnlyNominalAC)
            BaseConn.cmd.Parameters.AddWithValue("@ControlACwithClass", obj.bool_ControlACwithClass)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@FormType", obj.str_FormType)
            BaseConn.cmd.Parameters.AddWithValue("@Category", obj.str_Category)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            obj.dt_group = ds.Tables(0)
            If ds.Tables.Count > 1 Then
                obj.dt_LedgerDesc = ds.Tables(1)
            End If
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub ChangeLedgerDetails(ByVal Str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByRef _dt As DataTable, ByVal strParentLedgerID As String, ByVal _Flag As String, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        _ErrNo = 0
        _ErrStr = ""
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[ChangeLedgerDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@ParentLedgerID", strParentLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@DT", _dt)
            BaseConn.cmd.ExecuteNonQuery()
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
End Class
