'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes
Public Class DAL_FormGridSettings
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Public Sub Get_Structure(ByRef Obj As csFormGridSettings, ByVal _strPath As String, ByVal _strPwd As String, ByRef ErrNo As Integer, ByVal ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetFormGridSettingsforForm]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.str_BusinessPerionID)
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", Obj.MenuID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = New DataTable
            BaseConn.da.Fill(dt)
            Obj.dt_Mccb1 = ds.Tables(0)
            Obj.dt_FormGridSetting = ds.Tables(1)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Function Put_Structure(ByVal Obj As csFormGridSettings, ByRef str_DocumentNo As String, ByVal _strPath As String, ByVal _strPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_FormGridSettingsUpdate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.str_BusinessPerionID)
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", Obj.MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@dt_FormGridSetting", Obj.dt_FormGridSetting)
            BaseConn.cmd.ExecuteNonQuery()
        Catch ex As Exception
            _ErrString = ex.Message
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Put_Structure = _ErrString
    End Function
End Class
