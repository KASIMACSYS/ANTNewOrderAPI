'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Imports Classes

Public Class DAL_MasterMain
    Private BaseConn As New SQLConn()
    Private dt As DataTable
    'Public Sub Get_Structure(ByRef Obj As csMasterMain, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
    '    _ErrNo = 0
    '    _ErrStr = ""
    '    Try
    '        BaseConn.Open()
    '        dt = New DataTable
    '        BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetMasterDetails]", BaseConn.cnn)
    '        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    '        BaseConn.cmd.Parameters.AddWithValue("@TableName", Obj.str_TableName)
    '        BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPeriodID)
    '        BaseConn.cmd.Parameters.AddWithValue("@SiteID", Obj.str_SiteID)
    '        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    '        BaseConn.da.Fill(dt)
    '        Obj.dt_Master = dt
    '    Catch ex As Exception
    '        _ErrNo = 1
    '        _ErrStr = ex.Message
    '    Finally
    '        BaseConn.Close()
    '    End Try
    'End Sub
End Class
