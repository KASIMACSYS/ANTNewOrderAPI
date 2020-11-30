'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Imports Classes

Public Class DAL_StockValuation
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()

    'Public Sub Get_Structure(ByRef Obj As csStockValuation, ByRef ErrNo As Integer, ByRef ErrStr As String)
    '    ErrNo = 0
    '    ErrStr = ""
    '    Try
    '        BaseConn.Open()
    '        BaseConn.cmd = New SqlClient.SqlCommand("[sp_ParameterUpdate]", BaseConn.cnn)
    '        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    '        BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPerionID)
    '        BaseConn.cmd.Parameters.AddWithValue("@SiteID", Obj.str_SiteID)
    '        BaseConn.cmd.Parameters.AddWithValue("@Catagory", Obj.str_Catagory)
    '        BaseConn.cmd.Parameters.AddWithValue("@Type", Obj.str_Type)
    '        BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.str_Flag)
    '        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    '        BaseConn.da.Fill(dt)

    '        'Obj.dt_Types = ds.Tables(0)
    '        'Obj.dt_Parameter = ds.Tables(1)
    '        Obj.dt_Mccb = dt

    '    Catch ex As Exception
    '        ErrNo = 1
    '        ErrStr = ex.Message
    '    Finally
    '        BaseConn.Close()
    '    End Try
    'End Sub

End Class
