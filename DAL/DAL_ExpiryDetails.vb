'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Imports Classes
Public Class DAL_ExpiryDetails
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    'Public Function Get_Structure(ByVal Obj As csExpiryDetails) As csExpiryDetails
    '    Try
    '        BaseConn.Open()
    '        dt = New DataTable
    '        BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetExpiryDetails]", BaseConn.cnn)
    '        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    '        BaseConn.cmd.Parameters.AddWithValue("@SiteID", Obj.str_SiteID)
    '        BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPeriodID)
    '        BaseConn.cmd.Parameters.AddWithValue("@DocumentType", Obj.str_DocumentType)
    '        BaseConn.cmd.Parameters.AddWithValue("@Type", Obj.str_Type)
    '        BaseConn.cmd.Parameters.AddWithValue("@Name", Obj.str_Name)
    '        BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.str_Flag)
    '        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    '        BaseConn.da.Fill(dt)
    '        If dt.Rows.Count > 0 Then
    '            Obj.dt_ExpiryDetails = dt
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    Finally
    '        BaseConn.Close()
    '    End Try
    '    Return Get_Structure
    'End Function
End Class
