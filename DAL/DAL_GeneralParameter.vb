'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


'Imports Classes

'Public Class DAL_GeneralParameter
'    Dim dt As DataTable
'    Dim BaseConn As New SQLConn()

'    Public Function Get_Structure(ByVal Obj As csGeneralParameter, ByVal _StrDBPath As String, ByVal _StrDBPwd As String) As csGeneralParameter
'        Try
'            BaseConn.Open(_StrDBPath, _StrDBPwd)
'            BaseConn.cmd = New SqlClient.SqlCommand("[sp_ParameterUpdate]", BaseConn.cnn)
'            BaseConn.cmd.CommandType = CommandType.StoredProcedure
'            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.str_BusinessPerionID)
'            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Obj.str_SiteID)
'            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.str_Flag)
'            BaseConn.cmd.Parameters.AddWithValue("@Condition", Obj.str_Types)
'            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
'            Dim ds As New DataSet
'            BaseConn.da.Fill(ds)
'            Obj.dt_Types = ds.Tables(0)
'            'Obj.dt_Parameter = ds.Tables(1)
'            'Obj.dt_Mccb = ds.Tables(2)
'        Catch ex As Exception
'            MsgBox(ex.Message)
'        Finally
'            BaseConn.Close()
'        End Try
'        Return Get_Structure
'    End Function

'    Public Function Update_Parameter(ByVal obj As csGeneralParameter, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
'        Dim _ErrString As String = ""
'        ErrNo = 0
'        Try
'            BaseConn.Open(_StrDBPath, _StrDBPwd)
'            BaseConn.cmd = New SqlClient.SqlCommand("[sp_ParameterUpdate]", BaseConn.cnn)
'            BaseConn.cmd.CommandType = CommandType.StoredProcedure
'            BaseConn.cmd.Parameters.AddWithValue("@SiteID", obj.str_SiteID) 'obj.str_SiteID
'            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.str_BusinessPerionID)
'            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.str_Flag)
'            BaseConn.cmd.Parameters.AddWithValue("@Condition", obj.str_Types)
'            'BaseConn.cmd.Parameters.AddWithValue("@Combo_Name", obj.str_Types)
'            'BaseConn.cmd.Parameters.AddWithValue("@ParameterDT", obj.dt_Parameter)
'            BaseConn.cmd.Parameters.AddWithValue("@BaseTypeDT", obj.dt_Types)
'            'BaseConn.cmd.Parameters.AddWithValue("@MCCBDT", obj.dt_Mccb)
'            BaseConn.cmd.ExecuteNonQuery()
'        Catch ex As Exception
'            _ErrString = ex.Message
'            ErrNo = 1
'        Finally
'            BaseConn.Close()
'        End Try

'        Update_Parameter = _ErrString
'    End Function
'End Class
