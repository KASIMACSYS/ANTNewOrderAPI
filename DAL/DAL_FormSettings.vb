'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes
Public Class DAL_FormSettings
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General
    Public Sub Get_Structure(ByRef Obj As csFormSetting, ByVal _StrDBPath As String, ByVal _strDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _strDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetFormSettings]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.str_BusinessPerionID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", Obj.MenuID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = New DataTable
            BaseConn.da.Fill(dt)
            Obj.dt_PropertyMenuID = ds.Tables(0)
            Obj.dt_DefaultForm = ds.Tables(1)
            Obj.dt_GridMenuID = ds.Tables(2)
            Obj.dt_GridSetting = ds.Tables(3)
            Obj.dt_ReportMenuID = ds.Tables(4)
            Obj.dt_ReportSetting = ds.Tables(5)
            Obj.dt_PropertySubMenuID = ds.Tables(6)
            Obj.dt_PropertySub = ds.Tables(7)
            Obj.dt_ApprovalSetting = ds.Tables(8)
            Obj.dt_PropertySortViewMenuID = ds.Tables(9)
            Obj.dt_PropertySortView = ds.Tables(10)

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Function Put_Structure(ByVal Obj As csFormSetting, ByRef str_DocumentNo As String, ByVal _StrDBPath As String, ByVal _strDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _strDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[FormSettings]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.str_BusinessPerionID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", Obj.MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@dt_DefaultForm", Obj.dt)
            BaseConn.cmd.Parameters.AddWithValue("@dt_FormGridSetting", Obj.dt_GridSetting)
            BaseConn.cmd.Parameters.AddWithValue("@dt_FormReportSetting", Obj.dt_ReportSetting)
            BaseConn.cmd.Parameters.AddWithValue("@dt_PropertysettingSub", Obj.dt_PropertySub)
            BaseConn.cmd.Parameters.AddWithValue("@dt_ApprovalSetting", Obj.dt_ApprovalSetting)
            BaseConn.cmd.Parameters.AddWithValue("@dt_PropertysettingSort", Obj.dt_PropertySortView)
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(Obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(Obj.str_SiteID, _StrDBPath, _strDBPwd, Obj.str_BusinessPerionID, "", DateTime.Now, "", "FormSettings", Err.Number, "Error in " & Obj.str_Flag & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
        Put_Structure = _ErrString
    End Function
End Class
