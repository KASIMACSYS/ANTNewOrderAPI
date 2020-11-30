'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Imports Classes

Public Class DAL_ProjectMaster
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef Obj As csProjectMaster, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetProjectDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", 101)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectID", Obj.ObjProject.str_ProjectID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.ObjProjCommon.str_Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            If Obj.ObjProjCommon.str_Flag = "PROJECT" Then
                Obj.ObjProject.str_Description = ds.Tables(0).Rows(0)("Description").ToString()
                Obj.ObjProject.str_MerchantID = ds.Tables(0).Rows(0)("MerchantID").ToString()
                Obj.ObjProject.str_City = ds.Tables(0).Rows(0)("City").ToString()
                Obj.ObjProject.str_State = ds.Tables(0).Rows(0)("State").ToString()
                Obj.ObjProject.str_Country = ds.Tables(0).Rows(0)("Country").ToString()
                Obj.ObjProject.bool_Status = ds.Tables(0).Rows(0)("Status").ToString()
                Obj.ObjProject.dtp_StartDate = ds.Tables(0).Rows(0)("StartDate").ToString()
                Obj.ObjProject.dtp_EndDate = ds.Tables(0).Rows(0)("EndDate").ToString()
                Obj.ObjProject.dbl_BudgetAmount = ds.Tables(0).Rows(0)("BudgetAmount").ToString()
                Obj.ObjProject.int_EstimatedManDay = ds.Tables(0).Rows(0)("EstimatedManDay").ToString()
                Obj.ObjProject.int_DayHours = ds.Tables(0).Rows(0)("DayHours").ToString()
                Obj.ObjProject.int_ProductID = ds.Tables(0).Rows(0)("DependencyID").ToString()
                Obj.ObjProject.str_ContactPerson = ds.Tables(0).Rows(0)("ContactPerson").ToString()
                Obj.ObjProject.str_ContactDesignatin = ds.Tables(0).Rows(0)("ContactDesignation").ToString()
                Obj.ObjProject.str_ContactEmail = ds.Tables(0).Rows(0)("ContactEmail").ToString()
                Obj.ObjProject.str_ContactTelephone = ds.Tables(0).Rows(0)("ContactTelephone").ToString()
                Obj.ObjProject.str_ContactMobile = ds.Tables(0).Rows(0)("ContactMobile").ToString()
                Obj.ObjProject.str_DstLedgerID = ds.Tables(0).Rows(0)("DstLedgerID").ToString()
                Obj.ObjProject.int_PCCID = ds.Tables(0).Rows(0)("PCCID").ToString()
                If ds.Tables(1).Rows.Count > 0 Then
                    Obj.ObjProject.dt_LocationSub = ds.Tables(1)
                Else
                    If Obj.ObjProject.dt_LocationSub IsNot Nothing Then
                        Obj.ObjProject.dt_LocationSub.Clear()
                    End If
                End If
            End If

            If Obj.ObjProjCommon.str_Flag = "LOCATION" Then
                Obj.ObjLocation.str_ProjectLocation = ds.Tables(1).Rows(0)("ProjectLocation").ToString()
            End If

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Update_ProjMaster(ByVal obj As csProjectMaster, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("ProjectMasterUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID)

            BaseConn.cmd.Parameters.AddWithValue("@ProjectID", obj.ObjProject.str_ProjectID)
            BaseConn.cmd.Parameters.AddWithValue("@Description", obj.ObjProject.str_Description)
            BaseConn.cmd.Parameters.AddWithValue("@MerchantID", obj.ObjProject.str_MerchantID)
            BaseConn.cmd.Parameters.AddWithValue("@City", obj.ObjProject.str_City)
            BaseConn.cmd.Parameters.AddWithValue("@State", obj.ObjProject.str_State)
            BaseConn.cmd.Parameters.AddWithValue("@Country", obj.ObjProject.str_Country)
            BaseConn.cmd.Parameters.AddWithValue("@Status", obj.ObjProject.bool_Status)
            BaseConn.cmd.Parameters.AddWithValue("@StartDate", obj.ObjProject.dtp_StartDate)
            BaseConn.cmd.Parameters.AddWithValue("@EndDate", obj.ObjProject.dtp_EndDate)
            BaseConn.cmd.Parameters.AddWithValue("@BudgetAmount", obj.ObjProject.dbl_BudgetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@EstimatedManDay", obj.ObjProject.int_EstimatedManDay)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectLocation", obj.ObjLocation.str_ProjectLocation)
            BaseConn.cmd.Parameters.AddWithValue("@DayHours", obj.ObjProject.int_DayHours)
            BaseConn.cmd.Parameters.AddWithValue("@DependencyID", obj.ObjProject.int_ProductID)
            BaseConn.cmd.Parameters.AddWithValue("@ContactPerson", obj.ObjProject.str_ContactPerson)
            BaseConn.cmd.Parameters.AddWithValue("@ContactDesignation", obj.ObjProject.str_ContactDesignatin)
            BaseConn.cmd.Parameters.AddWithValue("@ContactEmail", obj.ObjProject.str_ContactEmail)
            BaseConn.cmd.Parameters.AddWithValue("@ContactTelephone", obj.ObjProject.str_ContactTelephone)
            BaseConn.cmd.Parameters.AddWithValue("@ContactMobile", obj.ObjProject.str_ContactMobile)
            BaseConn.cmd.Parameters.AddWithValue("@DstLedgerID", obj.ObjProject.str_DstLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@PCCID", obj.ObjProject.int_PCCID)

            BaseConn.cmd.Parameters.AddWithValue("@DT", obj.ObjProject.dt_LocationSub)

            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjProjCommon.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@ProjOrLocation", obj.ObjProjCommon.str_ProjOrLocation)
            BaseConn.cmd.Parameters.AddWithValue("@EditLocation", obj.ObjLocation.str_EditLocation)
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.int_BusinessPeriodID, "", DateTime.Now, "", "ProjecttMaster", Err.Number, "Error no " & obj.ObjProjCommon.str_Flag & " : " & obj.ObjProject.str_ProjectID & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        Update_ProjMaster = _ErrString
    End Function

End Class
