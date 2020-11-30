'======================================================================================
'$Author: Saravanan $
'$Rev: 674 $
'$Date: 2019-02-13 18:06:08 +0530 (Tue, 13 Feb 2019) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes
Public Class DAL_MaterialSampleOrder
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal str_SiteID As String, ByRef Obj As csMaterialSampleOrder, ByVal _MenuID As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[FE_GetMaterialSample]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@MaterialSampleID", Obj.ObjMaterialSampleOrderMain.str_MaterialSampleOrderID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.ObjMaterialSampleOrderMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.ObjMaterialSampleOrderMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", _MenuID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            Obj.ObjMaterialSampleOrderMain.str_JobOrderNo = ds.Tables(0).Rows(0)("JobOrderNo").ToString
            Obj.ObjMaterialSampleOrderMain.str_QtnNo = ds.Tables(0).Rows(0)("QtnNo").ToString
            Obj.ObjMaterialSampleOrderMain.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString
            Obj.ObjMaterialSampleOrderMain.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString
            Obj.ObjMaterialSampleOrderMain.str_Project = ds.Tables(0).Rows(0)("ProjectName").ToString
            Obj.ObjMaterialSampleOrderMain.str_Item = ds.Tables(0).Rows(0)("ItemName").ToString
            Obj.ObjMaterialSampleOrderMain.str_Coordinator = ds.Tables(0).Rows(0)("Coordinator").ToString
            Obj.ObjMaterialSampleOrderMain.str_Production = ds.Tables(0).Rows(0)("Production").ToString
            Obj.ObjMaterialSampleOrderMain.str_Location = ds.Tables(0).Rows(0)("Location").ToString
            Obj.ObjMaterialSampleOrderMain.str_BrandName = ds.Tables(0).Rows(0)("BrandName").ToString
            Obj.ObjMaterialSampleOrderMain.str_Contact = ds.Tables(0).Rows(0)("Contact").ToString
            Obj.ObjMaterialSampleOrderMain.str_Email = ds.Tables(0).Rows(0)("Email").ToString
            Obj.ObjMaterialSampleOrderMain.dtp_VoucherDate = ds.Tables(0).Rows(0)("VoucherDate").ToString
            Obj.ObjMaterialSampleOrderMain.dtp_IssueDate = ds.Tables(0).Rows(0)("IssueDate").ToString
            Obj.ObjMaterialSampleOrderMain.dtp_CompletionDate = ds.Tables(0).Rows(0)("CompletionDate").ToString
            Obj.ObjMaterialSampleOrderMain.str_CFCCompletionDate = ds.Tables(0).Rows(0)("CFCCompletionDate").ToString
            Obj.ObjMaterialSampleOrderMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString
            Obj.ObjMaterialSampleOrderMain.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
            Obj.ObjMaterialSampleOrderMain.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
            Obj.ObjMaterialSampleOrderMain.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
            Obj.ObjMaterialSampleOrderMain.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
            Obj.ObjMaterialSampleOrderMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()
            Obj.ObjMaterialSampleOrderMain.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
            Obj.ObjMaterialSampleOrderMain.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
            Obj.ObjMaterialSampleOrderMain.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus")
            If ds.Tables(1).Rows.Count > 0 Then
                Obj.ObjMaterialSampleOrderSub.dt_MaterialSampleOrder = ds.Tables(1)
            End If
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Function Update_MaterialSample(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal str_SiteID As String, ByVal obj As csMaterialSampleOrder, ByRef str_DocumentNo As String, ByRef outRevNo As Integer, ByRef ErrNo As Integer, ByRef ErrStr As String) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("FE_MaterialSampleUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.ObjMaterialSampleOrderMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Prefix", obj.ObjMaterialSampleOrderMain.str_Prefix)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.ObjMaterialSampleOrderMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.ObjMaterialSampleOrderMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@MaterialSampleID", obj.ObjMaterialSampleOrderMain.str_MaterialSampleOrderID)
            BaseConn.cmd.Parameters.AddWithValue("@JobOrderNo", obj.ObjMaterialSampleOrderMain.str_JobOrderNo)
            BaseConn.cmd.Parameters.AddWithValue("@qtnNo", obj.ObjMaterialSampleOrderMain.str_QtnNo)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.ObjMaterialSampleOrderMain.int_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Alias", obj.ObjMaterialSampleOrderMain.str_Alias)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectName", obj.ObjMaterialSampleOrderMain.str_Project)
            BaseConn.cmd.Parameters.AddWithValue("@ItemName", obj.ObjMaterialSampleOrderMain.str_Item)
            BaseConn.cmd.Parameters.AddWithValue("@Coordinator", obj.ObjMaterialSampleOrderMain.str_Coordinator)
            BaseConn.cmd.Parameters.AddWithValue("@Production", obj.ObjMaterialSampleOrderMain.str_Production)
            BaseConn.cmd.Parameters.AddWithValue("@Location", obj.ObjMaterialSampleOrderMain.str_Location)
            BaseConn.cmd.Parameters.AddWithValue("@BrandName", obj.ObjMaterialSampleOrderMain.str_BrandName)
            BaseConn.cmd.Parameters.AddWithValue("@Contact", obj.ObjMaterialSampleOrderMain.str_Contact)
            BaseConn.cmd.Parameters.AddWithValue("@Email", obj.ObjMaterialSampleOrderMain.str_Email)
            BaseConn.cmd.Parameters.AddWithValue("@VoucherDate", obj.ObjMaterialSampleOrderMain.dtp_VoucherDate)
            BaseConn.cmd.Parameters.AddWithValue("@IssueDate", obj.ObjMaterialSampleOrderMain.dtp_IssueDate)
            BaseConn.cmd.Parameters.AddWithValue("@CompletionDate", obj.ObjMaterialSampleOrderMain.dtp_CompletionDate)
            BaseConn.cmd.Parameters.AddWithValue("@CFCCompletionDate", obj.ObjMaterialSampleOrderMain.str_CFCCompletionDate)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.ObjMaterialSampleOrderMain.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.ObjMaterialSampleOrderMain.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.ObjMaterialSampleOrderMain.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.ObjMaterialSampleOrderMain.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.ObjMaterialSampleOrderMain.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.ObjMaterialSampleOrderMain.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.ObjMaterialSampleOrderMain.bool_ApprovedStatus)
            BaseConn.cmd.Parameters.AddWithValue("@MaterialSampleDT", obj.ObjMaterialSampleOrderSub.dt_MaterialSampleOrder)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjMaterialSampleOrderMain.str_Flag)
            BaseConn.cmd.Parameters.Add("@MaterialSampleIDOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            str_DocumentNo = BaseConn.cmd.Parameters("@MaterialSampleIDOut").Value.ToString
            outRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.ObjMaterialSampleOrderMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "StockTransfer", Err.Number, "Error in '" & obj.ObjMaterialSampleOrderMain.str_Flag & "'ED '" & obj.ObjMaterialSampleOrderMain.str_MaterialSampleOrderID & "' ", ex.Message, 5, 3, 1, 0)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
        Update_MaterialSample = _ErrString
    End Function

    Public Function Get_POTStatusDetails(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal str_SiteID As String, ByVal Obj As csMaterialSampleOrder, ByVal _MenuID As String) As csMaterialSampleOrder
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetPOTStatusDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure

            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.ObjMaterialSampleOrderMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@POTNo", Obj.ObjMaterialSampleOrderMain.str_POTNo)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            Obj.ObjMaterialSampleOrderMain.str_POTRef = ds.Tables(0).Rows(0)("POTRef").ToString
            Obj.ObjMaterialSampleOrderMain.str_Project = ds.Tables(0).Rows(0)("Project").ToString
            Obj.ObjMaterialSampleOrderMain.str_JobOrderNo = ds.Tables(0).Rows(0)("JONo").ToString
            Obj.ObjMaterialSampleOrderMain.str_Item = ds.Tables(0).Rows(0)("ItemDesc").ToString
            Obj.ObjMaterialSampleOrderMain.dtp_VoucherDate = ds.Tables(0).Rows(0)("VouDate").ToString
            Obj.ObjMaterialSampleOrderMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString

            Obj.ObjMaterialSampleOrderMain.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
            Obj.ObjMaterialSampleOrderMain.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
            Obj.ObjMaterialSampleOrderMain.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
            Obj.ObjMaterialSampleOrderMain.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
            Obj.ObjMaterialSampleOrderMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()
            Obj.ObjMaterialSampleOrderMain.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
            Obj.ObjMaterialSampleOrderMain.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
            Obj.ObjMaterialSampleOrderMain.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus")
            If ds.Tables(1).Rows.Count > 0 Then
                Obj.ObjMaterialSampleOrderSub.dt_MaterialSampleOrder = ds.Tables(1)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Get_POTStatusDetails = Obj
        Return Get_POTStatusDetails
    End Function
    Public Function Update_POTStatusDetails(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal str_SiteID As String, ByVal obj As csMaterialSampleOrder, ByRef str_DocumentNo As String, ByRef outRevNo As Integer, ByRef ErrNo As Integer, ByRef ErrStr As String) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("sp_POTStatusDetailsUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.ObjMaterialSampleOrderMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Prefix", obj.ObjMaterialSampleOrderMain.str_Prefix)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.ObjMaterialSampleOrderMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.ObjMaterialSampleOrderMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@POTNo", obj.ObjMaterialSampleOrderMain.str_POTNo)
            BaseConn.cmd.Parameters.AddWithValue("@JobOrderNo", obj.ObjMaterialSampleOrderMain.str_JobOrderNo)
            BaseConn.cmd.Parameters.AddWithValue("@POTRef", obj.ObjMaterialSampleOrderMain.str_POTRef)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectName", obj.ObjMaterialSampleOrderMain.str_Project)
            BaseConn.cmd.Parameters.AddWithValue("@ItemName", obj.ObjMaterialSampleOrderMain.str_Item)
            BaseConn.cmd.Parameters.AddWithValue("@VoucherDate", obj.ObjMaterialSampleOrderMain.dtp_VoucherDate)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.ObjMaterialSampleOrderMain.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.ObjMaterialSampleOrderMain.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.ObjMaterialSampleOrderMain.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.ObjMaterialSampleOrderMain.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.ObjMaterialSampleOrderMain.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.ObjMaterialSampleOrderMain.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.ObjMaterialSampleOrderMain.bool_ApprovedStatus)
            BaseConn.cmd.Parameters.AddWithValue("@POTStatusDetailsDT", obj.ObjMaterialSampleOrderSub.dt_MaterialSampleOrder)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjMaterialSampleOrderMain.str_Flag)
            BaseConn.cmd.Parameters.Add("@POTNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            str_DocumentNo = BaseConn.cmd.Parameters("@POTNoOut").Value.ToString
            outRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.ObjMaterialSampleOrderMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "POTStatus", Err.Number, "Error in '" & obj.ObjMaterialSampleOrderMain.str_Flag & "'ED '" & obj.ObjMaterialSampleOrderMain.str_MaterialSampleOrderID & "' ", ex.Message, 5, 3, 1, 0)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
        Update_POTStatusDetails = _ErrString
    End Function
End Class
