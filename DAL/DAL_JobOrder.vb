Imports Classes

Public Class DAL_JobOrder
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByVal _DBPath As String, ByVal _DBPwd As String, ByRef Obj As csJobOrder, ByRef ErrNo As Integer, ByRef ErrMsg As String)
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetJODetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.objJOMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@JONo", Obj.objJOMain.str_JONo)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.objJOMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", Obj.objJOMain.str_FG)

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            Obj.objproject.str_ProjectID = ""
            Obj.objproject.str_ProjectLocation = ""
            Obj.objproject.str_WorkOrderNo = ""

            If Obj.objJOMain.str_Flag = "JO" Then
                Obj.objJOMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()
                Obj.objJOMain.dtp_JODate = ds.Tables(0).Rows(0)("JODate").ToString()
                Obj.objJOMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
                Obj.objJOMain.str_SONo = ds.Tables(0).Rows(0)("SONo").ToString()
                Obj.objJOMain.str_JODesc = ds.Tables(0).Rows(0)("JODesc").ToString()
                Obj.objJOMain.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
                Obj.objJOMain.int_SalesManID = ds.Tables(0).Rows(0)("SalesManID").ToString()
                Obj.objJOMain.str_ProdUnitName = ds.Tables(0).Rows(0)("ProdUnitName").ToString()

                Obj.objJOMain.dtp_EstEndDate = ds.Tables(0).Rows(0)("EstEndDate").ToString()
                Obj.objJOMain.dtp_ActEndDate = ds.Tables(0).Rows(0)("ActEndDate").ToString()
                Obj.objJOMain.dbl_ManDays = ds.Tables(0).Rows(0)("ManDays").ToString()
                Obj.objJOMain.dbl_EstCost = ds.Tables(0).Rows(0)("EstCost").ToString()
                Obj.objJOMain.dbl_ActCost = ds.Tables(0).Rows(0)("ActCost").ToString()
                Obj.objJOMain.dbl_EstMatCost = ds.Tables(0).Rows(0)("EstMatCost").ToString()
                Obj.objJOMain.dbl_ActMatCost = ds.Tables(0).Rows(0)("ActMatCost").ToString()
                Obj.objJOMain.str_Status = ds.Tables(0).Rows(0)("Status").ToString()
                Obj.objJOMain.str_LpoNo = ds.Tables(0).Rows(0)("LpoNo").ToString()

                Obj.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
                Obj.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
                Obj.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
                Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
                Obj.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
                Obj.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
                Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus").ToString()

                If ds.Tables(1).Rows.Count > 0 Then
                    Obj.objJOSub.DT_JOItemDetailsFG = ds.Tables(1)
                End If

                If ds.Tables(2).Rows.Count > 0 Then
                    Obj.objJOVarBOM.DT_JOItemDetailsRM = ds.Tables(2)
                End If

                If ds.Tables(3).Rows.Count > 0 Then
                    Obj.objJOVarBOM.DT_BOMParam = ds.Tables(3)
                End If

                If ds.Tables(4).Rows.Count > 0 Then
                    Obj.objproject.str_ProjectID = ds.Tables(4).Rows(0)("ProjectID").ToString()
                    Obj.objproject.str_ProjectLocation = ds.Tables(4).Rows(0)("ProjectLocation").ToString()
                    Obj.objproject.str_WorkOrderNo = ds.Tables(4).Rows(0)("WorkOrderNo").ToString()
                End If
            Else
                Obj.objJOMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()
                Obj.objJOMain.dtp_JODate = Date.Now
                Obj.objJOMain.int_RevNo = 0
                Obj.objJOMain.str_SONo = Obj.objJOMain.str_JONo
                Obj.objJOMain.str_JODesc = ds.Tables(0).Rows(0)("Comment").ToString()
                Obj.objJOMain.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
                Obj.objJOMain.int_SalesManID = ds.Tables(0).Rows(0)("SalesManID").ToString()
                Obj.objJOMain.str_ProdUnitName = ""
                Obj.objJOMain.str_LpoNo = ds.Tables(0).Rows(0)("MerchantRef").ToString()

                Obj.objJOMain.dtp_EstEndDate = Date.Now
                Obj.objJOMain.dtp_ActEndDate = Date.Now
                Obj.objJOMain.dbl_ManDays = 0
                Obj.objJOMain.dbl_EstCost = 0
                Obj.objJOMain.dbl_ActCost = 0
                Obj.objJOMain.dbl_EstMatCost = 0
                Obj.objJOMain.dbl_ActMatCost = 0

                If ds.Tables(1).Rows.Count > 0 Then
                    Obj.objJOSub.DT_JOItemDetailsFG = ds.Tables(1)
                End If

                If ds.Tables(2).Rows.Count > 0 Then
                    Obj.objproject.str_ProjectID = ds.Tables(2).Rows(0)("ProjectID").ToString()
                    Obj.objproject.str_ProjectLocation = ds.Tables(2).Rows(0)("ProjectLocation").ToString()
                    Obj.objproject.str_WorkOrderNo = ds.Tables(2).Rows(0)("WorkOrderNo").ToString()
                End If
            End If

        Catch ex As Exception
            ErrNo = 1
            ErrMsg = ex.Message ' "Problem in Updating Invoice"
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub Get_StructureForJOProduction(ByVal _DBPath As String, ByVal _DBPwd As String, ByRef Obj As csJobOrder, ByRef ErrNo As Integer, ByRef ErrMsg As String)
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetJODetailsForProduction]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Obj.str_CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.objJOMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@JONo", Obj.objJOMain.str_JONo)

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            Obj.objJOMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()
            Obj.objJOMain.dtp_JODate = ds.Tables(0).Rows(0)("JODate").ToString()
            Obj.objJOMain.str_ProdStage = ds.Tables(0).Rows(0)("ProdStage").ToString()
            Obj.objJOMain.dtp_ProdDate = ds.Tables(0).Rows(0)("ProdDate").ToString()
            Obj.objJOMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.objJOMain.bit_UpdateInv = ds.Tables(0).Rows(0)("UpdateInv").ToString()
            Obj.objJOMain.str_Status = ds.Tables(0).Rows(0)("Status").ToString()

            Obj.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
            Obj.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
            Obj.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
            Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
            Obj.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
            Obj.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
            Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus")

            If ds.Tables(1).Rows.Count > 0 Then
                Obj.objJOSub.DT_JOItemDetailsFGProd = ds.Tables(1)
            End If

        Catch ex As Exception
            ErrNo = 1
            ErrMsg = ex.Message ' "Problem in Updating Invoice"
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Update_JO(ByVal _strPath As String, ByVal _strPwd As String, ByVal obj As csJobOrder, ByRef JONo As String, ByRef intRevNo As Integer, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("UpdateJobOrder", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.objJOMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.objJOMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.objJOMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.objJOMain.str_FormPrefix)

            BaseConn.cmd.Parameters.AddWithValue("@JODate", obj.objJOMain.dtp_JODate)
            BaseConn.cmd.Parameters.AddWithValue("@JONo", obj.objJOMain.str_JONo)
            BaseConn.cmd.Parameters.AddWithValue("@SONo", obj.objJOMain.str_SONo)
            BaseConn.cmd.Parameters.AddWithValue("@JODesc", obj.objJOMain.str_JODesc)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.objJOMain.int_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManID", obj.objJOMain.int_SalesManID)
            BaseConn.cmd.Parameters.AddWithValue("@ProdUnitName", obj.objJOMain.str_ProdUnitName)
            BaseConn.cmd.Parameters.AddWithValue("@LpoNo", obj.objJOMain.str_LpoNo)

            BaseConn.cmd.Parameters.AddWithValue("@EstEndDate", obj.objJOMain.dtp_EstEndDate)
            BaseConn.cmd.Parameters.AddWithValue("@ActEndDate", obj.objJOMain.dtp_ActEndDate)
            BaseConn.cmd.Parameters.AddWithValue("@ManDays", obj.objJOMain.dbl_ManDays)
            BaseConn.cmd.Parameters.AddWithValue("@EstCost", obj.objJOMain.dbl_EstCost)
            BaseConn.cmd.Parameters.AddWithValue("@ActCost", obj.objJOMain.dbl_ActCost)
            BaseConn.cmd.Parameters.AddWithValue("@EstMatCost", obj.objJOMain.dbl_EstMatCost)
            BaseConn.cmd.Parameters.AddWithValue("@ActMatCost", obj.objJOMain.dbl_ActMatCost)
            BaseConn.cmd.Parameters.AddWithValue("@Status", obj.objJOMain.str_Status)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.bool_ApprovedStatus)

            BaseConn.cmd.Parameters.AddWithValue("@ProjectID", obj.objproject.str_ProjectID)
            BaseConn.cmd.Parameters.AddWithValue("@WorkOrderNo", obj.objproject.str_WorkOrderNo)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectLocation", obj.objproject.str_ProjectLocation)

            BaseConn.cmd.Parameters.AddWithValue("@JOItemDetailsDTFG", obj.objJOSub.DT_JOItemDetailsFG)
            BaseConn.cmd.Parameters.AddWithValue("@JOItemDetailsDTRM", obj.objJOVarBOM.DT_JOItemDetailsRM)
            BaseConn.cmd.Parameters.AddWithValue("@JOBOMParameterDT", obj.objJOVarBOM.DT_BOMParam)

            BaseConn.cmd.Parameters.Add("@JONoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            JONo = BaseConn.cmd.Parameters("@JONoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_CID)
            ObjDalGeneral.Elog_Insert(obj.str_CID, _strPath, _strPwd, obj.objJOMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "BOM", ErrNo, "Error in " & obj.objJOMain.str_Flag & " : " & obj.objJOMain.str_JONo & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_JO = _ErrString
    End Function

    Public Function Update_JOForProduction(ByVal _strPath As String, ByVal _strPwd As String, ByVal obj As csJobOrder, ByRef JONo As String, ByRef intRevNo As Integer, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("sp_JobOrderProdStatusUpdated", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", obj.str_CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.objJOMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.objJOMain.str_MenuID)

            BaseConn.cmd.Parameters.AddWithValue("@JONo", obj.objJOMain.str_JONo)
            BaseConn.cmd.Parameters.AddWithValue("@ProdStage", obj.objJOMain.str_ProdStage)
            BaseConn.cmd.Parameters.AddWithValue("@ProdDate", obj.objJOMain.dtp_ProdDate)
            BaseConn.cmd.Parameters.AddWithValue("@UpdateInv", obj.objJOMain.bit_UpdateInv)
            BaseConn.cmd.Parameters.AddWithValue("@Status", obj.objJOMain.str_Status)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.objJOMain.str_Comment)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.bool_ApprovedStatus)

            BaseConn.cmd.Parameters.AddWithValue("@JOItemsFGProduction", obj.objJOSub.DT_JOItemDetailsFGProd)

            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            'BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            'Dim ds As New DataSet
            'BaseConn.da.Fill(ds)

            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_CID)
            ObjDalGeneral.Elog_Insert(obj.str_CID, _strPath, _strPwd, obj.objJOMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "JOProd", ErrNo, "Error in " & obj.objJOMain.str_Flag & " : " & obj.objJOMain.str_JONo & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        Update_JOForProduction = _ErrString
    End Function
End Class
