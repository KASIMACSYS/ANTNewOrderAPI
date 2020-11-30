Imports Classes
Public Class DAL_Estimation
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByVal _DBPath As String, ByVal _DBPwd As String, ByRef Obj As csEstimation, ByRef ErrNo As Integer, ByRef ErrMsg As String)
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetEstimationDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.objEstMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@EstNo", Obj.objEstMain.str_EstNo)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.objEstMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", Obj.objEstMain.str_MenuID)

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            Obj.objproject.str_ProjectID = ""
            Obj.objproject.str_ProjectLocation = ""
            Obj.objproject.str_WorkOrderNo = ""

            If Obj.objEstMain.str_Flag = "EST" Then
                Obj.objEstMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()
                Obj.objEstMain.dtp_EstDate = ds.Tables(0).Rows(0)("EstDate").ToString()
                Obj.objEstMain.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
                Obj.objEstMain.str_EstDesc = ds.Tables(0).Rows(0)("EstDesc").ToString()
                Obj.objEstMain.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
                Obj.objEstMain.int_SalesManID = ds.Tables(0).Rows(0)("SalesManID").ToString()
                Obj.objEstMain.str_MenuID = ds.Tables(0).Rows(0)("MenuID").ToString()
                Obj.objEstMain.str_InvRef = ds.Tables(0).Rows(0)("InvRef").ToString()

                'Obj.objEstMain.str_ProdUnitName = ds.Tables(0).Rows(0)("ProdUnitName").ToString()

                Obj.objEstMain.dtp_EstEndDate = ds.Tables(0).Rows(0)("EstEndDate").ToString()
                Obj.objEstMain.dtp_ActEndDate = ds.Tables(0).Rows(0)("ActEndDate").ToString()
                Obj.objEstMain.dbl_ManDays = ds.Tables(0).Rows(0)("ManDays").ToString()
                Obj.objEstMain.dbl_EstCost = ds.Tables(0).Rows(0)("EstCost").ToString()
                Obj.objEstMain.dbl_ActCost = ds.Tables(0).Rows(0)("ActCost").ToString()
                Obj.objEstMain.dbl_EstMatCost = ds.Tables(0).Rows(0)("EstMatCost").ToString()
                Obj.objEstMain.dbl_ActMatCost = ds.Tables(0).Rows(0)("ActMatCost").ToString()
                Obj.objEstMain.str_Status = ds.Tables(0).Rows(0)("Status").ToString()

                Obj.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
                Obj.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
                Obj.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
                Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
                Obj.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
                Obj.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
                Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus")

                If ds.Tables(1).Rows.Count > 0 Then
                    Obj.objEstSub.DT_EstItemDetailsFG = ds.Tables(1)
                End If

                'If ds.Tables(2).Rows.Count > 0 Then
                Obj.objEstVarBOM.DT_EstItemDetailsRM = ds.Tables(2)
                'End If

                'If ds.Tables(3).Rows.Count > 0 Then
                Obj.objEstVarBOM.DT_BOMParam = ds.Tables(3)
                'End If

                If ds.Tables(4).Rows.Count > 0 Then
                    Obj.objproject.str_ProjectID = ds.Tables(4).Rows(0)("ProjectID").ToString()
                    Obj.objproject.str_ProjectLocation = ds.Tables(4).Rows(0)("ProjectLocation").ToString()
                    Obj.objproject.str_WorkOrderNo = ds.Tables(4).Rows(0)("WorkOrderNo").ToString()
                End If
            ElseIf Obj.objEstMain.str_Flag = "POT" Then
                Obj.objEstSub.DT_EstItemDetailsFG = ds.Tables(0)
                Obj.objEstMain.dtp_EstDate = Date.Now
                Obj.objEstMain.dtp_EstEndDate = Date.Now
                Obj.objEstMain.dtp_ActEndDate = Date.Now
                Obj.objEstMain.dbl_ManDays = 0
                Obj.objEstMain.dbl_EstCost = 0
                Obj.objEstMain.dbl_ActCost = 0
                Obj.objEstMain.dbl_EstMatCost = 0
                Obj.objEstMain.dbl_ActMatCost = 0
                If ds.Tables(1).Rows.Count > 0 Then
                    Obj.objproject.str_ProjectID = ds.Tables(1).Rows(0)("ProjectID").ToString()
                    Obj.objproject.str_ProjectLocation = ds.Tables(1).Rows(0)("ProjectLocation").ToString()
                    Obj.objproject.str_WorkOrderNo = ds.Tables(1).Rows(0)("WorkOrderNo").ToString()
                End If
            Else
                Obj.objEstMain.int_BusinessPeriodID = ds.Tables(0).Rows(0)("BusinessPeriodID").ToString()
                Obj.objEstMain.dtp_EstDate = Date.Now
                Obj.objEstMain.int_RevNo = 0
                'Obj.objEstMain.str_SONo = Obj.objEstMain.str_JONo
                Obj.objEstMain.str_EstDesc = ds.Tables(0).Rows(0)("Comment").ToString()
                Obj.objEstMain.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
                Obj.objEstMain.int_SalesManID = ds.Tables(0).Rows(0)("SalesManID").ToString()
                Obj.objEstMain.str_ProdUnitName = ""

                Obj.objEstMain.dtp_EstEndDate = Date.Now
                Obj.objEstMain.dtp_ActEndDate = Date.Now
                Obj.objEstMain.dbl_ManDays = 0
                Obj.objEstMain.dbl_EstCost = 0
                Obj.objEstMain.dbl_ActCost = 0
                Obj.objEstMain.dbl_EstMatCost = 0
                Obj.objEstMain.dbl_ActMatCost = 0

                If ds.Tables(1).Rows.Count > 0 Then
                    Obj.objEstSub.DT_EstItemDetailsFG = ds.Tables(1)
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
    Public Function Update_Estimation(ByVal _strPath As String, ByVal _strPwd As String, ByVal obj As csEstimation, ByRef JONo As String, ByRef intRevNo As Integer, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("EstimationUpdated", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.objEstMain.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.objEstMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.objEstMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.objEstMain.str_FormPrefix)

            BaseConn.cmd.Parameters.AddWithValue("@EstDate", obj.objEstMain.dtp_EstDate)
            BaseConn.cmd.Parameters.AddWithValue("@EstNo", obj.objEstMain.str_EstNo)
            'BaseConn.cmd.Parameters.AddWithValue("@SONo", obj.objEstMain.str_SONo)
            BaseConn.cmd.Parameters.AddWithValue("@EstDesc", obj.objEstMain.str_EstDesc)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.objEstMain.int_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManID", obj.objEstMain.int_SalesManID)
            BaseConn.cmd.Parameters.AddWithValue("@InvRef", obj.objEstMain.str_InvRef)
            'BaseConn.cmd.Parameters.AddWithValue("@ProdUnitName", obj.objEstMain.str_ProdUnitName)

            BaseConn.cmd.Parameters.AddWithValue("@EstEndDate", obj.objEstMain.dtp_EstEndDate)
            BaseConn.cmd.Parameters.AddWithValue("@ActEndDate", obj.objEstMain.dtp_ActEndDate)
            BaseConn.cmd.Parameters.AddWithValue("@ManDays", obj.objEstMain.dbl_ManDays)
            BaseConn.cmd.Parameters.AddWithValue("@EstCost", obj.objEstMain.dbl_EstCost)
            BaseConn.cmd.Parameters.AddWithValue("@ActCost", obj.objEstMain.dbl_ActCost)
            BaseConn.cmd.Parameters.AddWithValue("@EstMatCost", obj.objEstMain.dbl_EstMatCost)
            BaseConn.cmd.Parameters.AddWithValue("@ActMatCost", obj.objEstMain.dbl_ActMatCost)
            BaseConn.cmd.Parameters.AddWithValue("@Status", obj.objEstMain.str_Status)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.bool_ApprovedStatus)

            BaseConn.cmd.Parameters.AddWithValue("@ApprovedLevel", obj.ApprovedLevel)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedComment", obj.ApprovedComment)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedHigherLevel", obj.ApprovedHigherLevel)

            BaseConn.cmd.Parameters.AddWithValue("@ProjectID", obj.objproject.str_ProjectID)
            BaseConn.cmd.Parameters.AddWithValue("@WorkOrderNo", obj.objproject.str_WorkOrderNo)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectLocation", obj.objproject.str_ProjectLocation)

            BaseConn.cmd.Parameters.AddWithValue("@ESTItemDetailsDTFG", obj.objEstSub.DT_EstItemDetailsFG)
            BaseConn.cmd.Parameters.AddWithValue("@ESTItemDetailsDTRM", obj.objEstVarBOM.DT_EstItemDetailsRM)
            BaseConn.cmd.Parameters.AddWithValue("@JOBOMParameterDT", obj.objEstVarBOM.DT_BOMParam)

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
            ObjDalGeneral.Elog_Insert(obj.str_CID, _strPath, _strPwd, obj.objEstMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "BOM", ErrNo, "Error in " & obj.objEstMain.str_Flag & " : " & obj.objEstMain.str_EstNo & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        Update_Estimation = _ErrString
    End Function
End Class
