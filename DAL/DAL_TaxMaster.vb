Imports Classes

Public Class DAL_TaxMaster
    Private ObjDalGeneral As DAL_General
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()


    Public Sub Get_Structure(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef Obj As csTaxMaster, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetTaxAgentDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.ObjTaxMasterMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@TaxID", Obj.ObjTaxMasterMain.int_TaxID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            Obj.ObjTaxMasterMain.str_TaxName = ds.Tables(0).Rows(0)("Name").ToString()
            Obj.ObjTaxMasterMain.str_TaxAgent = ds.Tables(0).Rows(0)("AgentTitle").ToString()
            Obj.ObjTaxMasterMain.str_TaxDesc = ds.Tables(0).Rows(0)("Description").ToString()
            Obj.ObjTaxMasterMain.str_SalesTax = ds.Tables(0).Rows(0)("SalesTaxAccount").ToString()
            Obj.ObjTaxMasterMain.str_PurchaseTax = ds.Tables(0).Rows(0)("PurchaseTaxAccount").ToString()
            Obj.ObjTaxMasterMain.bool_InActive = ds.Tables(0).Rows(0)("Status").ToString()
            Obj.ObjTaxMasterMain.str_TAN = ds.Tables(0).Rows(0)("TaxAgentNumber").ToString()
            Obj.ObjTaxMasterMain.str_TAAN = ds.Tables(0).Rows(0)("TaxAgentApprovalNumber").ToString()
            Obj.ObjTaxMasterMain.dtp_StartPeriodDate = ds.Tables(0).Rows(0)("PeriodStartDate")
            Obj.ObjTaxMasterMain.dtp_EndPeriodDate = ds.Tables(0).Rows(0)("PeriodEndDate")
            Obj.ObjTaxMasterMain.dtp_FAFCreationDate = ds.Tables(0).Rows(0)("FAFCreationDate")
            Obj.ObjTaxMasterMain.str_FAFVersion = ds.Tables(0).Rows(0)("FAFVersion").ToString()
            Obj.ObjTaxMasterMain.dt_TaxMaster = ds.Tables(1)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Sub GetTaxMasterDetailsALL(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByRef _DTTaxDetails As DataTable, ByRef ErrNo As Integer, ByRef ErrStr As String, Optional ByVal Flag As String = "")
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetTaxMasterDetailsALL]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure

            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _DTTaxDetails = ds.Tables(0)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Update_TaxAgent(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal obj As csTaxMaster, ByRef Int_TaxID As Integer, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("TaxAgentUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjTaxMasterMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@TaxID", obj.ObjTaxMasterMain.int_TaxID)
            BaseConn.cmd.Parameters.AddWithValue("@TaxName", obj.ObjTaxMasterMain.str_TaxName)
            BaseConn.cmd.Parameters.AddWithValue("@SalesTax", obj.ObjTaxMasterMain.str_SalesTax)
            BaseConn.cmd.Parameters.AddWithValue("@PurchaseTax", obj.ObjTaxMasterMain.str_PurchaseTax)
            BaseConn.cmd.Parameters.AddWithValue("@Description", obj.ObjTaxMasterMain.str_TaxDesc)
            BaseConn.cmd.Parameters.AddWithValue("@Agent", obj.ObjTaxMasterMain.str_TaxAgent)
            BaseConn.cmd.Parameters.AddWithValue("@Status", obj.ObjTaxMasterMain.bool_InActive)
            BaseConn.cmd.Parameters.AddWithValue("@TaxAgentNumber", obj.ObjTaxMasterMain.str_TAN)
            BaseConn.cmd.Parameters.AddWithValue("@TaxAgentApprovalNumber", obj.ObjTaxMasterMain.str_TAAN)
            BaseConn.cmd.Parameters.AddWithValue("@PeriodStartDate", obj.ObjTaxMasterMain.dtp_StartPeriodDate)
            BaseConn.cmd.Parameters.AddWithValue("@PeriodEndDate", obj.ObjTaxMasterMain.dtp_EndPeriodDate)
            BaseConn.cmd.Parameters.AddWithValue("@FAFCreationDate", obj.ObjTaxMasterMain.dtp_FAFCreationDate)
            BaseConn.cmd.Parameters.AddWithValue("@FAFVersion", obj.ObjTaxMasterMain.str_FAFVersion)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.ObjTaxMasterMain.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.ObjTaxMasterMain.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.ObjTaxMasterMain.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.ObjTaxMasterMain.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.ObjTaxMasterMain.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.ObjTaxMasterMain.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.ObjTaxMasterMain.bool_ApprovedStatus)
            BaseConn.cmd.Parameters.Add("@TaxIDOut", SqlDbType.Int, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            Int_TaxID = BaseConn.cmd.Parameters("@TaxIDOut").Value
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.int_BusinessPeriodID, obj.ObjTaxMasterMain.str_CreatedBy, obj.ObjTaxMasterMain.dtp_CreatedDate, "", "TaxMaster", Err.Number, "Error in " & obj.ObjTaxMasterMain.str_Flag & " : " & obj.ObjTaxMasterMain.int_TaxID & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        Update_TaxAgent = _ErrString
    End Function

    Public Sub Get_Structure(ByRef Obj As csTaxMaster, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetTaxConfigDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@TaxCode", Obj.ObjTaxMasterMain.str_TaxCode)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            Obj.ObjTaxMasterMain.int_TaxID = ds.Tables(0).Rows(0)("TaxID")
            Obj.ObjTaxMasterMain.str_TaxCode = ds.Tables(0).Rows(0)("TaxCode").ToString()
            Obj.ObjTaxMasterMain.str_TaxDesc = ds.Tables(0).Rows(0)("Description").ToString()
            Obj.ObjTaxMasterMain.dbl_SalesTaxPercentage = ds.Tables(0).Rows(0)("SalesPercentage").ToString()
            Obj.ObjTaxMasterMain.dbl_PurchaseTaxPercentage = ds.Tables(0).Rows(0)("PurchasePercentage").ToString()
            Obj.ObjTaxMasterMain.bool_InActive = ds.Tables(0).Rows(0)("Status").ToString()
            Obj.ObjTaxMasterMain.dbl_ReverseTax = ds.Tables(0).Rows(0)("ReverseTax").ToString()
            Obj.ObjTaxMasterMain.bool_PurchaseType = ds.Tables(0).Rows(0)("PurchaseType").ToString()
            'Obj.ObjTaxMasterMain.dt_TaxMaster = ds.Tables(1)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Update_TaxMaster(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal obj As csTaxMaster, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("TaxConfigUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjTaxMasterMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@TaxID", obj.ObjTaxMasterMain.int_TaxID)
            BaseConn.cmd.Parameters.AddWithValue("@TaxCode", obj.ObjTaxMasterMain.str_TaxCode)
            BaseConn.cmd.Parameters.AddWithValue("@PurchaseType", obj.ObjTaxMasterMain.bool_PurchaseType)
            BaseConn.cmd.Parameters.AddWithValue("@SalesTaxPercentage", obj.ObjTaxMasterMain.dbl_SalesTaxPercentage)
            BaseConn.cmd.Parameters.AddWithValue("@PurchaseTaxPercentage", obj.ObjTaxMasterMain.dbl_PurchaseTaxPercentage)
            BaseConn.cmd.Parameters.AddWithValue("@Description", obj.ObjTaxMasterMain.str_TaxDesc)
            BaseConn.cmd.Parameters.AddWithValue("@ReverseTax", obj.ObjTaxMasterMain.dbl_ReverseTax)
            BaseConn.cmd.Parameters.AddWithValue("@Status", obj.ObjTaxMasterMain.bool_InActive)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.ObjTaxMasterMain.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.ObjTaxMasterMain.dtp_CreatedDate)


            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            'Int_TaxID = BaseConn.cmd.Parameters("@TaxIDOut").Value
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.int_BusinessPeriodID, obj.ObjTaxMasterMain.str_CreatedBy, obj.ObjTaxMasterMain.dtp_CreatedDate, "", "TaxMaster", Err.Number, "Error in " & obj.ObjTaxMasterMain.str_Flag & " : " & obj.ObjTaxMasterMain.int_TaxID & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        Update_TaxMaster = _ErrString
    End Function
End Class
