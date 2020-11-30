Imports Classes
Public Class DAL_TaxFileReturn
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General
    'Public Sub Get_Structure(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal obj As csTaxFileReturn, ByRef ErrNo As Integer, ByRef ErrMsg As String)
    '    Try
    '        BaseConn.Open(_StrDBPath, _StrDBPwd)
    '        BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetTaxFileReturnDetails]", BaseConn.cnn)
    '        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    '        BaseConn.cmd.Parameters.AddWithValue("@SiteID", obj.str_SiteID) 'obj.str_SiteID
    '        BaseConn.cmd.Parameters.AddWithValue("@VouNo", obj.str_VouNo)
    '        BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.str_Flag)
    '        BaseConn.cmd.Parameters.AddWithValue("@Condition", obj.str_Condition)
    '        BaseConn.cmd.Parameters.AddWithValue("@FromDate", obj.dtp_FromDate)
    '        BaseConn.cmd.Parameters.AddWithValue("@ToDate", obj.dtp_ToDate)
    '        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    '        Dim ds As New DataSet
    '        BaseConn.da.Fill(ds)
    '        If obj.str_Flag = "ADD" Or (obj.str_Flag = "EDIT" And obj.str_Condition = "Open") Then
    '            obj.dt_TaxFileReturn = ds.Tables(0)
    '        Else
    '            obj.str_Description = ds.Tables(0).Rows(0)("Description").ToString()
    '            obj.str_TaxAgentUID = ds.Tables(0).Rows(0)("TaxAgentUID").ToString()
    '            obj.dtp_VouDate = ds.Tables(0).Rows(0)("VouDate")
    '            obj.dtp_FromDate = ds.Tables(0).Rows(0)("FromDate")
    '            obj.dtp_ToDate = ds.Tables(0).Rows(0)("ToDate")
    '            obj.str_Status = ds.Tables(0).Rows(0)("Status").ToString()
    '            obj.dt_TaxFileReturn = ds.Tables(1)
    '        End If
    '    Catch ex As Exception

    '    End Try


    'End Sub

    Public Sub GetInvoiceToFileTax(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _CID As String, ByVal _FromDate As Date,
                                        ByVal _ToDate As Date, ByRef _DTInvoiceForTax As DataTable)
        Try

            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetInvoiceToFileTax]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            _DTInvoiceForTax = ds.Tables(0)
        Catch ex As Exception

        End Try


    End Sub

    Public Sub Get_Structure(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _Flag As String, _
                                        ByVal _VouNo As String, ByRef _DSFileReturn As DataSet, ByRef _PayableAmount As Decimal, ByRef _ClaimableAmount As Decimal)
        Try

            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetTaxFileReturn]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", _VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.Add("@PayableAmount", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ClaimableAmount", SqlDbType.Decimal).Direction = ParameterDirection.Output

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _PayableAmount = BaseConn.cmd.Parameters("@PayableAmount").Value.ToString
            _ClaimableAmount = BaseConn.cmd.Parameters("@ClaimableAmount").Value.ToString
            _DSFileReturn = ds
        Catch ex As Exception

        End Try


    End Sub
    Public Function Put_Structure(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal obj As csTaxFileReturn, ByRef VouNo As String, ByRef intRevNo As Integer, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0

        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("UpdateTaxFileReturn", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_CID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", obj.str_VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.Str_FormPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@Description", obj.str_Description)
            BaseConn.cmd.Parameters.AddWithValue("@TaxAgentUID", obj.str_TaxAgentUID)
            BaseConn.cmd.Parameters.AddWithValue("@VouDate", obj.dtp_VouDate)

            BaseConn.cmd.Parameters.AddWithValue("@FromDate", obj.dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", obj.dtp_ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Status", obj.str_Status)
            'BaseConn.cmd.Parameters.AddWithValue("@DT_TaxFileReturn", obj.dt_TaxFileReturn)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.bool_ApprovedStatus)
            'BaseConn.cmd.Parameters.AddWithValue("@VouDate", obj.dtp_VouDate)
            'BaseConn.cmd.Parameters.AddWithValue("@VouDate", obj.dtp_VouDate)

            BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output

            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output

            BaseConn.cmd.ExecuteNonQuery()

            VouNo = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString

        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_CID)
            ObjDalGeneral.Elog_Insert(obj.str_CID, _StrDBPath, _StrDBPwd, 0, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "TaxFileReturn", Err.Number, "Error in " & obj.str_Flag & " : " & obj.str_VouNo & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
        Put_Structure = _ErrString
    End Function
    Public Sub Get_TaxFileReturnMain(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal obj As csTaxFileReturn, ByRef ErrNo As Integer, ByRef ErrMsg As String)
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetTaxFileReturnMain]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", obj.str_CID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.int_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@All", obj.bool_All)
            BaseConn.cmd.Parameters.AddWithValue("@Open", obj.bool_Open)
            BaseConn.cmd.Parameters.AddWithValue("@Submitted", obj.bool_Submitted)
            BaseConn.cmd.Parameters.AddWithValue("@DateType", obj.str_Condition)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", obj.dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", obj.dtp_ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@VouType", obj.str_VouType)
           
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            obj.dt_TaxFileReturn = ds.Tables(0)

        Catch ex As Exception
            ErrMsg = ex.ToString
        End Try
    End Sub

    Public Sub GetTaxFileRtnDate(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _CID As String, ByRef _FromDate As Date)
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetTaxFileRtnDate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.Add("@FromDate", SqlDbType.Date, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()

            _FromDate = BaseConn.cmd.Parameters("@FromDate").Value

        Catch ex As Exception
            'ErrMsg = ex.ToString
        End Try
    End Sub
    Public Sub Get_TaxFileReturnFromExcel(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal obj As csTaxFileReturn, ByRef ErrNo As Integer, ByRef ErrMsg As String)
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetTaxFileReturnFromExcel]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_CID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.str_Flag)
            'BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.int_LedgerID)
            'BaseConn.cmd.Parameters.AddWithValue("@All", obj.bool_All)
            'BaseConn.cmd.Parameters.AddWithValue("@Open", obj.bool_Open)
            'BaseConn.cmd.Parameters.AddWithValue("@Submitted", obj.bool_Submitted)
            'BaseConn.cmd.Parameters.AddWithValue("@DateType", obj.str_Condition)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", obj.dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", obj.dtp_ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@VouType", obj.str_VouType)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", obj.str_Condition)

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.cmd.CommandTimeout = 2000

            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            obj.dt_TaxFileReturn = ds.Tables(0)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub
End Class
