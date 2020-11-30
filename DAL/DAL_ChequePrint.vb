Imports Classes

Public Class DAL_ChequePrint

    Private BaseConn As New SQLConn()
    Private dt As DataTable
    Private ObjDalGeneral As DAL_General

    Public Function Get_Structure(ByRef Obj As csChequePrint, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _SiteID As Integer, _
                                  ByVal _Flag As String, ByRef _ErrNo As Integer, ByRef _ErrString As String) As DataTable
        dt = New DataTable
        _ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetChequePrintDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", Obj.str_VouNo)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            Obj.str_ConvertNo = ds.Tables(0).Rows(0)("ConvertNo").ToString()
            Obj.int_BankLedgerID = ds.Tables(0).Rows(0)("BankLedgerID").ToString()
            Obj.int_MerchantLedgerID = ds.Tables(0).Rows(0)("MerchantLedgerID").ToString()

            Obj.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
            Obj.int_PrintRevNo = ds.Tables(0).Rows(0)("PrintRevNo").ToString()
            Obj.str_Alias = ds.Tables(0).Rows(0)("Alias").ToString()
            Obj.str_ChequeNumber = ds.Tables(0).Rows(0)("ChequeNo").ToString()
            Obj.dbl_Amount = ds.Tables(0).Rows(0)("Amount").ToString()
            Obj.dtp_VouDate = ds.Tables(0).Rows(0)("VouDate").ToString()
            Obj.dtp_ChequeDate = ds.Tables(0).Rows(0)("ChequeDate").ToString()
            Obj.bool_ACpayeeonly = ds.Tables(0).Rows(0)("ACPayeeOnly").ToString()
            Obj.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.str_RefID = ds.Tables(0).Rows(0)("RefID").ToString()
            Obj.bool_Cancelled = ds.Tables(0).Rows(0)("Cancelled").ToString()

        Catch ex As Exception
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Sub Put_Structure(ByRef Obj As csChequePrint, ByRef _VouNo As String, ByRef _RevNo As Integer, ByVal _SiteID As String, ByVal _DBPath As String, ByVal _DBPwd As String, ByRef _ErrNo As Integer, ByRef _ErrString As String)
        _ErrString = ""
        _ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("ChequePrintUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", Obj.str_FormPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPeriodID)

            BaseConn.cmd.Parameters.AddWithValue("@VouNo", Obj.str_VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@ConvertNo", Obj.str_ConvertNo)
            BaseConn.cmd.Parameters.AddWithValue("@BankLedgerID", Obj.int_BankLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@MerchantLedgerID", Obj.int_MerchantLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", Obj.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@PrintRevNo", Obj.int_PrintRevNo)

            BaseConn.cmd.Parameters.AddWithValue("@Alias", Obj.str_Alias)
            BaseConn.cmd.Parameters.AddWithValue("@ChequeNo", Obj.str_ChequeNumber)
            BaseConn.cmd.Parameters.AddWithValue("@Amount", Obj.dbl_Amount)

            BaseConn.cmd.Parameters.AddWithValue("@VouDate", Obj.dtp_VouDate)
            BaseConn.cmd.Parameters.AddWithValue("@ChequeDate", Obj.dtp_ChequeDate)
            BaseConn.cmd.Parameters.AddWithValue("@ACPayeeOnly", Obj.bool_ACpayeeonly)

            BaseConn.cmd.Parameters.AddWithValue("@Comment", Obj.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@RefID", Obj.str_RefID)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", Obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", Obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", Obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", Obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", Obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", Obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", Obj.bool_ApprovedStatus)
            BaseConn.cmd.Parameters.AddWithValue("@Cancelled", Obj.bool_Cancelled)

            BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output

            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output

            BaseConn.cmd.ExecuteNonQuery()

            _VouNo = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
            _RevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString

            _ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(Obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(Obj.str_SiteID, _DBPath, _DBPwd, Obj.int_BusinessPeriodID, Obj.str_CreatedBy, Obj.dtp_CreatedDate, "", "ChequePrint", Err.Number, "Error in '" & Obj.str_Flag & "'ED '" & Obj.str_VouNo & "' ", ex.Message, 5, 3, 1, 0)
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub

End Class
