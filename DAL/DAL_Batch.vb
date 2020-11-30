Imports Classes

Public Class DAL_Batch
    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Function Update_Batch(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal obj As csBatchMaster, ByRef VouNo As String, ByVal _Flag As String, _
                                 ByVal _BusinessPeriodID As Integer, ByRef intRevNo As Integer, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0


        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("UpdateBatch", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.str_FormPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@BatchID", obj.str_BatchID)
            BaseConn.cmd.Parameters.AddWithValue("@BatchDesc", obj.str_BatchDesc)
            BaseConn.cmd.Parameters.AddWithValue("@MfgDate", obj.str_MfgDate)
            BaseConn.cmd.Parameters.AddWithValue("@ExpDate", obj.str_ExpDate)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", obj.str_ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@VouRef", obj.str_VouRef)
            BaseConn.cmd.Parameters.AddWithValue("@InActive", obj.bool_InActive)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.bool_ApprovedStatus)

            BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            VouNo = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString

        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, _BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "Batch", _
             Err.Number, "Error in " & _Flag & " : " & obj.str_BatchID & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_Batch = _ErrString
    End Function

    Public Sub Get_Structure(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef Obj As csBatchMaster, ByRef ErrNo As Integer, _
                             ByRef ErrStr As String)

        ErrNo = 0
        ErrStr = ""

        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetBatchMasterDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BatchID", Obj.str_BatchID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            Obj.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
            Obj.str_BatchDesc = ds.Tables(0).Rows(0)("BatchDesc").ToString()
            Obj.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
            Obj.str_MfgDate = ds.Tables(0).Rows(0)("MfgDate").ToString()
            Obj.str_ExpDate = ds.Tables(0).Rows(0)("ExpDate").ToString()
            Obj.str_ItemCode = ds.Tables(0).Rows(0)("ItemCode").ToString()
            Obj.str_VouRef = ds.Tables(0).Rows(0)("VouRef").ToString()
            Obj.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
            Obj.bool_InActive = ds.Tables(0).Rows(0)("InActive").ToString()

            Obj.str_CreatedBy = ds.Tables(0).Rows(0)("CreatedBy").ToString()
            Obj.dtp_CreatedDate = ds.Tables(0).Rows(0)("CreatedDate").ToString()
            Obj.str_LastUpdatedBy = ds.Tables(0).Rows(0)("LastUpdatedBy").ToString()
            Obj.dtp_LastUpdatedDate = ds.Tables(0).Rows(0)("LastUpdatedDate").ToString()
            Obj.str_ApprovedBy = ds.Tables(0).Rows(0)("ApprovedBy").ToString()
            Obj.dtp_ApprovedDate = ds.Tables(0).Rows(0)("ApprovedDate").ToString()
            Obj.bool_ApprovedStatus = ds.Tables(0).Rows(0)("ApprovedStatus")

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub Get_AllBatches(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByRef _DTBatch As DataTable, ByRef ErrNo As Integer, _
                           ByRef ErrStr As String)

        ErrNo = 0
        ErrStr = ""

        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetBatchMasterDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            _DTBatch = ds.Tables(0)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetExpiredBatchDetails(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As Integer, ByVal _Flag As String, ByRef _DTBatch As DataTable, ByRef ErrNo As Integer,
                           ByRef ErrStr As String)

        ErrNo = 0
        ErrStr = ""

        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetExpiredBatchDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            _DTBatch = ds.Tables(0)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub BatchInActive(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As Integer, ByVal _DTBatch As DataTable)
        BaseConn.Open(_StrDBPath, _StrDBPwd)
        BaseConn.cmd = New SqlClient.SqlCommand("BatchInactive", BaseConn.cnn)
        BaseConn.cmd.CommandType = CommandType.StoredProcedure
        BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID) 'obj.str_SiteID
        BaseConn.cmd.Parameters.AddWithValue("@DTBatch", _DTBatch)
        BaseConn.cmd.ExecuteNonQuery()
    End Sub
    Public Function GetBatchReport(ByVal _strPath As String, ByVal _strPwd As String, ByVal _SiteID As Integer, ByVal ItemCodeColl As DataTable,
                                   ByVal BatchID As String, ByVal Condition As String, ByVal DateType As String, ByVal Date1 As String, ByVal FromDate As Date,
                                   ByVal ToDate As Date, ByVal Flag As String, Optional ByVal BinID As String = "", Optional ByVal WH As String = "") As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetBatchReport]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemArray", ItemCodeColl)
            BaseConn.cmd.Parameters.AddWithValue("@BatchID", BatchID)
            BaseConn.cmd.Parameters.AddWithValue("@BinID", BinID)
            BaseConn.cmd.Parameters.AddWithValue("@WHID", WH)
            BaseConn.cmd.Parameters.AddWithValue("@DateType", DateType)
            BaseConn.cmd.Parameters.AddWithValue("@Date1", Date1)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", Condition)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Flag)
            'BaseConn.cmd.Parameters.AddWithValue("@VouRef", VouRef)
            'BaseConn.cmd.Parameters.AddWithValue("@COO", COO)
            'BaseConn.cmd.Parameters.AddWithValue("@ZeroSuppress", _ZeroSuppress)
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = ds.Tables(0)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Sub BatchCreateLPOItems(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _VouNo As String, ByVal _ConvertFrom As String, ByVal _DTItemBatch As DataTable, ByVal _Alias As String, ByVal _LedgerID As Integer, ByVal _UserName As String, ByRef _ErrStr As String, ByRef _ErrNo As Integer)
        _ErrNo = 0
        _ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("BatchCreateLPOItems", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", _VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", _UserName)
            BaseConn.cmd.Parameters.AddWithValue("@ConvertFrom", _ConvertFrom)
            BaseConn.cmd.Parameters.AddWithValue("@DTItemBatch", _DTItemBatch)
            BaseConn.cmd.Parameters.AddWithValue("@Alias", _Alias)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            _ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrStr = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.ToString
        Finally
            BaseConn.Close()
        End Try

    End Sub

    Public Function GetBatchProfit(ByVal str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal ItemCode As String, ByVal BatchID As String, ByVal VouNo As String, ByVal BinID As String, ByVal Serial As String, ByVal BatchRefFrom As String, ByVal BatchRefTo As String, ByVal Flag As String, Optional ByRef dt_PurExpense As DataTable = Nothing, Optional ByRef dt_SalesExpense As DataTable = Nothing, Optional ByRef dt_JV As DataTable = Nothing) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetBatchProfit]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@BatchID", BatchID)
            BaseConn.cmd.Parameters.AddWithValue("@BinID", BinID)
            BaseConn.cmd.Parameters.AddWithValue("@Serial", Serial)
            BaseConn.cmd.Parameters.AddWithValue("@BatchRefFrom", BatchRefFrom)
            BaseConn.cmd.Parameters.AddWithValue("@BatchRefTo", BatchRefTo)
            'BaseConn.cmd.Parameters.AddWithValue("@Date1", Date1)
            'BaseConn.cmd.Parameters.AddWithValue("@FromDate", FromDate)
            'BaseConn.cmd.Parameters.AddWithValue("@ToDate", ToDate)
            'BaseConn.cmd.Parameters.AddWithValue("@Condition", Condition)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Flag)
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            If Flag = "EXPORT" Then
                dt = ds.Tables(0)
                dt_PurExpense = ds.Tables(1)
                dt_SalesExpense = ds.Tables(2)
                dt_JV = ds.Tables(3)
            Else
                dt = ds.Tables(0)
            End If

        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
    Public Function GetBatchPriceList(ByVal str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal _Condition As String, ByVal _ItemCondition As String) As DataTable
        GetBatchPriceList = New DataTable
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetBatchPriceList]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", _Condition)
            BaseConn.cmd.Parameters.AddWithValue("@ItemConditon", _ItemCondition)
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            GetBatchPriceList = ds.Tables(0)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return GetBatchPriceList
    End Function

    Public Function Update_BatchPrice(ByVal str_SiteID As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _DTBatchPrice As DataTable, ByVal _Flag As String) As String
        Dim _ErrString As String = ""

        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("sp_UpdateBatchPrice", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@DTBatchPrice", _DTBatchPrice)

            'BaseConn.cmd.Parameters.Add("@VouNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            'BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            'BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            'BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            'VouNo = BaseConn.cmd.Parameters("@VouNoOut").Value.ToString
            'intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            'ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            '_ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString

        Catch ex As Exception
            _ErrString = ex.Message
            'ObjDalGeneral = New DAL_General(obj.str_SiteID)
            'ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, _BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "Batch", _
            ' Err.Number, "Error in " & _Flag & " : " & obj.str_BatchID & "", ex.Message, 5, 3, 1, ErrNo)
            'ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_BatchPrice = _ErrString
    End Function
End Class
