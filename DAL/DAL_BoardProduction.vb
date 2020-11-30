Imports Classes

Public Class DAL_BoardProduction

    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Structure(ByRef obj As csBoardProduction, ByVal _StrDBPath As String, ByVal str_SiteID As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String, ByVal _Flag As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GGBC_GetBoardProduction]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", obj.ObjProdConsumption.Str_VoucherID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.ObjProdConsumption.Str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjProdConsumption.Str_Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            obj.ObjProdConsumption.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
            obj.ObjProdConsumption.dtp_VouDate = ds.Tables(0).Rows(0)("VouDate").ToString()
            obj.ObjProdConsumption.Str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()

            obj.ObjProdConsumption.XMLData1 = ds.Tables(0).Rows(0)("XmlData1").ToString()
            obj.ObjProdConsumption.XMLData2 = ds.Tables(0).Rows(0)("XmlData2").ToString()
            obj.ObjProdConsumption.XMLData3 = ds.Tables(0).Rows(0)("XmlData3").ToString()

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Update_BoardProduction(ByVal obj As csBoardProduction, ByRef VouNo As String, ByRef intRevNo As Integer, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("GGBC_UpdateBoardProduction", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjProdConsumption.Str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.ObjProdConsumption.Str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.ObjProdConsumption.Str_FormPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", obj.ObjProdConsumption.Str_VoucherID)
            BaseConn.cmd.Parameters.AddWithValue("@VouDate", obj.ObjProdConsumption.dtp_VouDate)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.ObjProdConsumption.Str_Comment)

            BaseConn.cmd.Parameters.AddWithValue("@XMLDATA1", obj.ObjProdConsumption.XMLData1)
            BaseConn.cmd.Parameters.AddWithValue("@XMLDATA2", obj.ObjProdConsumption.XMLData2)
            BaseConn.cmd.Parameters.AddWithValue("@XMLDATA3", obj.ObjProdConsumption.XMLData3)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)

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
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "Budget", Err.Number, "Error in " & obj.ObjProdConsumption.Str_Flag & " : " & obj.ObjProdConsumption.Str_VoucherID & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_BoardProduction = _ErrString
    End Function


    Public Sub Get_ProdCons(ByRef obj As csBoardProduction_Ledger, ByVal _StrDBPath As String, ByVal str_SiteID As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String, ByVal _Flag As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_BoardProduction_Main]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", "")
            BaseConn.cmd.Parameters.AddWithValue("@Flag", "")

            BaseConn.cmd.Parameters.AddWithValue("@DateString", obj.str_Datestring)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", obj.dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", obj.dtp_ToDate)

            BaseConn.cmd.Parameters.AddWithValue("@All", obj.bool_All)
            BaseConn.cmd.Parameters.AddWithValue("@Board", obj.bool_Board)
            BaseConn.cmd.Parameters.AddWithValue("@Plaster", obj.bool_Plaster)
            BaseConn.cmd.Parameters.AddWithValue("@VAP", obj.bool_VAP)

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            obj._DTLedger = ds.Tables(0)
            'obj.ObjProdConsumption.int_RevNo = ds.Tables(0).Rows(0)("RevNo").ToString()
            'obj.ObjProdConsumption.dtp_VouDate = ds.Tables(0).Rows(0)("VouDate").ToString()
            'obj.ObjProdConsumption.Str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()

            'obj.ObjProdConsumption.XMLData1 = ds.Tables(0).Rows(0)("XmlData1").ToString()
            'obj.ObjProdConsumption.XMLData2 = ds.Tables(0).Rows(0)("XmlData2").ToString()
            'obj.ObjProdConsumption.XMLData3 = ds.Tables(0).Rows(0)("XmlData3").ToString()

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function GetDTforExport(ByVal str_SiteID As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _Flag As String, ByVal _DateString As String, ByVal _FromDate As Date, ByVal _ToDate As Date, ByRef ErrNo As Integer, ByRef ErrStr As String) As DataTable
        ErrNo = 0
        ErrStr = ""
        dt = New DataTable
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_BoardProduction_Report]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@DateString", _DateString)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = ds.Tables(0)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

End Class
