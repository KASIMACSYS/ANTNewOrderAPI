'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Imports Classes

Public Class DAL_ProductGrouping
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()

    Public Sub Get_Structure(ByRef Obj As csProductGrouping, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[ProductGroupingUpdated]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPerionID)

            BaseConn.cmd.Parameters.AddWithValue("@ProductID", Obj.str_ProductID)
            BaseConn.cmd.Parameters.AddWithValue("@Description", Obj.str_Description)
            BaseConn.cmd.Parameters.AddWithValue("@ParentID", Obj.str_ParentID)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", Obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", Obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", Obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", Obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCodeDT", Obj.dt_ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.str_Flag)
            'BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            'BaseConn.da.Fill(dt)
            BaseConn.cmd.ExecuteNonQuery()

            'Obj.dt_Types = ds.Tables(0)
            'Obj.dt_Parameter = ds.Tables(1)
            'Obj.dt_Mccb = ds.Tables(2)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Sub Get_Structure1(ByRef Obj As csProductGrouping, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef _ErrNo As Integer, ByRef _ErrStr As String, Optional ByVal _WHID As String = "")
        _ErrNo = 0
        _ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetProductGrouping]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPerionID)
            BaseConn.cmd.Parameters.AddWithValue("@ProductID", Obj.str_ProductID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemColumn", Obj.str_ItemColumn)
            BaseConn.cmd.Parameters.AddWithValue("@PriceShow", Obj.bool_Price)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", Obj.Condition)
            BaseConn.cmd.Parameters.AddWithValue("@WHID", _WHID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.cmd.CommandTimeout = 500
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            Obj.dt_selecteditem = ds.Tables(0)
            If Obj.dt_selecteditem.Columns.Contains("Code") = True Then
                Obj.dt_selecteditem.Columns.Remove("Code")
                Obj.dt_selecteditem.AcceptChanges()
            End If
            'If ds.Tables.Count > 1 Then
            '    Obj.dt_selectedall = ds.Tables(1)
            'End If
            '
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub Get_Structure2(ByRef DT As DataTable, ByVal _FromDate As Date, ByVal _ToDate As Date, ByVal _SiteID As String, ByVal int_BusinessPerionID As Integer, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        _ErrNo = 0
        _ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetStockMovementExport]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", int_BusinessPerionID)
            BaseConn.cmd.Parameters.AddWithValue("@DT", DT)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.cmd.CommandTimeout = 500
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            DT = ds.Tables(0)
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message.ToString
        End Try
    End Sub
    Public Sub Get_GetStockForproductgrouping(ByRef DT As DataTable, ByVal _ProductID As String, ByVal _SiteID As String, ByVal int_BusinessPerionID As Integer, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef _ErrNo As Integer, ByRef _ErrStr As String, ByVal _Flag As String)
        _ErrNo = 0
        _ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetStockForProductGrouping]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", int_BusinessPerionID)
            BaseConn.cmd.Parameters.AddWithValue("@ProductID", _ProductID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.cmd.CommandTimeout = 500
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            DT = ds.Tables(0)
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message.ToString
        End Try
    End Sub
End Class

