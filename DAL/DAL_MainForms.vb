'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Imports Classes
Public Class DAL_MainForms
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()

    Public Sub Get_Structure(ByRef Obj As csMainForms, ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_strDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetMainForms]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.str_BusinessPerionID)
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@MerchantID", Obj.str_MerchantName)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManID", Obj.str_SalesManID)
            BaseConn.cmd.Parameters.AddWithValue("@Project", Obj.str_Project)
            BaseConn.cmd.Parameters.AddWithValue("@CurrencyCode", Obj.str_CurrencyCode)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", Obj.dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", Obj.dtp_ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Date1", Obj.dtp_Date)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", Obj.MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Partial", Obj.bool_Partial)
            BaseConn.cmd.Parameters.AddWithValue("@Open", Obj.bool_Open)
            BaseConn.cmd.Parameters.AddWithValue("@Manually", Obj.bool_Manually)
            BaseConn.cmd.Parameters.AddWithValue("@Closed", Obj.bool_Closed)
            BaseConn.cmd.Parameters.AddWithValue("@Draft", Obj.bool_Draft)
            BaseConn.cmd.Parameters.AddWithValue("@Status", Obj.bool_Status)
            BaseConn.cmd.Parameters.AddWithValue("@Invoiced", Obj.bool_Invoiced)
            BaseConn.cmd.Parameters.AddWithValue("@Uninvoiced", Obj.bool_Uninvoiced)
            BaseConn.cmd.Parameters.AddWithValue("@Cancelled", Obj.bool_Cancelled)
            BaseConn.cmd.Parameters.AddWithValue("@NotPaid", Obj.bool_NotPaid)
            BaseConn.cmd.Parameters.AddWithValue("@Paid", Obj.bool_Paid)
            BaseConn.cmd.Parameters.AddWithValue("@PartiallyPaid", Obj.bool_PartiallyPaid)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", Obj.str_ApprovedStatus)
            BaseConn.cmd.Parameters.AddWithValue("@AccountingPeriod", Obj.str_AccountingPeriod)
            BaseConn.cmd.Parameters.AddWithValue("@AccountingPeriodFrom", Obj.int_AccountingPeriodFrom)
            ''BaseConn.cmd.Parameters.AddWithValue("@AccountingPeriodTo", Obj.int_AccountingPeriodTo)
            BaseConn.cmd.Parameters.AddWithValue("@WorkOrderNo", Obj.str_WorkOrderNo)
            BaseConn.cmd.Parameters.AddWithValue("@WHMaster", Obj.str_WHMaster)
            BaseConn.cmd.Parameters.AddWithValue("@SignatureType", Obj.str_SignatureType)
            BaseConn.cmd.Parameters.AddWithValue("@DateType", Obj.str_DateType)
            BaseConn.cmd.Parameters.AddWithValue("@User", Obj.str_User)
            BaseConn.cmd.Parameters.AddWithValue("@GrpID", Obj.int_GrpID)
            BaseConn.cmd.Parameters.AddWithValue("@GrpName", Obj.str_GrpName)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", Obj.str_User)
            BaseConn.cmd.Parameters.AddWithValue("@Options", Obj.str_Options)
            BaseConn.cmd.Parameters.AddWithValue("@Options2", Obj.bool_Options2)
            BaseConn.cmd.Parameters.AddWithValue("@Options3", Obj.bool_Options3)
            BaseConn.cmd.Parameters.AddWithValue("@Filter1", Obj.str_Filter1)
            BaseConn.cmd.Parameters.AddWithValue("@Filter2", Obj.str_Filter2)
            BaseConn.cmd.Parameters.AddWithValue("@QtyStatusAll", Obj.bool_QtyStatusAll)
            BaseConn.cmd.Parameters.AddWithValue("@QtyStatusOpen", Obj.bool_QtyStatusOpen)
            BaseConn.cmd.Parameters.AddWithValue("@QtyStatusPartial", Obj.bool_QtyStatusPartial)
            BaseConn.cmd.Parameters.AddWithValue("@QtyStatusClose", Obj.bool_QtyStatusClose)


            If Obj.MenuID <> "Menu_540" Then
                BaseConn.cmd.Parameters.AddWithValue("@LedgerDept", "")
            Else
                BaseConn.cmd.Parameters.AddWithValue("@LedgerDept", "")
            End If
            BaseConn.cmd.CommandTimeout = 1000
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            dt = New DataTable
            BaseConn.da.Fill(dt)
            BaseConn.da.Dispose()
            Obj.dt_Main = dt.Copy
            dt.Dispose()
            dt = Nothing
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub Get_BI_Sales(ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByVal Obj As csMainForms, ByVal _MerchantDT As DataTable, _
                            ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_strDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_BI_Sales]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.str_BusinessPerionID)
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Obj.str_SiteID)
            If _MerchantDT.Rows.Count > 0 AndAlso _MerchantDT.Rows(0)("Name").ToString <> "--ALL--" Then
                Obj.str_MerchantName = 1
            End If
            BaseConn.cmd.Parameters.AddWithValue("@MerchantID", Obj.str_MerchantName)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManID", Obj.str_SalesManID)
            BaseConn.cmd.Parameters.AddWithValue("@Project", Obj.str_Project)
            BaseConn.cmd.Parameters.AddWithValue("@CurrencyCode", Obj.str_CurrencyCode)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", Obj.dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", Obj.dtp_ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Date1", Obj.dtp_Date)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", Obj.MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Partial", Obj.bool_Partial)
            BaseConn.cmd.Parameters.AddWithValue("@Open", Obj.bool_Open)
            BaseConn.cmd.Parameters.AddWithValue("@Manually", Obj.bool_Manually)
            BaseConn.cmd.Parameters.AddWithValue("@Closed", Obj.bool_Closed)
            BaseConn.cmd.Parameters.AddWithValue("@Status", Obj.bool_Status)
            BaseConn.cmd.Parameters.AddWithValue("@Invoiced", Obj.bool_Invoiced)
            BaseConn.cmd.Parameters.AddWithValue("@Uninvoiced", Obj.bool_Uninvoiced)
            BaseConn.cmd.Parameters.AddWithValue("@Cancelled", Obj.bool_Cancelled)
            BaseConn.cmd.Parameters.AddWithValue("@NotPaid", Obj.bool_NotPaid)
            BaseConn.cmd.Parameters.AddWithValue("@Paid", Obj.bool_Paid)
            BaseConn.cmd.Parameters.AddWithValue("@PartiallyPaid", Obj.bool_PartiallyPaid)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", Obj.str_ApprovedStatus)
            BaseConn.cmd.Parameters.AddWithValue("@AccountingPeriod", Obj.str_AccountingPeriod)
            BaseConn.cmd.Parameters.AddWithValue("@AccountingPeriodFrom", Obj.int_AccountingPeriodFrom)
            ''BaseConn.cmd.Parameters.AddWithValue("@AccountingPeriodTo", Obj.int_AccountingPeriodTo)
            BaseConn.cmd.Parameters.AddWithValue("@WorkOrderNo", Obj.str_WorkOrderNo)
            BaseConn.cmd.Parameters.AddWithValue("@WHMaster", Obj.str_WHMaster)
            BaseConn.cmd.Parameters.AddWithValue("@SignatureType", Obj.str_SignatureType)
            BaseConn.cmd.Parameters.AddWithValue("@DateType", Obj.str_DateType)
            BaseConn.cmd.Parameters.AddWithValue("@User", Obj.str_User)
            BaseConn.cmd.Parameters.AddWithValue("@GrpID", Obj.int_GrpID)
            BaseConn.cmd.Parameters.AddWithValue("@GrpName", Obj.str_GrpName)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", Obj.str_User)
            BaseConn.cmd.Parameters.AddWithValue("@Options", Obj.str_Options)
            BaseConn.cmd.Parameters.AddWithValue("@Options2", Obj.bool_Options2)
            BaseConn.cmd.Parameters.AddWithValue("@Options3", Obj.bool_Options3)
            BaseConn.cmd.Parameters.AddWithValue("@DTMerchantLedger", _MerchantDT)

            If Obj.MenuID <> "Menu_540" Then
                BaseConn.cmd.Parameters.AddWithValue("@LedgerDept", "")
            Else
                BaseConn.cmd.Parameters.AddWithValue("@LedgerDept", "")
            End If
            BaseConn.cmd.CommandTimeout = 1000
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            dt = New DataTable
            BaseConn.da.Fill(dt)
            Obj.dt_Main = dt
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Sub Get_HealthCheck(ByRef Obj As csMainForms, ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_strDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[HealthCheck]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.str_BusinessPerionID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.str_Options)
            BaseConn.cmd.Parameters.AddWithValue("@StartDate", Obj.dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@EndDate", Obj.dtp_ToDate)
            BaseConn.cmd.CommandTimeout = 1000
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            dt = New DataTable
            BaseConn.da.Fill(dt)
            Obj.dt_Main = dt
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub UpdatePOSConfig(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, _CID As String, TagID As String, Value As String)
        Dim objDalGeneral As New DAL_General(_CID)

        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("UpdatePOSConfig", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@TagID", TagID)
            BaseConn.cmd.Parameters.AddWithValue("@Value", Value)

            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            'ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
        Catch ex As Exception
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetPOSConfig(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, _CID As String, TagID As String, ByRef Value As String)
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetPOSConfig]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@TagID", TagID)
            BaseConn.cmd.CommandTimeout = 1000
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            dt = New DataTable
            BaseConn.da.Fill(dt)
            Value = dt.Rows(0)(0)
        Catch ex As Exception

        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub UpdateGridLayout(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, _CID As String, _MenuID As String, _Flag As String,
                                   _GridLayoutName As String, _MasterGridLayoutXML As String, _ChildGridLayoutXML As String, _DefaultLayout As Boolean,
                                   ByRef _UserRights As DataTable, ByRef _GroupRights As DataTable, _UpdatedBy As String, _UpdatedDate As Date, ByRef ErrNo As Int16)
        Dim objDalGeneral As New DAL_General(_CID)

        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("GridLayoutUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", _MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@GridLayoutName", _GridLayoutName)
            BaseConn.cmd.Parameters.AddWithValue("@MasterGridLayout", _MasterGridLayoutXML)
            BaseConn.cmd.Parameters.AddWithValue("@ChildGridLayout", _ChildGridLayoutXML)
            BaseConn.cmd.Parameters.AddWithValue("@DefaultLayout", _DefaultLayout)
            BaseConn.cmd.Parameters.AddWithValue("@UserRights", objDalGeneral.DatatableToJSONString(_UserRights))
            BaseConn.cmd.Parameters.AddWithValue("@GroupRights", objDalGeneral.DatatableToJSONString(_GroupRights))
            BaseConn.cmd.Parameters.AddWithValue("@UpdatedBy", _UpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@UpdatedDate", _UpdatedDate)

            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            'ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
        Catch ex As Exception
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub UpdateFilterString(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, _CID As String, _MenuID As String, _Flag As String,
                                     _FilterName As String, _FilterString As String, _DefaultFilter As Boolean,
                                    ByRef _UserRights As DataTable, ByRef _GroupRights As DataTable, _UpdatedBy As String, _UpdatedDate As Date, ByRef ErrNo As Int16)

        Dim objDalGeneral As New DAL_General(_CID)
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("MainFormFilterUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", _MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@FilterName", _FilterName)
            BaseConn.cmd.Parameters.AddWithValue("@FilterString", _FilterString)
            BaseConn.cmd.Parameters.AddWithValue("@DefaultFilter", _DefaultFilter)
            BaseConn.cmd.Parameters.AddWithValue("@UserRights", objDalGeneral.DatatableToJSONString(_UserRights))
            BaseConn.cmd.Parameters.AddWithValue("@GroupRights", objDalGeneral.DatatableToJSONString(_GroupRights))
            BaseConn.cmd.Parameters.AddWithValue("@UpdatedBy", _UpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@UpdatedDate", _UpdatedDate)
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            'ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
        Catch ex As Exception
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
    End Sub


End Class

