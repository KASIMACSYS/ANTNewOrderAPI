'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Imports Classes
Imports System.Data.SqlClient
Imports Newtonsoft.Json
Imports System.Web.Script.Serialization

Public Class DAL_General
    Dim dt, dt1 As DataTable
    Dim BaseConn As New SQLConn()
    Dim CID As String

    Public Sub New(ByVal siteid As String)
        Me.CID = siteid
    End Sub

    Public Function Load_ComboDTWithAllData(ByVal _strPath As String, ByVal _strPwd As String, ByVal TableName As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_LoadComboDTWithAllData]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@TableName", TableName)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Sub test(ByVal _strPath As String, ByVal _strPwd As String)
        BaseConn.Open(_strPath, _strPwd)
        BaseConn.cmd = New SqlClient.SqlCommand("[sp_GETFTAVATAuditFile]", BaseConn.cnn)
        BaseConn.cmd.CommandType = CommandType.StoredProcedure
        BaseConn.cmd.Parameters.AddWithValue("@SiteID", 101)
        BaseConn.cmd.Parameters.AddWithValue("@FromDate", "2017-11-01")
        BaseConn.cmd.Parameters.AddWithValue("@ToDate", "2017-12-31")
        BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
        BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output

        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
        Dim ds As New DataSet
        BaseConn.da.Fill(ds)
        ds.WriteXml("Result")

        Dim ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
        Dim _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
    End Sub

    Public Function Load_ComboDT(ByVal _strPath As String, ByVal _strPwd As String, ByVal TableName As String, ByVal DisplayMember As String, ByVal ValueMember As String, ByVal Condition As String, ByVal Sorting As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[LoadComboDT]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@TableName", TableName)
            BaseConn.cmd.Parameters.AddWithValue("@DisplayMember", DisplayMember)
            BaseConn.cmd.Parameters.AddWithValue("@ValueMember", ValueMember)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", Condition)
            BaseConn.cmd.Parameters.AddWithValue("@Sorting", Sorting)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
            BaseConn.da.Dispose()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
    Public Function Load_ComboDTMultiLanguage(ByVal _strPath As String, ByVal _strPwd As String, ByVal TableName As String, ByVal DisplayMember As String, ByVal ValueMember As String, ByVal Condition As String, ByVal Sorting As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[LoadComboDTForMultiLanguage]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@TableName", TableName)
            BaseConn.cmd.Parameters.AddWithValue("@DisplayMember", DisplayMember)
            BaseConn.cmd.Parameters.AddWithValue("@ValueMember", ValueMember)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", Condition)
            BaseConn.cmd.Parameters.AddWithValue("@Sorting", Sorting)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
            BaseConn.da.Dispose()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function


    Public Function Load_ComBoPrefix(ByVal _strPath As String, ByVal _strPwd As String, ByVal TableName As String, ByVal DisplayMember As String, ByVal Condition As String, ByVal Sorting As String, ByVal PrefixLength As Integer) As DataTable
        Try
            'select substring(pip,5,len(pip)-4) from Pur_InvPosting 
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[LoadComboPrefix]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@TableName", TableName)
            BaseConn.cmd.Parameters.AddWithValue("@DisplayMember", DisplayMember)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", Condition)
            BaseConn.cmd.Parameters.AddWithValue("@Sorting", Sorting)
            BaseConn.cmd.Parameters.AddWithValue("@PrefixLength", PrefixLength)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
            BaseConn.da.Dispose()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
    Public Function Load_ItemMCCB(ByVal _strPath As String, ByVal _strpwd As String, ByVal _UserID As Integer, Optional ByVal _Condition As String = "") As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strpwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[LoadItemMccb]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@UserID", _UserID)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", _Condition)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.cmd.Dispose()
            BaseConn.da.Fill(dt)
            BaseConn.da.Dispose()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function Load_ItemMCCBByVoucherNumber(ByVal _strPath As String, ByVal _strpwd As String, ByVal _VouType As String, ByVal _VouNo As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strpwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[LoadItemMccbByVoucherNumber]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@VouType", _VouType)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", _VouNo)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
            BaseConn.da.Dispose()

        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function LoadBarcodeItems(ByVal _strPath As String, ByVal _strpwd As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strpwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_LoadBarcodeItems]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", CID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Sub LoadPOSItems(ByVal _strPath As String, ByVal _strpwd As String, ByVal _SiteID As String, ByRef _DTCategory As DataTable, ByRef _DTPOSItems As DataTable)
        Try
            Dim ds As New DataSet
            BaseConn.Open(_strPath, _strpwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_LoadPOSItems]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(ds)
            _DTCategory = ds.Tables(0)
            _DTPOSItems = ds.Tables(1)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function GetPOSCategoryAllItems(ByVal _strPath As String, ByVal _strpwd As String, ByVal _SiteID As String) As DataTable
        Try
            Dim ds As New DataSet
            BaseConn.Open(_strPath, _strpwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetPOSCategoryAllItems]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(ds)
            Return ds.Tables(0)

        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Function

    Public Sub GetItemForPOS(ByVal _strPath As String, ByVal _strpwd As String, ByVal _SiteID As String, ByVal _ItemCode As String,
                                ByRef _ItemDesc As String, ByRef _Unit As String, ByRef _Price As Decimal, ByRef _Tax As String)
        Try
            Dim ds As New DataSet
            BaseConn.Open(_strPath, _strpwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetItemForPOS]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", _ItemCode)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(ds)
            _ItemDesc = ds.Tables(0).Rows(0)("ItemDesc")
            _Unit = ds.Tables(0).Rows(0)("Unit")
            _Tax = ds.Tables(0).Rows(0)("Tax")
            _Price = ds.Tables(0).Rows(0)("Price")
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetPOSCategory(ByVal _strPath As String, ByVal _strpwd As String, ByVal _SiteID As String, ByRef _DTCategory As DataTable)
        Try
            Dim ds As New DataSet
            BaseConn.Open(_strPath, _strpwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetPOSCategory]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(ds)
            _DTCategory = ds.Tables(0)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    'Public Sub GetPOSCategoryItems(ByVal _strPath As String, ByVal _strpwd As String, ByVal _SiteID As String, ByVal _POSCategory As String, ByRef _DTCategoryItems As DataTable)
    '    Try
    '        Dim ds As New DataSet
    '        BaseConn.Open(_strPath, _strpwd)
    '        BaseConn.cmd = New SqlClient.SqlCommand("[GetPOSCategoryItems]", BaseConn.cnn)
    '        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    '        BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
    '        BaseConn.cmd.Parameters.AddWithValue("@POSCategory", _POSCategory)
    '        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    '        BaseConn.da.Fill(ds)
    '        _DTCategoryItems = ds.Tables(0)
    '    Catch ex As Exception
    '        MsgBox("Error" & ex.Message)
    '    Finally
    '        BaseConn.Close()
    '    End Try
    'End Sub
    Public Function LoadMCCBWithLedger(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _SiteID As String, ByVal _TableName As String,
                                       Optional ByVal _FormType As String = "", Optional ByVal _Condition As String = Nothing, Optional ByVal _MenuID As String = Nothing, Optional ByVal _WithInActive As Boolean = False) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[LoadMCCBWithLedger]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@TableName", _TableName.ToUpper)
            BaseConn.cmd.Parameters.AddWithValue("@Type", _FormType)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", _Condition)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", _MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@WithInActive", _WithInActive)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
            BaseConn.da.Dispose()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function Get_DataTable(ByVal _strPath As String, ByVal _strPwd As String, ByVal CID As Integer, ByVal TableName As String, ByVal FieldName As String, ByVal Condition As String, ByVal Sorting As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[getDataTable]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@TableName", TableName)
            BaseConn.cmd.Parameters.AddWithValue("@FieldName", FieldName)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", Condition)
            BaseConn.cmd.Parameters.AddWithValue("@Sorting", Sorting)
            BaseConn.cmd.CommandTimeout = 1000
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
            BaseConn.da.Dispose()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function getMainFormReport(ByVal _strPath As String, ByVal _strPwd As String, ByVal SiteID As String, ByVal TableName As String, ByVal Condition As DataTable) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetMainFormReport]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@TableName", TableName)
            'BaseConn.cmd.Parameters.AddWithValue("@ColumnName", ColumnName)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", Condition)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
    Public Function GetMultibleTable(ByVal SiteID As String, ByVal _strPath As String, ByVal _strpwd As String, ByVal Flag As String, ByVal Condition As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strpwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetMultipleTable]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Flag)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", Condition)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    'Public Function getDatatTable(ByVal TableName As String, ByVal FieldName As String, ByVal Condition As String, ByVal Sorting As String) As DataTable
    '    Try
    '        dt = New DataTable
    '        BaseConn.Open()
    '        BaseConn.cmd = New SqlClient.SqlCommand("[sp_getDataTable]", BaseConn.cnn)
    '        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    '        BaseConn.cmd.Parameters.AddWithValue("@SiteID", SiteID)
    '        BaseConn.cmd.Parameters.AddWithValue("@TableName", TableName)
    '        BaseConn.cmd.Parameters.AddWithValue("@FieldName", FieldName)
    '        BaseConn.cmd.Parameters.AddWithValue("@Condition", Condition)
    '        BaseConn.cmd.Parameters.AddWithValue("@Sorting", Sorting)
    '        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    '        BaseConn.da.Fill(dt)
    '    Catch ex As Exception
    '        MsgBox("Error" & ex.Message)
    '    Finally
    '        BaseConn.Close()
    '    End Try
    '    Return dt
    'End Function

    ''Public Function getBillwise(ByVal Str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal int_BusinessPeroidID As Integer, ByVal MerchantLedgerID As String, ByVal SalesManLedgerID As String, ByVal dtp_FromDate As Date, ByVal dtp_ToDate As Date) As DataTable
    ''    ''getBillwise = ObjDALGeneral.getBillwise(Str_SiteID, _strPath, _strPwd, int_BusinessPeroidID, MerchantLedgerID, SalesManLedgerID, dtp_FromDate, dtp_ToDate)
    ''    Try
    ''        dt = New DataTable
    ''        BaseConn.Open(_strPath, _strPwd)
    ''        BaseConn.cmd = New SqlClient.SqlCommand("[sp_getBillwiseReport]", BaseConn.cnn)
    ''        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    ''        BaseConn.cmd.Parameters.AddWithValue("@SiteID", Str_SiteID)
    ''        BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", int_BusinessPeroidID)
    ''        BaseConn.cmd.Parameters.AddWithValue("@MerchantLedgerID", MerchantLedgerID)
    ''        BaseConn.cmd.Parameters.AddWithValue("@SalesManID", SalesManLedgerID)
    ''        BaseConn.cmd.Parameters.AddWithValue("@FromDate", dtp_FromDate)
    ''        BaseConn.cmd.Parameters.AddWithValue("@ToDate", dtp_ToDate)
    ''        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    ''        Dim ds As New DataSet
    ''        BaseConn.da.Fill(ds)
    ''        dt = ds.Tables(0)
    ''    Catch ex As Exception
    ''        MsgBox("Error" & ex.Message)
    ''    Finally
    ''        BaseConn.Close()
    ''    End Try
    ''    Return dt
    ''End Function

    Public Function getBillwise_Old(ByVal Str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal int_BusinessPeroidID As Integer, ByVal MerchantLedgerID As String, ByVal SalesManLedgerID As String, ByVal dtp_FromDate As Date, ByVal dtp_ToDate As Date, ByVal NotPaid As Boolean, ByVal Paid As Boolean, ByVal PartialPaid As Boolean) As DataTable
        ''getBillwise = ObjDALGeneral.getBillwise(Str_SiteID, _strPath, _strPwd, int_BusinessPeroidID, MerchantLedgerID, SalesManLedgerID, dtp_FromDate, dtp_ToDate)
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_getBillwiseReport]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", int_BusinessPeroidID)
            BaseConn.cmd.Parameters.AddWithValue("@MerchantLedgerID", MerchantLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManID", SalesManLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", dtp_ToDate)

            BaseConn.cmd.Parameters.AddWithValue("@NotPaid", NotPaid)
            BaseConn.cmd.Parameters.AddWithValue("@FullyPaid", Paid)
            BaseConn.cmd.Parameters.AddWithValue("@PartialPaid", PartialPaid)
            BaseConn.cmd.CommandTimeout = 300

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

    Public Function getBillwise(ByVal Str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal int_BusinessPeroidID As Integer, ByVal MerchantLedgerDT As DataTable, ByVal SalesManLedgerID As String, ByVal dtp_FromDate As Date, ByVal dtp_ToDate As Date, ByVal NotPaid As Boolean, ByVal Paid As Boolean, ByVal PartialPaid As Boolean) As DataTable
        ''getBillwise = ObjDALGeneral.getBillwise(Str_SiteID, _strPath, _strPwd, int_BusinessPeroidID, MerchantLedgerID, SalesManLedgerID, dtp_FromDate, dtp_ToDate)
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            'BaseConn.cmd = New SqlClient.SqlCommand("[sp_getBillwiseReport]", BaseConn.cnn)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetBillwiseReport]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Str_SiteID)
            'BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", int_BusinessPeroidID)--TO DO  Enable 
            BaseConn.cmd.Parameters.AddWithValue("@MerchantLedgerDT", MerchantLedgerDT.DefaultView.ToTable(False, "LedgerID", "Name"))
            'BaseConn.cmd.Parameters.AddWithValue("@SalesManID", SalesManLedgerID)--TO DO  Enable 
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", dtp_ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@NotPaid", NotPaid)
            BaseConn.cmd.Parameters.AddWithValue("@FullyPaid", Paid)
            BaseConn.cmd.Parameters.AddWithValue("@PartialPaid", PartialPaid)
            BaseConn.cmd.CommandTimeout = 1000
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

    Public Function getBillwisePurchase(ByVal _strPath As String, ByVal _strPwd As String, ByVal Str_SiteID As String, ByVal MerchantLedgerDT As DataTable, ByVal dtp_FromDate As Date,
                                        ByVal dtp_ToDate As Date, ByVal NotPaid As Boolean, ByVal Paid As Boolean, ByVal PartialPaid As Boolean) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetBillwisePurchase]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@MerchantLedgerDT", MerchantLedgerDT.DefaultView.ToTable(False, "LedgerID", "Name"))
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", dtp_ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@NotPaid", NotPaid)
            BaseConn.cmd.Parameters.AddWithValue("@FullyPaid", Paid)
            BaseConn.cmd.Parameters.AddWithValue("@PartialPaid", PartialPaid)
            BaseConn.cmd.CommandTimeout = 1000
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

    Public Function GetAgingDetails(ByRef _Path As String, ByRef _Pwd As String, ByRef _CID As Integer, ByRef _LedgerID As DataTable, ByRef _ActiveOnly As Boolean,
                                     ByRef _FromDate As Date, ByRef _ToDate As Date, ByRef _IncludePDC As Boolean, ByRef _Aging As Boolean, ByVal _Summary As Boolean,
                                     ByRef _Project As String, ByRef _FormName As String, ByRef _DTAgingSlot As DataTable) As DataTable

        Try
            dt = New DataTable
            BaseConn.Open(_Path, _Pwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetAgingDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@ActiveOnly", _ActiveOnly)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@IncludePDC", _IncludePDC)
            BaseConn.cmd.Parameters.AddWithValue("@Aging", _Aging)
            BaseConn.cmd.Parameters.AddWithValue("@Summary", _Summary)
            BaseConn.cmd.Parameters.AddWithValue("@Project", _Project)
            BaseConn.cmd.Parameters.AddWithValue("@FormName", _FormName)
            BaseConn.cmd.Parameters.AddWithValue("@AgingSlotDT", _DTAgingSlot)
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

    Public Function getStockReport(ByVal str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal int_BusinessPeroidID As Integer,
                                   ByVal ItemCode As String, ByVal dtp_FromDate As Date, ByVal dtp_ToDate As Date, ByRef dbl_Stock As Double, ByRef dbl_WHStock As Double,
                                   ByRef dbl_Cost As Double, ByRef CostType As String, ByVal strRPTType As String, ByVal bool_UpdateCost As Boolean,
                                   Optional ByVal ItemCodeColl As DataTable = Nothing, Optional ByVal _WHID As Integer = 0, Optional ByVal LedgerID As Integer = Nothing,
                                   Optional ByRef dt_WHStock As DataTable = Nothing) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetStockReport]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", int_BusinessPeroidID)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", dtp_ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@RptType", strRPTType)
            BaseConn.cmd.Parameters.AddWithValue("@UpdateCost", bool_UpdateCost)
            BaseConn.cmd.Parameters.AddWithValue("@ItemArray", ItemCodeColl)
            If Not _WHID = 0 Then
                BaseConn.cmd.Parameters.AddWithValue("@WHID", _WHID)
            End If
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", LedgerID)
            BaseConn.cmd.Parameters.Add("@CalcWAC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@Stock", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@CostType", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@WHStockOUT", SqlDbType.Float).Direction = ParameterDirection.Output

            BaseConn.cmd.CommandTimeout = 1000
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = ds.Tables(0)
            If ds.Tables.Count = 2 Then
                dt_WHStock = ds.Tables(1)
            End If
            BaseConn.da.Dispose()
            dbl_Cost = BaseConn.cmd.Parameters("@CalcWAC").Value
            dbl_Stock = BaseConn.cmd.Parameters("@Stock").Value
            CostType = BaseConn.cmd.Parameters("@CostType").Value.ToString
            dbl_WHStock = BaseConn.cmd.Parameters("@WHStockOUT").Value.ToString
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function getItemWiseProfit(ByVal str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal int_BusinessPeroidID As Integer,
                                   ByVal ItemCode As String, ByVal dtp_FromDate As Date, ByVal dtp_ToDate As Date, ByVal Date1 As String, ByVal _MenuID As String,
                                   ByVal _LedgerID As Integer, ByVal _Condition As String, Optional ByVal ItemCodeColl As DataTable = Nothing) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_ItemWiseProfit]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", int_BusinessPeroidID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", dtp_ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Date1", Date1)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", _Condition)
            BaseConn.cmd.Parameters.AddWithValue("@ItemArray", ItemCodeColl)
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

    Public Function PriceTypeExistsForUser(ByVal _strPath As String, ByVal _strPwd As String, ByVal _CID As Integer, ByVal _UserID As Integer, ByVal _PriceType As String) As Boolean
        PriceTypeExistsForUser = False

        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[ItemPriceTypeExists]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@UserID", _UserID)
            BaseConn.cmd.Parameters.AddWithValue("@PriceType", _PriceType)
            BaseConn.cmd.Parameters.Add("@PriceTypeExists", SqlDbType.Bit).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            PriceTypeExistsForUser = BaseConn.cmd.Parameters("@PriceTypeExists").Value
        Catch ex As Exception
            '_ErrString = ex.Message
            'ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Return PriceTypeExistsForUser
    End Function

    Public Sub GetPriceByPriceType(ByVal _SiteID As String, ByVal _strPath As String, ByVal _strPwd As String,
                                    ByVal _ToDate As Date, ByVal _ItemCode As String, ByVal _MerchantLedger As String,
                                   ByVal _PriceType As String, ByVal _Cost As Double, ByRef _Price As Double)
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetPriceByPriceType]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", _ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@MerchantLedger", _MerchantLedger)
            BaseConn.cmd.Parameters.AddWithValue("@PriceType", _PriceType)
            BaseConn.cmd.Parameters.AddWithValue("@Cost", _Cost)
            BaseConn.cmd.Parameters.Add("@Price", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            _Price = BaseConn.cmd.Parameters("@Price").Value.ToString
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetItemDiscPercentage(ByVal _SiteID As String, ByVal _strPath As String, ByVal _strPwd As String,
                                    ByVal _ItemCode As String, ByVal _DiscType As String, ByRef _DiscPercentage As Double)
        Try
            'dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetItemDiscPercentage]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", _ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@DiscType", _DiscType)
            'BaseConn.cmd.Parameters.AddWithValue("@DiscPercentage", _DiscPercentage)

            BaseConn.cmd.Parameters.Add("@DiscPercentage", SqlDbType.Float).Direction = ParameterDirection.Output

            BaseConn.cmd.ExecuteNonQuery()
            _DiscPercentage = BaseConn.cmd.Parameters("@DiscPercentage").Value
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try

    End Sub

    Public Sub GetTaxDetails(ByVal _SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal _FormType As String,
                             ByVal _Tax As String, ByRef _TaxPercentage As Double, ByRef _TaxClaimable As Boolean)
        Try
            'dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetTaxDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@FormType", _FormType)
            BaseConn.cmd.Parameters.AddWithValue("@Tax", _Tax)
            BaseConn.cmd.Parameters.Add("@TaxPercentage", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@TaxClaimable", SqlDbType.Float).Direction = ParameterDirection.Output

            BaseConn.cmd.ExecuteNonQuery()
            _TaxPercentage = BaseConn.cmd.Parameters("@TaxPercentage").Value
            _TaxClaimable = BaseConn.cmd.Parameters("@TaxClaimable").Value
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try

    End Sub

    Public Function GetBaseDropDownList(ByVal _SiteID As String, ByVal _strPath As String, ByVal _strPwd As String,
                                   ByVal _Condition As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetBaseDropDownList]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", _Condition)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = ds.Tables(0)
            BaseConn.ds.Dispose()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Sub GetUnDeliveredQty(ByVal _SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal _ItemCode As String,
                                 ByRef _UnDelLPOQty As Double, ByRef _UnDelSOQty As Double, ByRef _UnDelJOQty As Double)
        Try
            'dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetUnDeliveredQty]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", _ItemCode)

            BaseConn.cmd.Parameters.Add("@UnDelLPOQty", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@UnDelSOQty", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@UnDelJOQty", SqlDbType.Float).Direction = ParameterDirection.Output

            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()

            _UnDelLPOQty = BaseConn.cmd.Parameters("@UnDelLPOQty").Value.ToString
            _UnDelSOQty = BaseConn.cmd.Parameters("@UnDelSOQty").Value.ToString
            _UnDelJOQty = BaseConn.cmd.Parameters("@UnDelJOQty").Value.ToString

        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try

    End Sub

    Public Function getStockValuation(ByVal str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal int_BusinessPeroidID As Integer, ByVal Str_Category As String, ByVal Str_ItemCode As String, ByVal Str_StockCount As String, ByVal dtp_FromDate As Date, ByVal dtp_ToDate As Date) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_getStockValuation]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", int_BusinessPeroidID)
            BaseConn.cmd.Parameters.AddWithValue("@Category", Str_Category)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", Str_ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@StockCount", Str_StockCount)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", dtp_ToDate)
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

    Public Function getStockFastMovingandReorder(ByVal Str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String,
                 ByVal int_BusinessPeroidID As Integer, ByVal Str_Condition As String, ByVal dtp_FromDate As Date, ByVal dtp_ToDate As Date,
                 ByVal _FormText As String, ByVal _Type As String, ByVal _OrderBy As String, Optional ByVal _WHID As String = "--ALL--", Optional ByVal ItemCollection As DataTable = Nothing) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetStockFastMovingandReorder]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", int_BusinessPeroidID)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", Str_Condition)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", dtp_ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@FormText", _FormText)
            BaseConn.cmd.Parameters.AddWithValue("@WHID", _WHID)
            BaseConn.cmd.Parameters.AddWithValue("@Type", _Type)
            BaseConn.cmd.Parameters.AddWithValue("@OrderBy", _OrderBy)
            BaseConn.cmd.Parameters.AddWithValue("@ItemArray", ItemCollection)
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function getItemsReorder(ByVal Str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal _NoofMonths As Integer, Optional ByVal _ItemCode As String = "", Optional ByVal _Flag As String = "") As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[ItemReorder]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@NoOfMonths", _NoofMonths)
            BaseConn.cmd.Parameters.AddWithValue("@Name", _ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function getItemsAging(ByVal Str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal _BSPID As Integer, ByVal _GivenDate As Date, Optional ByVal ItemColl As DataTable = Nothing, Optional ByVal _Name As String = "") As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[ItemAging]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@BSPID", _BSPID)
            BaseConn.cmd.Parameters.AddWithValue("@GivenDate", _GivenDate)
            BaseConn.cmd.Parameters.AddWithValue("@ItemArray", ItemColl)
            BaseConn.cmd.Parameters.AddWithValue("@Name", _Name)
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function Get_CheckVoucherExist(ByVal _CID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal int_BusinessPeroidID As Integer, ByVal Str_VoucherNo As String, ByVal Str_VoucherTable As String, ByVal Str_VoucherField As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[CheckVoucherExist]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", int_BusinessPeroidID)
            BaseConn.cmd.Parameters.AddWithValue("@VoucherNo", Str_VoucherNo)
            BaseConn.cmd.Parameters.AddWithValue("@VoucherTable", Str_VoucherTable)
            BaseConn.cmd.Parameters.AddWithValue("@VoucherField", Str_VoucherField)
            BaseConn.cmd.Parameters.Add("@VoucherExists", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            ErrNo = BaseConn.cmd.Parameters("@VoucherExists").Value
        Catch ex As Exception
            _ErrString = ex.Message
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return _ErrString
    End Function

    ''Public Function Get_CheckVoucherExist(ByVal Str_SiteID As String, ByVal int_BusinessPeroidID As Integer, ByVal Str_VoucherNo As String, ByVal Str_VoucherTable As String, ByVal Str_VoucherField As String, ByRef ErrNo As Integer) As String
    ''    Dim _ErrString As String = ""
    ''    ErrNo = 0
    ''    Try
    ''        dt = New DataTable
    ''        BaseConn.Open()
    ''        BaseConn.cmd = New SqlClient.SqlCommand("[sp_CheckVoucherExist]", BaseConn.cnn)
    ''        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    ''        BaseConn.cmd.Parameters.AddWithValue("@SiteID", SiteID)
    ''        BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", int_BusinessPeroidID)
    ''        BaseConn.cmd.Parameters.AddWithValue("@VoucherNo", Str_VoucherNo)
    ''        BaseConn.cmd.Parameters.AddWithValue("@VoucherTable", Str_VoucherTable)
    ''        BaseConn.cmd.Parameters.AddWithValue("@VoucherField", Str_VoucherField)
    ''        BaseConn.cmd.Parameters.Add("@VoucherExists", SqlDbType.Int).Direction = ParameterDirection.Output
    ''        BaseConn.cmd.ExecuteNonQuery()
    ''        ErrNo = BaseConn.cmd.Parameters("@VoucherExists").Value
    ''    Catch ex As Exception
    ''        _ErrString = ex.Message
    ''        ErrNo = 1
    ''    Finally
    ''        BaseConn.Close()
    ''    End Try
    ''    Return _ErrString
    ''End Function


    ''Public Function Get_MenuDetails(ByVal str_MainSiteID As String, ByVal str_UserID As String, ByRef dt_Menu As DataTable, ByRef dt_MenuOptions As DataTable) As DataTable
    ''    BaseConn.Open()
    ''    dt = New DataTable
    ''    BaseConn.cmd = New SqlClient.SqlCommand("select MenuMgt.CustomText,MenuMgt.MenuGroup, MenuMgt.MenuID from [" & str_MainSiteID & "_MenuMgt] as MenuMgt, [" & str_MainSiteID & "_UserMgt] as UserMgt, [" & str_MainSiteID & "_GroupMgtSub] as GroupPermission where UserMgt.UserName=@prm1 and GroupPermission.GroupID=UserMgt.GroupID and MenuMgt.MenuID=GroupPermission.MenuID group by MenuMgt.MenuID,MenuMgt.CustomText,MenuMgt.MenuGroup", BaseConn.cnn)
    ''    BaseConn.cmd.CommandType = CommandType.Text
    ''    BaseConn.cmd.Parameters.AddWithValue("@prm1", str_UserID)
    ''    BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    ''    BaseConn.da.Fill(dt)
    ''    dt_Menu = dt

    ''    dt1 = New DataTable
    ''    BaseConn.cmd = New SqlClient.SqlCommand("select GroupMgt.GroupID,GroupMgtSub.Options,GroupMgtSub.MenuID, GroupMgt.CreatedBy from [" & str_MainSiteID & "_GroupMgt] as  GroupMgt, [" & str_MainSiteID & "_UserMgt] as  UserMgt, [" & str_MainSiteID & "_GroupMgtSub] as GroupMgtSub where GroupMgt.GroupID=UserMgt.GroupID and UserMgt.UserName=@prm1", BaseConn.cnn)
    ''    BaseConn.cmd.CommandType = CommandType.Text
    ''    BaseConn.cmd.Parameters.AddWithValue("@prm1", str_UserID)
    ''    BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    ''    BaseConn.da.Fill(dt1)
    ''    dt_MenuOptions = dt1
    ''    BaseConn.Close()
    ''End Function

    Public Function GetFormDefaults(ByVal CID As String, ByVal _StrPath As String, ByVal _StrPwd As String, ByVal MenuID As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_StrPath, _StrPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetFormDefaults]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", MenuID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
            BaseConn.da.Dispose()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try


        Return dt
    End Function

    ''Public Function GetFormDefaults(ByVal SiteID As String, ByVal MenuID As String) As DataTable
    ''    Try
    ''        BaseConn.Open()
    ''        dt = New DataTable
    ''        BaseConn.cmd = New SqlClient.SqlCommand("select TagID,TagValue from [" & SiteID & "_FormPropertySettings] where MenuID=@prm1", BaseConn.cnn)
    ''        ''BaseConn.cmd = New SqlClient.SqlCommand("select TagID,TagValue from 101_FormPropertySettings where MenuID=@prm1)", BaseConn.cnn)
    ''        BaseConn.cmd.CommandType = CommandType.Text
    ''        BaseConn.cmd.Parameters.AddWithValue("@prm1", MenuID)
    ''        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    ''        BaseConn.da.Fill(dt)
    ''        GetFormDefaults = dt
    ''        Return GetFormDefaults
    ''    Catch ex As Exception
    ''        BaseConn.Close()
    ''    End Try

    ''End Function

    Public Sub GetMerchantExpiryDetails(ByVal Str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal _LedgerID As Integer, ByRef _RtnMsg As String, ByVal _Flag As String, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetMerchantExpiryDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.cmd.Parameters.Add("@RtnMsg", SqlDbType.NVarChar, 1000).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 1000
            BaseConn.cmd.ExecuteNonQuery()
            _RtnMsg = BaseConn.cmd.Parameters("@RtnMsg").Value.ToString
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub


    Public Function getFormGridSettings(ByVal SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal MenuID As String, _LngCode As Integer, _GridID As Integer) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetFormGridSettings]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@LngCode", _LngCode)
            BaseConn.cmd.Parameters.AddWithValue("@GridID", _GridID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
            BaseConn.da.Dispose()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    ''Public Function getFormGridSettings(ByVal SiteID As String, ByVal MenuID As String) As DataTable
    ''    Try
    ''        dt = New DataTable
    ''        BaseConn.Open()
    ''        BaseConn.cmd = New SqlClient.SqlCommand("[sp_getFormGridSettings]", BaseConn.cnn)
    ''        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    ''        BaseConn.cmd.Parameters.AddWithValue("@SiteID", SiteID)
    ''        BaseConn.cmd.Parameters.AddWithValue("@MenuID", MenuID)
    ''        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    ''        BaseConn.da.Fill(dt)
    ''    Catch ex As Exception
    ''        MsgBox("Error" & ex.Message)
    ''    Finally
    ''        BaseConn.Close()
    ''    End Try
    ''    Return dt
    ''End Function

    'Public Function GetBusinessPeriodID(ByVal SiteID As String, ByRef intBusinessPeriodID As Integer) As Integer
    '    BaseConn.Open()
    '    BaseConn.cmd = New SqlClient.SqlCommand("select max(BusinessPeriodID) from [" + SiteID + "_BusinessPeriodMaster]", BaseConn.cnn)
    '    BaseConn.cmd.CommandType = CommandType.Text
    '    BaseConn.dr = BaseConn.cmd.ExecuteReader()
    '    If BaseConn.dr.HasRows Then
    '        BaseConn.dr.Read()
    '        intBusinessPeriodID = BaseConn.dr(0).ToString
    '    End If
    'End Function

    'Public Function GetLoginUser(ByVal SiteID As String, ByRef dt_GetLoginUser As DataTable, ByVal strUserName As String, ByRef errNo As Integer) As DataTable
    '    Dim ErrStr As String = ""
    '    errNo = 0
    '    BaseConn.Open()
    '    BaseConn.cmd = New SqlClient.SqlCommand("select G.GroupName,U.UserName,U.DefaultSite from [" + SiteID + "_GroupMgt] as G inner join [" + SiteID + "_UserMgt] as U on G.GroupID=u.GroupID where U.UserName=@UserName", BaseConn.cnn)
    '    BaseConn.cmd.CommandType = CommandType.Text
    '    BaseConn.cmd.Parameters.AddWithValue("@UserName", strUserName)
    '    BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    '    BaseConn.da.Fill(dt_GetLoginUser)
    '    Return dt_GetLoginUser
    'End Function

    'Public Function ConfigParameter(ByVal SiteID As String, ByRef dt_ConfigParam As DataTable, ByRef errNo As Integer) As DataTable
    '    Dim ErrStr As String = ""
    '    SiteID = "Main"
    '    errNo = 0
    '    BaseConn.Open()
    '    BaseConn.cmd = New SqlClient.SqlCommand("select Tag,Value from [" + SiteID + "_ConfigParam]", BaseConn.cnn)
    '    BaseConn.cmd.CommandType = CommandType.Text
    '    BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    '    BaseConn.da.Fill(dt_ConfigParam)
    '    Return dt_ConfigParam
    'End Function

    ''Public Function Get_CustomerDetails(ByVal obj As Object, ByVal SiteID As String, ByVal MerchantID As String) As Object
    ''    Try
    ''        dt = New DataTable
    ''        BaseConn.Open()
    ''        BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetMerchantDetails]", BaseConn.cnn)
    ''        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    ''        BaseConn.cmd.Parameters.AddWithValue("@SiteID", obj.str_SiteID)
    ''        BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.objSalesOrderMain.int_BusinessPeriodID)
    ''        BaseConn.cmd.Parameters.AddWithValue("@MerchantID", MerchantID)
    ''        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    ''        BaseConn.da.Fill(dt)
    ''        obj.objCustomerDetails.str_Address = dt.Rows(0)("Address").ToString
    ''        obj.objCustomerDetails.str_Mobile = dt.Rows(0)("Mobile").ToString
    ''        obj.objCustomerDetails.str_Tel = dt.Rows(0)("Tel").ToString
    ''        obj.objCustomerDetails.str_Aging = dt.Rows(0)("PayTerm").ToString
    ''        obj.objCustomerDetails.str_Contact = dt.Rows(0)("Contact").ToString
    ''    Catch ex As Exception
    ''        MsgBox("Error in select customer details ")
    ''    Finally
    ''        BaseConn.Close()
    ''    End Try

    ''    Return Nothing
    ''End Function
    '''' <summary>
    '''' Used to retrice WHID,WHDesc from Warehouse master table
    '''' </summary>
    '''' <param name="DT_Combo"></param>
    '''' <param name="SiteID"></param>
    '''' <param name="MenuID"></param>
    '''' <param name="ErrNo"></param>
    '''' <Author>KM1007</Author>
    Public Sub GetDynDGVCombo(ByRef DT_Combo As DataTable, ByVal SiteID As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal MenuID As String, ByVal Header As String, ByVal Condition As String, ByRef ErrNo As Integer)
        Try
            ErrNo = 0
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetDynDGVCombo]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Header", Header)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", Condition)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(DT_Combo)
        Catch ex As Exception
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub
    ''Public Function GetData4VouMatching(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal BusinessPeriodID As String, ByVal FormName As String, ByVal Flag As String, ByVal VouNo As String, ByVal LedgerID As String, ByVal _CurrCode As String) As DataTable
    ''    Try
    ''        dt = New DataTable
    ''        BaseConn.Open(_DBPath, _DBPwd)
    ''        'BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetData4VouMatchingCheque]", BaseConn.cnn)
    ''        BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetData4VouMatching]", BaseConn.cnn)
    ''        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    ''        BaseConn.cmd.Parameters.AddWithValue("@SiteID", SiteID)
    ''        BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", BusinessPeriodID)
    ''        BaseConn.cmd.Parameters.AddWithValue("@FormName", FormName)
    ''        BaseConn.cmd.Parameters.AddWithValue("@Flag", Flag)
    ''        BaseConn.cmd.Parameters.AddWithValue("@VouNo", VouNo)
    ''        BaseConn.cmd.Parameters.AddWithValue("@LedgerID", LedgerID)
    ''        BaseConn.cmd.Parameters.AddWithValue("@CurrCode", _CurrCode)
    ''        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    ''        BaseConn.da.Fill(dt)
    ''    Catch ex As Exception
    ''        MsgBox("Error" & ex.Message)
    ''    Finally
    ''        BaseConn.Close()
    ''    End Try
    ''    Return dt
    ''End Function

    'Public Function GetData4VouMatching(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal BusinessPeriodID As String, ByVal FormName As String, ByVal VouType As String, ByVal VouNo As String, ByVal LedgerID As String, ByVal _CurrCode As String) As DataTable
    Public Function GetData4VouMatching(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal BusinessPeriodID As String, ByVal FormName As String,
                                       ByVal VouType As String, ByVal VouNo As String, ByVal LedgerID As String, ByVal _CurrCode As String,
                                        ByVal _CRDR As String, Optional ByVal _PayType As String = "", Optional ByVal _RefNo As Integer = Nothing,
                                       Optional ByVal _BCRef As Integer = Nothing) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_DBPath, _DBPwd)
            ''BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetData4VouMatching_Dynamic]", BaseConn.cnn)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetData4VouMatching]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@VouType", VouType)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@FormName", FormName)
            BaseConn.cmd.Parameters.AddWithValue("@CurrCode", _CurrCode)
            BaseConn.cmd.Parameters.AddWithValue("@CRDR", _CRDR)
            BaseConn.cmd.Parameters.AddWithValue("@PayType", _PayType)
            BaseConn.cmd.Parameters.AddWithValue("@RefNo", _RefNo)
            BaseConn.cmd.Parameters.AddWithValue("@BCRef", _BCRef)
            BaseConn.cmd.CommandTimeout = 1000
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Sub GetItemHistory(ByRef dt_history As DataTable, ByVal _strPath As String, ByVal _strPwd As String, ByVal int_BusinessPeriodID As Integer, ByVal Str_LedgerID As String, ByVal Str_SiteID As String, ByVal Str_ItemCode As String, ByVal Str_Flag As String, ByVal dtp_FromDate As Date, ByVal dtp_ToDate As Date, ByVal bool_Status As Boolean, Optional ByVal bool_ViewAllMerchant As Boolean = False, Optional ByVal str_MCCBAll As String = "")
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetItemHistory]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", Str_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", Str_ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@ViewAllMerchant", bool_ViewAllMerchant)
            BaseConn.cmd.Parameters.AddWithValue("@MCCBAll", str_MCCBAll)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", dtp_ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Status", bool_Status)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
            dt_history = dt
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub BaseCurrencyFormat(ByVal _strSiteID As String, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _CurrencyCode As String, ByRef _DecimalPoint As String, ByRef _MajorCurrency As String, ByRef _MinorCurrency As String)
        'Try
        '    dt = New DataTable
        '    BaseConn.Open()
        '    BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetMerchantDetails]", BaseConn.cnn)
        '    BaseConn.cmd.CommandType = CommandType.StoredProcedure
        '    BaseConn.cmd.Parameters.AddWithValue("@SiteID", obj.str_SiteID)
        '    BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.objSalesOrderMain.int_BusinessPeriodID)
        '    BaseConn.cmd.Parameters.AddWithValue("@MerchantID", MerchantID)
        '    BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
        '    BaseConn.da.Fill(dt)
        '    _DecimalPoint = dt.Rows(0)("Address").ToString
        '    _MajorCurrency = dt.Rows(0)("Mobile").ToString
        '    _MinorCurrency = dt.Rows(0)("Tel").ToString

        'Catch ex As Exception
        '    MsgBox("Error in select customer details ")
        'Finally
        '    BaseConn.Close()
        'End Try

    End Sub
    'Public Function Get_LedgerStatement(ByVal _strPath As String, ByVal _strpwd As String, ByVal _SiteID As String, ByVal _Ledger As Integer,
    '                                    ByVal _FrmDate As Date, ByVal _ToDate As Date, ByVal _Condition As String, ByVal _LedgerType As String, ByVal _IncludePdc As Boolean) As DataTable
    '    Try
    '        dt = New DataTable
    '        BaseConn.Open(_strPath, _strpwd)

    '        BaseConn.cmd = New SqlClient.SqlCommand("[sp_LedgerStatement]", BaseConn.cnn)

    '        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    '        BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
    '        BaseConn.cmd.Parameters.AddWithValue("@Ledger", _Ledger)
    '        BaseConn.cmd.Parameters.AddWithValue("@FrmDate", _FrmDate)
    '        BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
    '        BaseConn.cmd.Parameters.AddWithValue("@Condition", _Condition.ToUpper)
    '        BaseConn.cmd.Parameters.AddWithValue("@LedgerType", _LedgerType.ToUpper)
    '        BaseConn.cmd.Parameters.AddWithValue("@PDC", _IncludePdc)
    '        'BaseConn.cmd.Parameters.AddWithValue("@BSPeriod", _BusPeriod)
    '        'BaseConn.cmd.Parameters.AddWithValue("@BusStartDate", _BusStartDate)
    '        BaseConn.cmd.CommandTimeout = 1000

    '        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    '        BaseConn.da.Fill(dt)
    '    Catch ex As Exception
    '        MsgBox("Error" & ex.Message)
    '    Finally
    '        BaseConn.Close()
    '    End Try
    '    Return dt
    'End Function

    Public Function GetLedgerStatement(ByVal _strPath As String, ByVal _strpwd As String, ByVal _CID As String, ByVal _Ledger As DataTable,
                                        ByVal _FrmDate As Date, ByVal _ToDate As Date, ByVal _ZeroSuppress As Boolean, ByVal _ActiveOnly As Boolean, ByVal _ReportName As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strpwd)

            BaseConn.cmd = New SqlClient.SqlCommand("[GetLedgerStatement]", BaseConn.cnn)

            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@Ledger", _Ledger)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FrmDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@ZeroSuppress", _ZeroSuppress)
            BaseConn.cmd.Parameters.AddWithValue("@ActiveOnly", _ActiveOnly)
            BaseConn.cmd.Parameters.AddWithValue("@ReportName", _ReportName)
            BaseConn.cmd.CommandTimeout = 1000

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function


    Public Function GetLedgerAdvance(ByVal _strPath As String, ByVal _strpwd As String, ByVal _CID As String, ByVal _Ledger As DataTable,
                                        ByVal _FrmDate As Date, ByVal _ToDate As Date, ByVal _ZeroSuppress As Boolean, ByVal _ActiveOnly As Boolean, ByVal _ReportName As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strpwd)

            BaseConn.cmd = New SqlClient.SqlCommand("[GetLedgerAdvance]", BaseConn.cnn)

            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@Ledger", _Ledger)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FrmDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@ZeroSuppress", _ZeroSuppress)
            BaseConn.cmd.Parameters.AddWithValue("@ActiveOnly", _ActiveOnly)
            BaseConn.cmd.Parameters.AddWithValue("@ReportName", _ReportName)
            BaseConn.cmd.CommandTimeout = 1000

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function Get_ExpensesAndIncomeStatement(ByVal _strPath As String, ByVal _strpwd As String, ByVal _SiteID As String, ByVal _Ledger As Integer,
                                        ByVal _FrmDate As Date, ByVal _ToDate As Date, ByVal _Condition As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strpwd)

            BaseConn.cmd = New SqlClient.SqlCommand("[sp_LedgerStatementExpensesAndIncome]", BaseConn.cnn)

            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Ledger", _Ledger)
            BaseConn.cmd.Parameters.AddWithValue("@FrmDate", _FrmDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", _Condition.ToUpper)
            'BaseConn.cmd.Parameters.AddWithValue("@BSPeriod", _BusPeriod)
            'BaseConn.cmd.Parameters.AddWithValue("@BusStartDate", _BusStartDate)
            BaseConn.cmd.CommandTimeout = 1000

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
    Public Function Get_LedgerStatementShowAcOnly(ByVal _strPath As String, ByVal _strpwd As String, ByVal _SiteID As String, ByVal _Ledger As Integer, ByVal _FrmDate As Date,
                                       ByVal _ToDate As Date, ByVal _showacconly As Boolean) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strpwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_LedgerStatementAdvanceonly]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Ledger", _Ledger)
            BaseConn.cmd.Parameters.AddWithValue("@FrmDate", _FrmDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            'BaseConn.cmd.Parameters.AddWithValue("@BSPeriod", _BusPeriod)
            'BaseConn.cmd.Parameters.AddWithValue("@BusStartDate", _BusStartDate)
            BaseConn.cmd.CommandTimeout = 1000

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
    Public Function Get_AdvanceSummary(ByVal _strPath As String, ByVal _strpwd As String, ByVal _SiteID As String, ByVal _Ledger As Integer, ByVal _Condition As String, ByVal _FrmDate As Date,
                                           ByVal _ToDate As Date, ByVal _showacconly As Boolean) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strpwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_AdvanceSummary]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Ledger", _Ledger)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", _Condition)
            BaseConn.cmd.Parameters.AddWithValue("@FrmDate", _FrmDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            'BaseConn.cmd.Parameters.AddWithValue("@BSPeriod", _BusPeriod)
            'BaseConn.cmd.Parameters.AddWithValue("@BusStartDate", _BusStartDate)
            BaseConn.cmd.CommandTimeout = 1000

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Sub Elog_Insert(ByVal strSiteID As String, ByVal _strDBPath As String, ByVal _strDBPwd As String, ByVal _intBusinessPeriodID As Integer, ByVal strUsernName As String, ByVal dtpCreatedDate As DateTime, ByVal strGroupName As String, ByVal strFormName As String, ByVal intErrorNo As Integer, ByVal strDescription As String, ByVal strErrorDescription As String, ByVal intLevel As Integer, ByVal intType As Integer, ByVal intSource As Integer, ByRef ErrNo As Integer)
        'Dim _ErrString As String = ""
        'ErrNo = 0
        Try
            BaseConn.Open(_strDBPath, _strDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[ElogUpdate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", strSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _intBusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", strUsernName)
            BaseConn.cmd.Parameters.AddWithValue("@DateTime_", dtpCreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@GroupName", strGroupName)
            BaseConn.cmd.Parameters.AddWithValue("@FormName", strFormName)
            BaseConn.cmd.Parameters.AddWithValue("@ErrorNo", intErrorNo)
            BaseConn.cmd.Parameters.AddWithValue("@Desc_", strDescription)
            BaseConn.cmd.Parameters.AddWithValue("@ErrorDesc", strErrorDescription)
            BaseConn.cmd.Parameters.AddWithValue("@Lvl", intLevel)
            BaseConn.cmd.Parameters.AddWithValue("@Type", intType)
            BaseConn.cmd.Parameters.AddWithValue("@Source", intSource)
            BaseConn.cmd.ExecuteNonQuery()
        Catch ex As Exception
            '_ErrString = ex.Message
            'ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        'Elog_Insert = _ErrString
    End Sub

    Public Function getSpecificReport(ByVal _SiteID As String, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal Condition As String, ByVal Flag As String, Optional ByVal dtItem As DataTable = Nothing, Optional ByVal Filter As String = "", Optional ByVal Condition1 As Date = Nothing) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetSpecificReport]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", Condition)
            BaseConn.cmd.Parameters.AddWithValue("@Condition1", Condition1)
            BaseConn.cmd.Parameters.AddWithValue("@Filter", Filter)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Flag)
            BaseConn.cmd.Parameters.AddWithValue("@DT", dtItem)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
    Public Function Get_SpecificExcelExport(ByVal _SiteID As String, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal Condition As String, ByVal Flag As String, Optional ByVal dtItem As DataTable = Nothing) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetSpecificExcelExport]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", Condition)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Flag)
            BaseConn.cmd.Parameters.AddWithValue("@DT", dtItem)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function


    Public Function GetProfitability(ByVal _strPath As String, ByVal _strDBPwd As String, ByVal _SiteID As String, ByVal _LedgerID As DataTable, ByVal _FromDate As Date,
                                     ByVal _ToDate As Date, ByVal _Flag As String, ByVal _Suppress As Boolean, ByRef _ErrNo As Integer, ByRef _ErrStr As String) As DataTable
        dt = New DataTable
        _ErrStr = ""
        _ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetProfitability]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            _ErrStr = ex.Message
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function GetInvoiceOutstanding(ByVal _strPath As String, ByVal _strDBPwd As String, ByVal _SiteID As String, ByVal _LedgerID As DataTable, ByVal _FromDate As Date,
                                     ByVal _ToDate As Date, ByVal _Flag As String, ByVal _SalesmanLedger As String,
                                     ByVal _Suppress As Boolean, ByRef _ErrNo As Integer, ByRef _ErrStr As String) As DataTable
        dt = New DataTable
        _ErrStr = ""
        _ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetInvoiceOutstanding]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManLedger", _SalesmanLedger)
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            _ErrStr = ex.Message
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function GetProfitabilityItem(ByVal _strPath As String, ByVal _strDBPwd As String, ByVal _SiteID As String, ByVal _ItemCode As DataTable, ByVal _FromDate As Date,
                                     ByVal _ToDate As Date, ByRef _ErrNo As Integer, ByRef _ErrStr As String) As DataTable
        dt = New DataTable
        _ErrStr = ""
        _ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetProfitabilityItem]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", _ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            _ErrStr = ex.Message
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function Get_SalesManSales(ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _intBusinessPeriodID As Integer, ByVal _Ledger As String, ByVal _FrmDate As Date,
                                            ByVal _ToDate As Date, ByVal _Condition As String) As DataTable
        Try
            BaseConn.Open(_strDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetSalesManSales]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@Name", _Ledger)
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _intBusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FrmDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", _Condition)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            dt = New DataTable
            BaseConn.da.Fill(dt)
            'dt = dt
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function Get_MonthlySales(ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _Ledger As DataTable,
                                     ByVal _NoofMonths As Integer, ByVal _FromDate As Date, ByVal _ToDate As Date, ByVal _IsSalesMan As Boolean,
                                     ByVal _Type As String, Optional ByVal _Day1 As Integer = 0, Optional ByVal _Day2 As Integer = 31, Optional _Condition As String = "") As DataTable
        Try
            BaseConn.Open(_strDBPath, _StrDBPwd)
            If _Type = "ITEM" Then
                BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetMonthlySalesItem]", BaseConn.cnn)
            Else
                BaseConn.cmd = New SqlClient.SqlCommand("[GetMonthlySales]", BaseConn.cnn)
            End If
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            If _Type = "ITEM" Then
                BaseConn.cmd.Parameters.AddWithValue("@ItemCode", _Ledger)
            Else
                BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _Ledger)
                BaseConn.cmd.Parameters.AddWithValue("@Day1", _Day1)
                BaseConn.cmd.Parameters.AddWithValue("@Day2", _Day2)
            End If

            BaseConn.cmd.Parameters.AddWithValue("@NoOfMonths", _NoofMonths)
            BaseConn.cmd.Parameters.AddWithValue("@StartDate", _FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@EndDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@IsSalesMan", _IsSalesMan)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", _Condition)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            dt = New DataTable
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function GetMonthlySales(ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByVal _CID As String, ByVal _LedgerID As DataTable,
                                     ByVal _FromDate As Date, ByVal _ToDate As Date, ByVal _IsSalesMan As Boolean, ByVal _Day1 As Integer, ByVal _Day2 As Integer) As DataTable
        Try
            BaseConn.Open(_strDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetMonthlySales]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@IsSalesMan", _IsSalesMan)
            BaseConn.cmd.Parameters.AddWithValue("@Day1", _Day1)
            BaseConn.cmd.Parameters.AddWithValue("@Day2", _Day2)

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            dt = New DataTable
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function GetMonthlySalesItem(ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByVal _CID As String,
                                     ByVal _FromDate As Date, ByVal _ToDate As Date, ByVal _ItemCode As DataTable,
                                     ByVal _Day1 As Integer, ByVal _Day2 As Integer) As DataTable
        Try
            BaseConn.Open(_strDBPath, _StrDBPwd)

            BaseConn.cmd = New SqlClient.SqlCommand("[GetMonthlySalesItem]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@Itemcode", _ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Day1", _Day1)
            BaseConn.cmd.Parameters.AddWithValue("@Day2", _Day2)

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            dt = New DataTable
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Sub Get_ExpPopUp(ByRef popup As Boolean, ByVal str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal str_UserName As String, ByVal str_flag As String, ByVal Bool_Hide As Boolean)
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetExpiryPopUp]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", str_UserName)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", str_flag)
            BaseConn.cmd.Parameters.AddWithValue("@Hide", Bool_Hide)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            If ds.Tables(0).Rows(0)("ShowPopUp").ToString = True AndAlso ds.Tables(0).Rows(0)("HidePopUp").ToString = True Then
                popup = True
            Else
                popup = False
            End If

        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub Get_ExpiryDetails(ByRef dt_EmpExp As DataTable, ByRef dt_CusExp As DataTable, ByRef dt_VenExp As DataTable, ByRef dt_Asset As DataTable, ByRef dt_AdvRequest As DataTable, ByVal str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal str_UserName As String, ByVal str_flag As String, ByVal Bool_Hide As Boolean)
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetExpiryPopUp]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", str_UserName)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", str_flag)
            BaseConn.cmd.Parameters.AddWithValue("@Hide", Bool_Hide)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt_CusExp = ds.Tables(0)
            dt_VenExp = ds.Tables(1)
            dt_EmpExp = ds.Tables(2)
            dt_Asset = ds.Tables(3)
            dt_AdvRequest = ds.Tables(4)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub Update_PopUp(ByVal str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal str_UserName As String, ByVal str_Flag As String, ByVal bool_hide As Boolean)
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetExpiryPopUp]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", str_UserName)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@Hide", bool_hide)
            BaseConn.cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetItemMasterColumn(ByVal Str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByRef _dt As DataTable, ByVal _Flag As String, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        _ErrNo = 0
        _ErrStr = ""
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_ItemMasterColumn]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@dt_ProductGrp", _dt)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _dt = ds.Tables(0)
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub MerchantGrouping(ByVal Str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByRef _dt As DataTable, ByVal strParentLedgerID As String, ByVal _Flag As String, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        _ErrNo = 0
        _ErrStr = ""
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[MerchantGrouping]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@ParentLedgerID", strParentLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@DT", _dt)
            BaseConn.cmd.ExecuteNonQuery()
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub UnderCostItemsCount(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _StrSiteID As String,
                          ByVal _VouNo As String, ByVal _Flag As String, ByRef UnderCostCount As Integer)

        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[IsItemshasUnderCost]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _StrSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", _VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.Add("@UnderCostCount", SqlDbType.Float).Direction = ParameterDirection.Output

            BaseConn.cmd.ExecuteNonQuery()
            UnderCostCount = BaseConn.cmd.Parameters("@UnderCostCount").Value

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub Get_DayBook(ByVal _strPath As String, ByVal _strpwd As String, ByVal _CID As String, ByVal _DateRange As Boolean, ByVal _FrmDate As Date, ByVal _ToDate As Date, ByRef dt As DataTable, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        _ErrNo = 0
        _ErrStr = ""
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strpwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetDayBook]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@DateRange", _DateRange)
            BaseConn.cmd.Parameters.AddWithValue("@FrmDate", _FrmDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Function Get_Aging(ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByRef _dt As DataTable, ByVal _Ledger As String, ByVal _DateType As String, ByVal _FrmDate As Date,
                                           ByVal _ToDate As Date, ByVal _Condition As String, ByVal Date_Type As String) As DataTable
        Try
            BaseConn.Open(_strDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetAgingReport]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@DT", _dt)
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _Ledger)
            BaseConn.cmd.Parameters.AddWithValue("@DateType", _DateType)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FrmDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Condition)
            BaseConn.cmd.Parameters.AddWithValue("@Date_Type", Date_Type)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.cmd.CommandTimeout = 500
            dt = New DataTable
            BaseConn.da.Fill(dt)
            'dt = dt
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function Load_ComboDistinctDT(ByVal _strPath As String, ByVal _strPwd As String, ByVal TableName As String, ByVal DisplayMember As String, ByVal ValueMember As String, ByVal Condition As String, ByVal Sorting As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_LoadComboDistinctDT]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@TableName", TableName)
            BaseConn.cmd.Parameters.AddWithValue("@DisplayMember", DisplayMember)
            BaseConn.cmd.Parameters.AddWithValue("@ValueMember", ValueMember)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", Condition)
            BaseConn.cmd.Parameters.AddWithValue("@Sorting", Sorting)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function GetLedgerDetails(ByVal _strPath As String, ByVal _strPwd As String, ByVal _LedgerID As Integer) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetLedgerDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Sub GetRatio(ByVal _strPath As String, ByVal _strPwd As String, ByVal _strSiteID As String, ByVal _ItemCode As String, ByRef _Ratio As Double, ByRef _strPrimaryUnit As String, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        _ErrNo = 0
        _ErrStr = ""
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetRatio]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", _ItemCode)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
            _Ratio = dt.Rows(0)("Ratio").ToString
            _strPrimaryUnit = dt.Rows(0)("PrimaryUnit").ToString
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetItemUOMRatio(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _strSiteID As String, ByVal _ItemCode As String,
                               ByVal _UOM As String, ByRef _UOMRatio As Double, ByRef _ErrNo As Integer, ByRef _ErrStr As String)

        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetItemUOMRatio]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _strSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Itemcode", _ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@UOM", _UOM)
            BaseConn.cmd.Parameters.Add("@UOMRatio", SqlDbType.Float).Direction = ParameterDirection.Output

            BaseConn.cmd.ExecuteNonQuery()
            _UOMRatio = BaseConn.cmd.Parameters("@UOMRatio").Value

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetUnit(ByVal _strPath As String, ByVal _strPwd As String, ByVal _strSiteID As String, ByVal _ItemCode As String, ByRef dt_Unit As DataTable, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        _ErrNo = 0
        _ErrStr = ""
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetUnit]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", _ItemCode)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
            dt_Unit = dt
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub Get_WHMaster(ByVal SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByRef dt_WHMaster As DataTable, ByVal Condition As String, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        _ErrNo = 0
        _ErrStr = ""
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[LoadWHMaster]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", Condition)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
            dt_WHMaster = dt
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub

    'Public Sub Update_WPSMaster(ByVal _strPath As String, ByVal _strPWD As String, ByVal _SiteID As String, ByVal WPSID As String, ByVal BankName As String, ByVal AllowHeader As Boolean, ByVal CompanyCode As String, ByVal PreBankName As String, ByVal Flag As String, ByVal dt_Main As DataTable, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
    '    _ErrNo = 0
    '    _ErrStr = ""
    '    Try
    '        dt = New DataTable
    '        BaseConn.Open(_strPath, _strPWD)
    '        BaseConn.cmd = New SqlClient.SqlCommand("[sp_WPSMasterUpdate]", BaseConn.cnn)
    '        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    '        BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
    '        BaseConn.cmd.Parameters.AddWithValue("@PreBankName", PreBankName)
    '        BaseConn.cmd.Parameters.AddWithValue("@WPSID", WPSID)
    '        BaseConn.cmd.Parameters.AddWithValue("@BankName", BankName)
    '        BaseConn.cmd.Parameters.AddWithValue("@AllowHeader", AllowHeader)
    '        BaseConn.cmd.Parameters.AddWithValue("@CompanyCode", CompanyCode)
    '        BaseConn.cmd.Parameters.AddWithValue("@Flag", Flag)
    '        BaseConn.cmd.Parameters.AddWithValue("@DT", dt_Main)
    '        BaseConn.cmd.ExecuteNonQuery()
    '    Catch ex As Exception
    '        _ErrNo = 1
    '        _ErrStr = ex.Message.ToString
    '    Finally
    '        BaseConn.Close()
    '    End Try
    'End Sub


    Public Sub Update_WPSMaster(ByVal _strPath As String, ByVal _strPWD As String, ByVal _SiteID As String, ByVal WPSID As String, ByVal BankName As String, ByVal AllowHeader As Boolean, ByVal CompanyCode As String, ByVal RoutingCode As String, ByVal PreBankName As String, ByVal Flag As String, ByVal dt_Main As DataTable, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        _ErrNo = 0
        _ErrStr = ""
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPWD)
            BaseConn.cmd = New SqlClient.SqlCommand("[WPSMasterUpdate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@PreBankName", PreBankName)
            BaseConn.cmd.Parameters.AddWithValue("@WPSID", WPSID)
            BaseConn.cmd.Parameters.AddWithValue("@BankName", BankName)
            BaseConn.cmd.Parameters.AddWithValue("@AllowHeader", AllowHeader)
            BaseConn.cmd.Parameters.AddWithValue("@CompanyCode", CompanyCode)
            BaseConn.cmd.Parameters.AddWithValue("@RoutingCode", RoutingCode)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Flag)
            BaseConn.cmd.Parameters.AddWithValue("@DT", dt_Main)
            BaseConn.cmd.ExecuteNonQuery()
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub Get_WPSReport(ByVal _strPath As String, ByVal _strPWD As String, ByVal _strSiteID As String, ByVal _TableName As String, ByVal _FieldName As String, ByVal _CustomText As String, ByVal _Condition As String, ByRef dt As DataTable, ByVal dtpSalaryMonth As Date, ByVal _strFormat As String, ByVal _Flag As String, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        _ErrNo = 0
        _ErrStr = ""
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPWD)
            BaseConn.cmd = New SqlClient.SqlCommand("[WPSReport]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _strSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@TableName", _TableName)
            BaseConn.cmd.Parameters.AddWithValue("@FieldName", _FieldName)
            BaseConn.cmd.Parameters.AddWithValue("@CustomText", _CustomText)
            BaseConn.cmd.Parameters.AddWithValue("@Format", _strFormat)
            BaseConn.cmd.Parameters.AddWithValue("@SalaryMonth", dtpSalaryMonth)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", _Condition)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub

    'Public Sub Get_WPSReport(ByVal _strPath As String, ByVal _strPWD As String, ByVal _strSiteID As String, ByVal _TableName As String, ByVal _FieldName As String, ByVal _CustomText As String, ByVal _Condition As String, ByRef dt As DataTable, ByVal _Flag As String, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
    '    _ErrNo = 0
    '    _ErrStr = ""
    '    Try
    '        dt = New DataTable
    '        BaseConn.Open(_strPath, _strPWD)
    '        BaseConn.cmd = New SqlClient.SqlCommand("[sp_WPSReport]", BaseConn.cnn)
    '        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    '        BaseConn.cmd.Parameters.AddWithValue("@SiteID", _strSiteID)
    '        BaseConn.cmd.Parameters.AddWithValue("@TableName", _TableName)
    '        BaseConn.cmd.Parameters.AddWithValue("@FieldName", _FieldName)
    '        BaseConn.cmd.Parameters.AddWithValue("@CustomText", _CustomText)
    '        BaseConn.cmd.Parameters.AddWithValue("@Condition", _Condition)
    '        BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
    '        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    '        BaseConn.da.Fill(dt)
    '    Catch ex As Exception
    '        _ErrNo = 1
    '        _ErrStr = ex.Message.ToString
    '    Finally
    '        BaseConn.Close()
    '    End Try
    'End Sub

    Public Function GetMissMatchItems(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal GivenItems As DataTable, ByRef ErrNo As Integer) As DataTable
        Dim _ErrString As String = ""
        ErrNo = 0
        dt = New DataTable
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetExcelMisMatchItems]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@ReceivedItemDT", GivenItems)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = ds.Tables(0)
        Catch ex As Exception
            _ErrString = ex.Message
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function GetItemsInItemMaster(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal Flag As String, ByVal GivenItems As DataTable, ByRef ErrNo As Integer) As DataTable
        Dim _ErrString As String = ""
        ErrNo = 0
        dt = New DataTable
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetItemsInItemMaster]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@ReceivedItemDT", GivenItems)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = ds.Tables(0)
        Catch ex As Exception
            _ErrString = ex.Message
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
    Public Function GetBarCodeInItemMaster(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal GivenItems As DataTable, ByRef ErrNo As Integer) As DataTable
        Dim _ErrString As String = ""
        ErrNo = 0
        dt = New DataTable
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetBarCodeInItemMaster]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@ReceivedItemDT", GivenItems)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = ds.Tables(0)
        Catch ex As Exception
            _ErrString = ex.Message
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
    Public Function GetBarCodeNotInItemMaster(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal GivenItems As DataTable, ByRef ErrNo As Integer) As DataTable
        Dim _ErrString As String = ""
        ErrNo = 0
        dt = New DataTable
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetBarCodeNotInItemMaster]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@ReceivedItemDT", GivenItems)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = ds.Tables(0)
        Catch ex As Exception
            _ErrString = ex.Message
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
    Public Sub Get_HRMain(ByRef Obj As csHRMain, ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByVal ErrNo As Integer, ByVal ErrStr As String, Optional ByVal _CurVacation As String = Nothing)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_strDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetHRMain]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", Obj.str_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Category", Obj.str_Category)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", Obj.dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", Obj.dtp_ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Date1", Obj._Date)
            BaseConn.cmd.Parameters.AddWithValue("@CurVacation", _CurVacation)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            Obj.dt_Main = ds.Tables(0)
            If ds.Tables.Count >= 2 Then
                Obj.dt_sub = ds.Tables(1)
            End If
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ""
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub HR_OTUpdate(ByVal _dtOT As DataTable, ByVal _dtsub As DataTable, ByVal _UserName As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As Integer, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        _ErrNo = 0
        _ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[UpdateOT]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", _UserName)
            BaseConn.cmd.Parameters.AddWithValue("@DTOT", _dtOT)
            BaseConn.cmd.Parameters.AddWithValue("@DTProject", _dtsub)
            BaseConn.cmd.ExecuteNonQuery()
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Load_ComboMCCB(ByVal _strPath As String, ByVal _strPwd As String, ByVal TableName As String, ByVal ColumnName As String, ByVal Condition As String, ByVal SortBy As String, ByVal Sorting As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[LoadComboMCCB]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@TableName", TableName)
            BaseConn.cmd.Parameters.AddWithValue("@ColumnName", ColumnName)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", Condition)
            BaseConn.cmd.Parameters.AddWithValue("@SortBy", SortBy)
            BaseConn.cmd.Parameters.AddWithValue("@Sorting", Sorting)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function Get_LedgerStatementSummary(ByVal _strPath As String, ByVal _strpwd As String, ByVal _SiteID As String, ByVal _Ledger As Integer,
                                              ByVal _FrmDate As Date, ByVal _ToDate As Date, ByVal _LedgerSummary As Boolean, ByVal _LedgerType As String,
                                               ByVal _Condition As String, ByVal _ZeroSuppress As Boolean) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strpwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_LedgerStatementSummary]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Ledger", _Ledger)
            BaseConn.cmd.Parameters.AddWithValue("@FrmDate", _FrmDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerType", _LedgerType)
            BaseConn.cmd.Parameters.AddWithValue("@Type", _Condition)
            BaseConn.cmd.Parameters.AddWithValue("@ZeroSuppress", _ZeroSuppress)
            BaseConn.cmd.CommandTimeout = 1000
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Sub DialogItemPrice(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _BusinessPeriodID As Integer, ByVal _PriceShow As Boolean, ByVal _Flag As String, ByRef dt_ItemPrice As DataTable, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            dt = New DataTable
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[DialogItemPrice]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@PriceShow", _PriceShow)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
        dt_ItemPrice = dt
    End Sub
    Public Sub ItemSerial(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal FromDate As Date, ByVal ToDate As Date, ByVal _Date As String, ByVal _ItemCode As String, ByVal _MenuID As String, ByVal _Flag As String,
                          ByVal MerchantID As String, ByRef dt_ItemSerial As DataTable, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            dt = New DataTable
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetIMEIReport]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", _ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Date1", _Date)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", _MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@MerchantID", MerchantID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
        dt_ItemSerial = dt
    End Sub
    Public Sub GetUnMatchedVoucher(ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _BSPeriod As Integer,
                         ByVal _LedgerID As Integer, ByRef _DTCredit As DataTable, ByRef _DTDebit As DataTable, ByVal ErrNo As Integer, ByVal ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_strDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetUnMatchedVoucher]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BSPeriod)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _DTCredit = ds.Tables(0)
            _DTDebit = ds.Tables(1)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ""
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub VoucherMatchingUpdated(ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _VouType As String, ByVal _CRDR As String,
                         ByVal _FormName As String, ByVal _VouNo As String, ByVal _BCRef As Integer, ByVal _RefNo As String, ByVal _TCNetAmount As Double,
                         ByVal _VouMatching As DataTable, ByVal _BSPeriodID As Integer, ByVal _CreatedBy As String, ByVal _CreatedDate As Date, ByVal _ErrNo As Integer, ByVal _ErrStr As String)
        _ErrNo = 0
        _ErrStr = ""
        Try
            BaseConn.Open(_strDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[VoucherMatchingUpdate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@VouType", _VouType)
            BaseConn.cmd.Parameters.AddWithValue("@CRDR", _CRDR)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", _VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@FormName", _FormName)
            BaseConn.cmd.Parameters.AddWithValue("@BC_Ref", _BCRef)
            BaseConn.cmd.Parameters.AddWithValue("@RefNo", _RefNo)
            BaseConn.cmd.Parameters.AddWithValue("@TCNetAmount", _TCNetAmount)
            BaseConn.cmd.Parameters.AddWithValue("@MatchingDT", _VouMatching)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BSPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", _CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", _CreatedDate)
            BaseConn.cmd.ExecuteNonQuery()
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ""
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function GetItemParamValues(ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _ItemCode As String,
                               ByVal _DTDynColDetails As DataTable, ByVal ErrNo As Integer, ByVal ErrStr As String) As DataTable
        GetItemParamValues = New DataTable

        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_strDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetItemParamValues]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", _ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@DynCol", _DTDynColDetails)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            GetItemParamValues = ds.Tables(0)

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ""
        Finally
            BaseConn.Close()
        End Try
    End Function

    Public Function GetSalesMan(ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByVal _CID As String,
                              ByVal _DTSalesMan As DataTable, ByVal _Flag As String, ByVal ErrNo As Integer, ByVal ErrStr As String) As DataTable
        GetSalesMan = New DataTable

        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_strDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetSalesMan]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManDT", _DTSalesMan)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            GetSalesMan = ds.Tables(0)
            BaseConn.da.Dispose()
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ""
        Finally
            BaseConn.Close()
        End Try
    End Function

    Public Sub GetConversionHistory(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _FormType As String, ByVal _VouNo As String, ByRef _dtHistory As DataTable, ByRef _ErrNo As Integer, ByRef _ErrStr As String, Optional ByVal _Type As String = "")
        _ErrNo = 0
        _ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetConversionHistory]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Form", _FormType)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", _VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@Type", _Type)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _dtHistory = ds.Tables(0)
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetMinMaxValue(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _BusinessPeriodID As Integer, ByVal _ItemCode As String, ByVal _FromDate As Date, ByVal _ToDate As Date, ByRef _Stock As Double, ByRef _POStock As Double, ByRef _SOStock As Double, ByRef _MinStock As Double, ByRef _MaxStock As Double, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        _ErrNo = 0
        _ErrStr = ""

        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetMinMaxValue]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", _ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.Add("@Stock", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@POStock", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@SOStock", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@MinStock", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@MaxStock", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            _Stock = BaseConn.cmd.Parameters("@Stock").Value
            _POStock = BaseConn.cmd.Parameters("@POStock").Value
            _SOStock = BaseConn.cmd.Parameters("@SOStock").Value
            _MinStock = BaseConn.cmd.Parameters("@MinStock").Value
            _MaxStock = BaseConn.cmd.Parameters("@MaxStock").Value
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try

    End Sub

    Public Sub GetLedgerCumBalance(ByVal _strDBPath As String, ByVal _strDBPwd As String, ByVal _strSiteID As String, ByVal _intBusinessPeriodID As Integer, ByVal _intLedgerID As Integer, ByVal _intParentAccount As Integer, ByVal _ToDate As Date, ByVal _strFlag As String, ByRef _dtBalance As DataTable, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        _ErrNo = 0
        _ErrStr = ""

        Try
            BaseConn.Open(_strDBPath, _strDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetLedgerCumBalance]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _strSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _intBusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _intLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@ParentAccount", _intParentAccount)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _strFlag)
            BaseConn.cmd.Parameters.AddWithValue("@IsReplicate", 1)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _dtBalance = ds.Tables(0)
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try

    End Sub

    Public Function GetLedgerBalance(ByVal Str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String,
                                   ByVal dtp_from As Date, ByVal dtp_to As Date, ByVal _ZeroSuppress As Boolean,
                                   ByVal _ShowInActive As Boolean, ByVal _Type As Integer, ByVal _LedgerID As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetLedgerBalance]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", dtp_from)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", dtp_to)
            BaseConn.cmd.Parameters.AddWithValue("@InActive", _ShowInActive)
            BaseConn.cmd.Parameters.AddWithValue("@ZeroSuppress", _ZeroSuppress)
            BaseConn.cmd.Parameters.AddWithValue("@PostClose", _Type)

            BaseConn.cmd.CommandTimeout = 1000
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

    Public Function GetExpenseLedgerWithoutCOGS(ByVal Str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String,
                                   ByVal dtp_from As Date, ByVal dtp_to As Date, ByVal _ReportLevel As String, ByVal _ZeroSuppress As Boolean,
                                   ByVal _ShowInActive As Boolean, ByVal _Type As Integer, ByVal _LedgerID As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetExpenseLedgerWithoutCOGS]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", dtp_from)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", dtp_to)
            BaseConn.cmd.Parameters.AddWithValue("@ReportLevel", _ReportLevel)
            BaseConn.cmd.Parameters.AddWithValue("@ZeroSuppress", _ZeroSuppress)
            BaseConn.cmd.Parameters.AddWithValue("@ShowInActive", _ShowInActive)
            BaseConn.cmd.Parameters.AddWithValue("@Type", _Type)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.CommandTimeout = 1000
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

    Public Sub GetLedgerCumBalanceDashboard(ByVal _strDBPath As String, ByVal _strDBPwd As String, ByVal _strSiteID As String,
                                            ByVal _intBusinessPeriodID As Integer, ByVal _LedgerID As Integer, ByVal _ToDate As Date, ByVal _TopRecord As Integer,
                                            ByVal _SortBy As String, ByRef _dtResult As DataTable, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        _ErrNo = 0
        _ErrStr = ""

        Try
            BaseConn.Open(_strDBPath, _strDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetLedgerDetailsDashboard]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _strSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _intBusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@TopRecord", _TopRecord)
            BaseConn.cmd.Parameters.AddWithValue("@SortBy", _SortBy)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _dtResult = ds.Tables(1)
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try

    End Sub

    Public Sub Get_UserPermissionFlag(ByVal Str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal _Flag As String, ByVal _UserName As String,
                                      ByVal _Functionality As String, ByVal _Module As String, ByVal _VouRef As String, ByRef _RowCnt As Integer, ByRef ErrNo As Integer)
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetUserPermissionFlag]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", _UserName)
            BaseConn.cmd.Parameters.AddWithValue("@Functionality", _Functionality)
            BaseConn.cmd.Parameters.AddWithValue("@Module", _Module)
            BaseConn.cmd.Parameters.AddWithValue("@VouRef", _VouRef)
            BaseConn.cmd.Parameters.Add("@RowCnt", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            _RowCnt = BaseConn.cmd.Parameters("@RowCnt").Value
        Catch ex As Exception
            _ErrString = ex.Message
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub UpdateDynPermissionFlag(ByVal Str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal _UserName As String,
                                      ByVal _Functionality As String, ByVal _Module As String, ByVal _VouRef As String, ByRef ErrNo As Integer)
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_UpdateDynamicPermission]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", _UserName)
            BaseConn.cmd.Parameters.AddWithValue("@Functionality", _Functionality)
            BaseConn.cmd.Parameters.AddWithValue("@Module", _Module)
            BaseConn.cmd.Parameters.AddWithValue("@VouRef", _VouRef)
            BaseConn.cmd.ExecuteNonQuery()

        Catch ex As Exception
            _ErrString = ex.Message
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetProjectAccounts(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _BusinessPeriodID As Integer, ByVal _Project As String, ByRef _Ledger As String, ByVal _intMerchantID As Integer, ByVal _ComboDate As String, ByVal _FromDate As Date, ByVal _ToDate As Date, ByVal _Flag As String, ByRef _dtProjectAccounts As DataTable, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        _ErrNo = 0
        _ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[ProjectExpensesBreakUp]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Project", _Project)
            BaseConn.cmd.Parameters.AddWithValue("@Ledger", _Ledger)
            BaseConn.cmd.Parameters.AddWithValue("@MerchantID", _intMerchantID)
            BaseConn.cmd.Parameters.AddWithValue("@ComboDate", _ComboDate)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            If ds.Tables.Count > 0 Then
                _dtProjectAccounts = ds.Tables(0)
            End If
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetRetentionInvoiceNo(ByVal _strSiteID As String, ByVal _strDBPath As String, ByVal _strDBPwd As String, ByVal _BusinessPeriodID As Integer, ByVal _SISNo As String, ByVal _LedgerID As String, ByVal _ProjectID As String, ByVal _CurrencyCode As String, ByVal _Flag As String, ByRef _dtRetMatching As DataTable, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        _ErrNo = 0
        _ErrStr = ""
        Try
            BaseConn.Open(_strDBPath, _strDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetRetentionInvoiceNos]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _strSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@SISNo", _SISNo)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectID", _ProjectID)
            BaseConn.cmd.Parameters.AddWithValue("@InvCurrency", _CurrencyCode)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            If ds.Tables.Count > 0 Then
                _dtRetMatching = ds.Tables(0)
            End If
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetReport(ByVal _strSiteID As String, ByVal _strDBPath As String, ByVal _strDBPwd As String, ByVal _BusinessPeriodID As Integer, ByVal _GroupID As String, ByVal _MenuID As String, ByRef _dtReport As DataTable, ByRef _ErrNo As Integer, ByRef _ErrStr As String, Optional ByVal _Flag As String = "")
        _ErrNo = 0
        _ErrStr = ""
        Try
            BaseConn.Open(_strDBPath, _strDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetReport]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _strSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@GroupID", _GroupID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", _MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _dtReport = ds.Tables(0)
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetActiveSessionUserID(ByVal _CID As String, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _UserName As String, ByVal _PCName As String, ByVal _LoginTime As Date, ByVal _SessionID As String, ByVal _IP As String, ByRef _AllReadyLogin As Boolean, ByVal _Flag As String, ByRef _ErrNo As Integer, ByRef _ErrString As String)
        _ErrString = ""
        _ErrNo = 0
        Try
            dt = New DataTable
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetActiveSessionUserID]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", _UserName)
            BaseConn.cmd.Parameters.AddWithValue("@PCName", _PCName)
            BaseConn.cmd.Parameters.AddWithValue("@LoginTime", _LoginTime)
            BaseConn.cmd.Parameters.AddWithValue("@SessionID", _SessionID)
            BaseConn.cmd.Parameters.AddWithValue("@IP", _IP)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.Add("@AllreadyLogin", SqlDbType.Bit).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            _AllReadyLogin = BaseConn.cmd.Parameters("@AllreadyLogin").Value
        Catch ex As Exception
            _ErrString = ex.Message
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function ValidateSession(ByVal _SiteID As String, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _UserName As String, ByVal _SessionID As String, ByRef _ErrNo As Integer, ByRef _ErrString As String) As Boolean
        _ErrString = ""
        _ErrNo = 0
        Dim _SessionExpired As Boolean = False
        Try
            dt = New DataTable
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[ValidateSession]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", _UserName)
            BaseConn.cmd.Parameters.AddWithValue("@SessionID", _SessionID)
            BaseConn.cmd.Parameters.Add("@SessionExpired", SqlDbType.Bit).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            _SessionExpired = BaseConn.cmd.Parameters("@SessionExpired").Value
        Catch ex As Exception
            _ErrString = ex.Message
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return _SessionExpired
    End Function

    Public Sub GetDualPwd_Validate(ByVal _SiteID As String, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _LoginUserID As String, ByVal _LoginUserName As String, ByVal _LoginGroupName As String, ByVal _DualUserName As String, ByVal _DualPwd As String, ByVal _FromMenuID As String, ByVal _PCNameandIP As String, ByRef _ErrNo As Integer, ByRef _ErrString As String)
        _ErrNo = 0
        _ErrString = ""
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[DualPwdValidate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@LoginUserID", _LoginUserID)
            BaseConn.cmd.Parameters.AddWithValue("@LoginUserName", _LoginUserName)
            BaseConn.cmd.Parameters.AddWithValue("@LoginGroupName", _LoginGroupName)
            BaseConn.cmd.Parameters.AddWithValue("@DualUserName", _DualUserName)
            BaseConn.cmd.Parameters.AddWithValue("@DualPassword", _DualPwd)
            BaseConn.cmd.Parameters.AddWithValue("@FromMenuID", _FromMenuID)
            BaseConn.cmd.Parameters.AddWithValue("@PCNameandIP", _PCNameandIP)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
        Catch ex As Exception
            _ErrNo = 1
            Dim SPErrString As String = ex.Message.ToString
            If SPErrString = "2" Then
                _ErrString = "Invalid Password"
                _ErrNo = 2
            ElseIf SPErrString = "3" Then
                _ErrString = "Group level low"
                _ErrNo = 3
            Else
                _ErrString = ex.Message
                _ErrNo = 4
            End If
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetVouEnquiryDaysDetails(ByVal _SiteID As String, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _BusinessPeridID As Integer, ByRef _dtEnquiryDetails As DataTable, ByVal _Flag As String, ByRef _ErrNo As Integer, ByRef _ErrString As String)
        _ErrNo = 0
        _ErrString = ""
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetEnquiryDaysDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeridID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _dtEnquiryDetails = ds.Tables(0)
        Catch ex As Exception
            _ErrNo = 1
            _ErrString = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
    ''' <summary>
    ''' Clsing businessperiod
    ''' </summary>
    ''' <param name="_SiteID"></param>
    ''' <param name="_DBPath"></param>
    ''' <param name="_DBPwd"></param>
    ''' <param name="_BusinessPeridID"></param>
    ''' <param name="_UserName"></param>
    ''' <param name="_ErrNo"></param>
    ''' <param name="_ErrString"></param>
    ''' <remarks></remarks>
    Public Sub CloseBunisessPeriod(ByVal _SiteID As String, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _BusinessPeridID As Integer, ByVal _BSPCloseDate As Date, ByVal _UserName As String, ByRef _ErrNo As Integer, ByRef _ErrString As String)
        _ErrNo = 0
        _ErrString = ""
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[BusinessPeriodCutOverLedgers]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeridID)
            BaseConn.cmd.Parameters.AddWithValue("@BSPCloseDate", _BSPCloseDate)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", _UserName)
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output

            BaseConn.cmd.CommandTimeout = 10000
            BaseConn.cmd.ExecuteNonQuery()

            _ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub UpdateClosingEntries(ByVal _SiteID As String, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _BusinessPeridID As Integer, ByVal _BSPCloseDate As Date, ByVal _UserName As String, ByRef _ErrNo As Integer, ByRef _ErrString As String)
        _ErrNo = 0
        _ErrString = ""
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_UpdateClosingEntries]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BPID", _BusinessPeridID)
            BaseConn.cmd.Parameters.AddWithValue("@BSPClosingDate", _BSPCloseDate)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", _UserName)
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output

            BaseConn.cmd.CommandTimeout = 1000
            BaseConn.cmd.ExecuteNonQuery()

            _ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub TrialBalanceMonthWise(ByVal _SiteID As String, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _BusinessPeridID As Integer, ByVal _Date As Date, ByRef _dtPreMonth As DataTable, ByRef _dtSelMonth As DataTable, ByRef _dtComparitive As DataTable, ByRef _dtCumulative As DataTable, ByVal _Flag As String, ByRef _ErrNo As Integer, ByRef _ErrString As String)
        _ErrNo = 0
        _ErrString = ""
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_TrialBalanceMonthWise]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeridID)
            BaseConn.cmd.Parameters.AddWithValue("@Month", _Date)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _dtPreMonth = ds.Tables(0)
            _dtSelMonth = ds.Tables(1)
            _dtComparitive = ds.Tables(2)
            _dtCumulative = ds.Tables(3)
        Catch ex As Exception
            _ErrNo = 1
            _ErrString = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Function Get_ItemProfitability(ByVal SiteID As String, ByVal _strPath As String, ByVal _strDBPwd As String, ByVal _intBusinessPeriodID As Integer, ByVal ItemCodeColl As DataTable, ByVal Description As String, ByVal _FromDate As Date, ByVal _ToDate As Date, ByVal _DateFlag As String, ByVal MerchantLedgerID As String, ByVal InvNo As String, ByVal SalesManID As String, ByVal Branch As String, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        dt = New DataTable
        _ErrStr = ""
        _ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_ItemProfitability]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _intBusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemArray", ItemCodeColl)
            BaseConn.cmd.Parameters.AddWithValue("@ItemName", Description)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@DateFlag", _DateFlag)
            BaseConn.cmd.Parameters.AddWithValue("@MerchantLedgerID", MerchantLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@InvNo", InvNo)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManID", SalesManID)
            BaseConn.cmd.Parameters.AddWithValue("@BranchID", Branch)
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            _ErrStr = ex.Message
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Sub ItemStockByAllCompany(ByVal _SiteID As String, ByVal SiteName As String, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _BusinessPeridID As Integer, ByVal _ItemDT As DataTable, ByRef _ReturnDT As DataTable, ByRef _ErrNo As Integer, ByRef _ErrString As String)
        _ReturnDT = New DataTable
        _ErrString = ""
        _ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[ItemStockByAllCompany]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@SiteName", SiteName)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeridID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemDT", _ItemDT)
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(_ReturnDT)
        Catch ex As Exception
            _ErrString = ex.Message
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Sub StockInCompanys(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _BusinessPeridID As Integer, ByVal _ItemDT As DataTable, ByVal _sitedetailsDT As DataTable, ByRef dt As DataTable, ByRef _ErrNo As Integer, ByRef _ErrString As String)
        dt = New DataTable
        _ErrString = ""
        _ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_StockInCompanys]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeridID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemDT", _ItemDT)
            BaseConn.cmd.Parameters.AddWithValue("@SiteDetailsDT", _sitedetailsDT)
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            _ErrString = ex.Message
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function getProjectwiseStockReport(ByVal str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal int_BusinessPeroidID As Integer,
                                   ByVal ItemCode As String, ByVal dtp_FromDate As Date, ByVal dtp_ToDate As Date, ByRef dbl_Stock As Double,
                                   ByRef dbl_Cost As Double, ByRef CostType As String, ByVal strRPTType As String, ByVal bool_UpdateCost As Boolean,
                                   Optional ByVal ItemCodeColl As DataTable = Nothing, Optional ByVal _WHID As Integer = 0) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            If ItemCodeColl IsNot Nothing Then
                BaseConn.cmd = New SqlClient.SqlCommand("[sp_ProjectwiseStockReport]", BaseConn.cnn)
            End If
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", int_BusinessPeroidID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", dtp_ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Project", strRPTType)
            BaseConn.cmd.Parameters.AddWithValue("@UpdateCost", bool_UpdateCost)
            If ItemCodeColl IsNot Nothing Then
                BaseConn.cmd.Parameters.AddWithValue("@ItemArray", ItemCodeColl)
            End If
            If Not _WHID = 0 Then
                BaseConn.cmd.Parameters.AddWithValue("@WHID", _WHID)
            End If

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

    Public Sub License(ByVal _SiteID As String, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _BusinessPeridID As Integer, ByVal _UserName As String, ByVal _InputString As String, ByRef _ErrNo As Integer, ByRef _ErrString As String)
        _ErrNo = 0
        _ErrString = ""
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[ImportLicense]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeridID)
            BaseConn.cmd.Parameters.AddWithValue("@InputString", _InputString)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", _UserName)
            BaseConn.cmd.ExecuteNonQuery()
        Catch ex As Exception
            _ErrString = ex.Message
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub

    'Public Function getComparitiveMonthlySales(ByVal str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal _StartDay As Integer, _
    '                                           ByVal _EndDay As Integer, ByVal _FirstMonth As Date, ByVal _LastMonth As Date) As DataTable
    '    Try
    '        dt = New DataTable
    '        BaseConn.Open(_strPath, _strPwd)
    '        BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetMonthlyComparitiveSales]", BaseConn.cnn)
    '        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    '        BaseConn.cmd.Parameters.AddWithValue("@SiteID", str_SiteID)
    '        BaseConn.cmd.Parameters.AddWithValue("@StartDay", _StartDay)
    '        BaseConn.cmd.Parameters.AddWithValue("@EndDay", _EndDay)
    '        BaseConn.cmd.Parameters.AddWithValue("@FirstMonth", _FirstMonth)
    '        BaseConn.cmd.Parameters.AddWithValue("@LastMonth", _LastMonth)

    '        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    '        Dim ds As New DataSet
    '        BaseConn.da.Fill(ds)
    '        dt = ds.Tables(0)
    '    Catch ex As Exception
    '        MsgBox("Error" & ex.Message)
    '    Finally
    '        BaseConn.Close()
    '    End Try
    '    Return dt
    'End Function

    Public Function getComparitiveMonthlySales(ByVal _strDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal _Ledger As String,
                                     ByVal _NoofMonths As Integer, ByVal _FromDate As Date, ByVal _ToDate As Date, ByVal _IsSalesMan As Boolean,
                                     ByVal _Day1 As Integer, ByVal _Day2 As Integer) As DataTable
        Try
            BaseConn.Open(_strDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetComparitiveMonthlySales]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _Ledger)

            BaseConn.cmd.Parameters.AddWithValue("@NoOfMonths", _NoofMonths)
            BaseConn.cmd.Parameters.AddWithValue("@StartDate", _FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@EndDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@IsSalesMan", _IsSalesMan)
            BaseConn.cmd.Parameters.AddWithValue("@Day1", _Day1)
            BaseConn.cmd.Parameters.AddWithValue("@Day2", _Day2)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            dt = New DataTable
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Sub FileUpload(ByVal _SiteID As String, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _BusinessPeridID As Integer, ByVal _VouNo As String, ByVal _VouType As String, ByVal _MenuID As String, ByVal _Desc1 As String, ByVal _Desc2 As String, ByVal _FileName As String, ByVal _FileType As String, ByVal _CreatedBy As String, ByVal _Image() As Byte, ByRef _SlNo As String, ByVal _Flag As String, ByRef _ErrNo As Integer, ByRef _ErrString As String)
        _ErrNo = 0
        _ErrString = ""
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[FileUploads]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeridID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", _VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@VouType", _VouType)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", _MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Desc1", _Desc1)
            BaseConn.cmd.Parameters.AddWithValue("@Desc2", _Desc2)
            BaseConn.cmd.Parameters.AddWithValue("@FileName", _FileName)
            BaseConn.cmd.Parameters.AddWithValue("@FileType", _FileType)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", _CreatedBy)

            If _Image Is Nothing Then
                Dim photoParam As New SqlParameter("@Image", SqlDbType.Image)
                photoParam.Value = DBNull.Value
                BaseConn.cmd.Parameters.Add(photoParam)
            Else
                BaseConn.cmd.Parameters.AddWithValue("@Image", DirectCast(_Image, Object))
            End If


            BaseConn.cmd.Parameters.AddWithValue("@SlNo", _SlNo)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)

            BaseConn.cmd.Parameters.AddWithValue("@SlNoOut", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 200).Direction = ParameterDirection.Output

            BaseConn.cmd.ExecuteNonQuery()

            _SlNo = BaseConn.cmd.Parameters("@SlNoOut").Value.ToString
            _ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetAccuredLiabilityAndPurchaseExpenseLedger(ByVal _strPath As String, ByVal _strPwd As String, ByRef _AccuredLiability As Integer,
                                    ByRef _PurchaseExpense As Integer)
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetAccuredLiabilityAndPurchaseExpenseLedger]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", CID)
            BaseConn.cmd.Parameters.Add("@AccuredLiability", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@PurchaseExpense", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            _AccuredLiability = BaseConn.cmd.Parameters("@AccuredLiability").Value
            _PurchaseExpense = BaseConn.cmd.Parameters("@PurchaseExpense").Value
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function GetLedgerDesc(ByVal _SiteID As String, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _LedgerID As Integer,
                          ByRef _ErrNo As Integer, ByRef _ErrString As String) As String
        GetLedgerDesc = String.Empty
        _ErrNo = 0
        _ErrString = ""
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetLedgerDesc]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.Add("@LedgerDesc", SqlDbType.NVarChar, 250).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            GetLedgerDesc = BaseConn.cmd.Parameters("@LedgerDesc").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return GetLedgerDesc
    End Function

    Public Sub Get_DefaultLedger(ByVal _SiteID As String, ByVal StrDBPath As String, ByVal StrDBPwd As String, ByRef _DTBaseDropDown As DataTable, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(StrDBPath, StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetDefaultLedger]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", "DefaultLedger")
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _DTBaseDropDown = ds.Tables(0)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetPLAndPCC(ByRef _DBPath As String, ByRef _DBPwd As String, ByRef _CID As String, ByRef _VouNo As String, ByRef _VouType As String,
                                ByRef _DTTxnLedgers As DataTable, ByRef _DTCostCentre As DataTable, ByRef _ErrNo As Integer, ByRef _ErrStr As String)

        _ErrNo = 0
        _ErrStr = ""
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetTxnLedgers]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", _VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@VouType", _VouType)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _DTTxnLedgers = ds.Tables(0)
            _DTCostCentre = ds.Tables(1)
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try

    End Sub
    Public Sub GetTxnLedgers(ByVal _SiteID As String, ByVal StrDBPath As String, ByVal StrDBPwd As String, ByVal _VouNo As String, ByVal _VouType As String,
                                ByRef _DTTxnLedgers As DataTable, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(StrDBPath, StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetTxnLedgers]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", _VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@Type", _VouType)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _DTTxnLedgers = ds.Tables(0)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function GetPurchaseExpenseLedger(ByVal _strPath As String, ByVal _strPwd As String) As Integer
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetPurchaseExpenseLedger]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", CID)
            BaseConn.cmd.Parameters.Add("@PurchaseExpense", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            GetPurchaseExpenseLedger = BaseConn.cmd.Parameters("@PurchaseExpense").Value
            Return GetPurchaseExpenseLedger
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Function


    Public Sub Load_ItemBatch(ByVal _strPath As String, ByVal _strpwd As String, ByVal _WHID As Integer, ByVal _ItemCode As String,
                               ByVal _VouRef As String, ByVal _Type As String, ByRef _DTItemBatch As DataTable,
                             ByRef _DTItemBin As DataTable, ByRef _DTItemSerial As DataTable, ByRef _BatchFlag As Boolean, ByRef _BinFlag As Boolean, ByRef _SerialFlag As Boolean, ByVal _Flag As String)
        Try
            BaseConn.Open(_strPath, _strpwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[LoadItemBatch]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@WHID", _WHID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", _ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@VouRef", _VouRef)
            BaseConn.cmd.Parameters.AddWithValue("@Type", _Type)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)

            BaseConn.cmd.Parameters.Add("@BatchFlag", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@BinFlag", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@SerialFlag", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _DTItemBatch = ds.Tables(0)
            _DTItemBin = ds.Tables(1)
            _DTItemSerial = ds.Tables(2)

            _BatchFlag = BaseConn.cmd.Parameters("@BatchFlag").Value
            _BinFlag = BaseConn.cmd.Parameters("@BinFlag").Value
            _SerialFlag = BaseConn.cmd.Parameters("@SerialFlag").Value
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub LoadItemFromBin(ByVal _strPath As String, ByVal _strpwd As String, ByVal _WHID As Integer, ByRef _DTItemFromBin As DataTable)
        Try
            BaseConn.Open(_strPath, _strpwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_LoadItemFromBin]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@WHID", _WHID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            _DTItemFromBin = ds.Tables(0)

        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function GetLedgerBalanceSTD(ByVal Str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String,
                                  ByVal dtp_from As Date, ByVal dtp_to As Date, ByVal _ReportLevel As String, ByVal _ZeroSuppress As Boolean,
                                  ByVal _ShowInActive As Boolean, ByVal _Type As Integer, ByVal _LedgerID As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_TrialBalance_STD]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", dtp_from)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", dtp_to)
            BaseConn.cmd.Parameters.AddWithValue("@ReportLevel", _ReportLevel)
            BaseConn.cmd.Parameters.AddWithValue("@ZeroSuppress", _ZeroSuppress)
            BaseConn.cmd.Parameters.AddWithValue("@ShowInActive", _ShowInActive)
            BaseConn.cmd.Parameters.AddWithValue("@Type", _Type)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.CommandTimeout = 1000
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

    Public Sub GetSingleLedgerBalance(ByVal _SiteID As String, ByVal StrDBPath As String, ByVal StrDBPwd As String, ByVal _LedgerID As String, ByVal _ToDate As Date,
                             ByRef _BalanceAmount As Double, ByRef ErrNo As Integer, ByRef ErrStr As String)
        _BalanceAmount = 0
        ErrNo = 0
        ErrStr = ""

        Try
            BaseConn.Open(StrDBPath, StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetSingleLedgerBalance]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.Add("@Balance", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            _BalanceAmount = BaseConn.cmd.Parameters("@Balance").Value
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function GetPurchasePrice(ByVal _strPath As String, ByVal _strPwd As String, ByVal _ItemCode As String,
                                        ByVal _FormType As String) As Double
        Try
            Dim _Price As Double = 0
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetPurchasePrice]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", _ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@FormType", _FormType)
            BaseConn.cmd.Parameters.Add("@Price", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            _Price = BaseConn.cmd.Parameters("@Price").Value
            Return _Price
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Function

    Public Function GetPDCAmount(ByVal _strPath As String, ByVal _strPwd As String, ByVal _LedgerID As Integer) As Decimal
        Try
            GetPDCAmount = 0
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetPDCAmount]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.Add("@PDCAmount", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            GetPDCAmount = BaseConn.cmd.Parameters("@PDCAmount").Value
            Return GetPDCAmount
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Function

    Public Function GetLeastCreditDate(ByVal _strPath As String, ByVal _strPwd As String, ByVal _LedgerID As Integer, ByVal _IncludePDC As Boolean) As Integer
        Try
            Dim NoOfCreditDays As Integer = 0
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetLeastCreditDate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@IncludePDC", _IncludePDC)
            BaseConn.cmd.Parameters.Add("@NoOfCreditDays", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            NoOfCreditDays = BaseConn.cmd.Parameters("@NoOfCreditDays").Value
            Return NoOfCreditDays
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Function

    Public Sub GetBarCode(ByVal _strPath As String, ByVal _strpwd As String, ByVal _ItemCode As String, ByRef _DTBarCode As DataTable)
        Try
            BaseConn.Open(_strPath, _strpwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetBarCode]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", _ItemCode)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            _DTBarCode = ds.Tables(0)

        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetSalaryForEmployeeGratity(ByVal _strPath As String, ByVal _strPwd As String, ByVal _LedgerID As Integer, ByVal _SalaryMonth As Date, ByRef _NetAmount As Decimal)
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetSalaryforEmployeeGratuity]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@SalaryMonth", _SalaryMonth)
            BaseConn.cmd.Parameters.Add("@NetAmountOUT", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            _NetAmount = BaseConn.cmd.Parameters("@NetAmountOUT").Value
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetNegativeStockItems(ByVal _SiteID As String, ByVal StrDBPath As String, ByVal StrDBPwd As String, ByVal _VouNo As String, ByVal _Flag As String,
                             ByRef _DTNegativeStockItems As DataTable, ByVal ItemDT As DataTable)
        _DTNegativeStockItems = New DataTable
        Try
            BaseConn.Open(StrDBPath, StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetNegativeStockItems]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", _VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@ItemDT", ItemDT)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _DTNegativeStockItems = ds.Tables(0)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Sub SalesManWiseProfit(ByVal _SiteID As String, ByVal SalesMan As String, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _BusinessPeridID As Integer, ByVal Date_ As String, ByVal FromDate As Date, ByVal ToDate As Date, ByVal _ItemDT As DataTable, ByRef _ReturnDT As DataTable, ByRef _ErrNo As Integer, ByRef _ErrString As String)
        _ReturnDT = New DataTable
        _ErrString = ""
        _ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_SalesManWiseProfit]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@SalesMan", SalesMan)
            BaseConn.cmd.Parameters.AddWithValue("@ItemDT", _ItemDT)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeridID)
            BaseConn.cmd.Parameters.AddWithValue("@Date_", Date_)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", ToDate)

            BaseConn.cmd.CommandTimeout = 500
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(_ReturnDT)
        Catch ex As Exception
            _ErrString = ex.Message
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Sub Get_ChequePrint(ByVal _strPath As String, ByVal _strpwd As String, ByVal _BusinessPeridID As Integer, ByRef _DT As DataTable)
        Try
            BaseConn.Open(_strPath, _strpwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetChequePrintDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeridID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", "")
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _DT = ds.Tables(0)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub UpdateItemImage(ByVal _SiteID As String, ByVal _strPath As String, ByVal _strpwd As String, ByVal _ItemCode As String, ByVal _Image() As Byte, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        _ErrStr = ""
        _ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strpwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_UpdateItemImage]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", _ItemCode)
            If _Image Is Nothing Then
                Dim photoParam As New SqlParameter("@Image", SqlDbType.Image)
                photoParam.Value = DBNull.Value
                BaseConn.cmd.Parameters.Add(photoParam)
            Else
                BaseConn.cmd.Parameters.AddWithValue("@Image", DirectCast(_Image, Object))
            End If
            BaseConn.cmd.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Sub Get_EOSLeaveSalary(ByVal _SiteID As String, ByVal _strPath As String, ByVal _strpwd As String, ByRef _dt As DataTable, ByVal _LedgerID As Integer, ByVal _EmployeeType As String, ByVal _MenuID As String, ByVal _Flag As String, ByVal Date_ As String, ByVal FromDate As Date, ByVal ToDate As Date, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        Try
            BaseConn.Open(_strPath, _strpwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_EOSLeaveSalaryLedger]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", 101)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", _MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@EmployeeType", _EmployeeType)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@Date1", Date_)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", ToDate)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _dt = ds.Tables(0)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function LoadMerchantBasedOnSalesMan(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _SiteID As String, ByVal _TableName As String, ByVal _DT As DataTable,
                               Optional ByVal _FormType As String = "", Optional ByVal _Condition As String = Nothing, Optional ByVal _MenuID As String = Nothing) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[LoadMerchantBasedOnSalesMan]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@TableName", _TableName.ToUpper)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManDT", _DT)
            BaseConn.cmd.Parameters.AddWithValue("@Type", _FormType)
            BaseConn.cmd.Parameters.AddWithValue("@Condition", _Condition)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", _MenuID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Sub GetFormApprovalSettings(ByVal _strPath As String, ByVal _strPwd As String, ByVal _SiteID As String, ByVal _MenuID As String,
                                      ByVal _UserName As String, ByVal _GroupID As Integer, ByRef _ApproveLevel As Integer, ByRef _ApproveHigherLevel As Boolean)
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetFormApprovalSettings]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", _MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", _UserName)
            BaseConn.cmd.Parameters.AddWithValue("@GroupID", _GroupID)

            BaseConn.cmd.Parameters.Add("@ApprovedLevel", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ApprovedHigherLevel", SqlDbType.Bit).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            _ApproveLevel = BaseConn.cmd.Parameters("@ApprovedLevel").Value
            _ApproveHigherLevel = BaseConn.cmd.Parameters("@ApprovedHigherLevel").Value
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function GetFormApprovalDetails(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _MenuID As String, ByVal _SiteID As String,
                                           ByVal _VouNo As String, ByVal _Flag As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetFormApprovalDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", _MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", _VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Sub getStockReportExcel(ByVal str_SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal int_BusinessPeroidID As Integer,
                                   ByVal ItemCode As String, ByVal dtp_FromDate As Date, ByVal dtp_ToDate As Date, ByRef dbl_Stock As Double,
                                   ByRef dbl_Cost As Double, ByRef CostType As String, ByVal strRPTType As String, ByVal bool_UpdateCost As Boolean,
                                   Optional ByVal ItemCodeColl As DataTable = Nothing, Optional ByVal _WHID As Integer = 0, Optional ByRef _dt_Excelmain As DataTable = Nothing, Optional ByRef _dt_excelSub As DataTable = Nothing)

        Try

            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            'If ItemCodeColl IsNot Nothing Then
            BaseConn.cmd = New SqlClient.SqlCommand("[GetStockReport4CollExcel]", BaseConn.cnn)
            'Else
            '    BaseConn.cmd = New SqlClient.SqlCommand("[sp_getStockReport]", BaseConn.cnn)
            'End If

            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", int_BusinessPeroidID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", dtp_FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", dtp_ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@RptType", strRPTType)
            BaseConn.cmd.Parameters.AddWithValue("@UpdateCost", bool_UpdateCost)
            If ItemCodeColl IsNot Nothing Then
                BaseConn.cmd.Parameters.AddWithValue("@ItemArray", ItemCodeColl)
            End If
            If Not _WHID = 0 Then
                BaseConn.cmd.Parameters.AddWithValue("@WHID", _WHID)
            End If


            BaseConn.cmd.Parameters.Add("@CalcWAC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@Stock", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@CostType", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output


            BaseConn.cmd.CommandTimeout = 1000
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            _dt_Excelmain = ds.Tables(0)
            dbl_Cost = BaseConn.cmd.Parameters("@CalcWAC").Value.ToString
            dbl_Stock = BaseConn.cmd.Parameters("@Stock").Value.ToString
            CostType = BaseConn.cmd.Parameters("@CostType").Value.ToString
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        'Return dt
    End Sub

    Public Function GetEmployeeSalaryDetails(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _SiteID As String, ByVal _Flag As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_EmployeeSalaryDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function Get_AttendanceWithOT(ByVal SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal LedgerID_ As Integer, ByVal PSMonth_ As Date, ByVal Flag_ As String, ByRef _ErrNo As Integer, ByRef _ErrStr As String) As DataTable
        _ErrNo = 0
        _ErrStr = ""
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetAttendanceWithOT]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", LedgerID_)
            BaseConn.cmd.Parameters.AddWithValue("@PSMonth", PSMonth_)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Flag_)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message.ToString
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function


    Public Function GetMissMatchedItems(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _SiteID As String, ByVal GivenItems As DataTable, ByRef ErrNo As Integer) As DataTable
        Dim _ErrString As String = ""
        ErrNo = 0
        dt = New DataTable
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetMismatchedItems]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@ReceivedItemDT", GivenItems)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = ds.Tables(0)
        Catch ex As Exception
            _ErrString = ex.Message
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Sub GetDefaultSalesman(ByVal _strPath As String, ByVal _strPwd As String, ByVal _SiteID As Integer, ByVal _LedgerID As Integer,
                                      ByRef _DefaultSalesman As Integer)
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetDefaultSalesman]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)

            BaseConn.cmd.Parameters.Add("@SalesmanID", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            _DefaultSalesman = BaseConn.cmd.Parameters("@SalesmanID").Value
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetItemCostMinMaxPrice(ByVal _SiteID As Integer, ByVal _strPath As String, ByVal _strPwd As String,
                                   ByVal _ItemCode As String, ByRef _CostPrice As Decimal, ByRef _MinSellPrice As Decimal, ByRef _MaxPurPrice As Decimal)
        Try
            'dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetItemCostMinMaxPrice]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", _ItemCode)

            BaseConn.cmd.Parameters.Add("@CostPrice", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@MinSellPrice", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@MaxPurPrice", SqlDbType.Float).Direction = ParameterDirection.Output

            BaseConn.cmd.ExecuteNonQuery()
            _CostPrice = BaseConn.cmd.Parameters("@CostPrice").Value
            _MinSellPrice = BaseConn.cmd.Parameters("@MinSellPrice").Value
            _MaxPurPrice = BaseConn.cmd.Parameters("@MaxPurPrice").Value
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try

    End Sub

    Public Sub ImportRVPVfromExcel(ByVal _strPath As String, ByVal _strPwd As String, ByVal _SiteID As String, ByVal _BSID As Integer,
                                      ByVal _RVPVDT As DataTable, ByVal _CreatedBy As String, ByRef ErrNo As Integer, ByRef _ErrDesc As String)
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("SP_ImportRVPVfromExcel", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BSID)
            BaseConn.cmd.Parameters.AddWithValue("@RVPVDT", _RVPVDT)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", _CreatedBy)

            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 2000
            BaseConn.cmd.ExecuteNonQuery()

            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            _ErrDesc = _ErrString
        Catch ex As Exception
            _ErrString = ex.Message
            'ObjDalGeneral = New DAL_General(_SiteID)
            Elog_Insert(_SiteID, _strPath, _strPwd, _BSID, _CreatedBy, Date.Now, "", "RVPVImport", Err.Number, "Error in Import from Excel :", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

    End Sub
    Public Function GetTaxLedgerSummary(ByVal _SiteID As String, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _FromDate As Date,
                                    ByVal _ToDate As Date, ByVal _Flag As String, ByVal _TRNNo As String, ByVal _Area As String, ByVal _SalesType As String,
                                    ByVal _GroupbyTRN As Boolean, ByVal _GroupbyArea As Boolean, ByVal _GroupbySalesType As Boolean, ByVal _MerchantName As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetTaxLedgerSummary]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@GroupbyTRN", _GroupbyTRN)
            BaseConn.cmd.Parameters.AddWithValue("@GroupbyArea", _GroupbyArea)
            BaseConn.cmd.Parameters.AddWithValue("@GroupbySalesType", _GroupbySalesType)
            BaseConn.cmd.Parameters.AddWithValue("@TRNNo", _TRNNo)
            BaseConn.cmd.Parameters.AddWithValue("@Area", _Area)
            BaseConn.cmd.Parameters.AddWithValue("@SalesType", _SalesType)
            BaseConn.cmd.Parameters.AddWithValue("@MerchantName", _MerchantName)
            BaseConn.cmd.CommandTimeout = 2000

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
    Public Sub GetDailyReport(ByVal _SiteID As String, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _FromDate As Date, ByVal _ToDate As Date, ByVal _DebtorsOS As Double, ByVal _Flag As String, ByRef dtMain As DataTable, ByRef dt_Sal As DataTable, ByRef dt_Pur As DataTable, ByRef dt_PettyCash As DataTable, ByRef dt_TaxAmount As DataTable)
        Try
            dt = New DataTable
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetDailyReports]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@DebtorsOS", _DebtorsOS)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            If CID = "101" Then
                dtMain = ds.Tables(0)
                dt_Sal = ds.Tables(1)
            Else
                dtMain = ds.Tables(0)
                dt_Sal = ds.Tables(1)
                dt_Pur = ds.Tables(2)
                dt_PettyCash = ds.Tables(3)
                dt_TaxAmount = ds.Tables(4)
            End If

        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Sub GetBatchprofitReport(ByVal _SiteID As String, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _Flag As String, ByVal VouNo As String, ByRef dtMain As DataTable, ByRef dt_Exp As DataTable)
        Try
            dt = New DataTable
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetBatchProfitReport]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dtMain = ds.Tables(0)
            dt_Exp = ds.Tables(1)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Function GetItemMasterDetails(ByVal _strPath As String, ByVal _strPwd As String, ByVal SiteID As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetItemMasterDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", SiteID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
            BaseConn.da.Dispose()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function GetStockMovementSpecific(ByVal _strPath As String, ByVal _strPwd As String, ByVal _SiteID As String, ByVal _BSID As Integer, ByVal _FromDate As Date, ByVal _ToDate As Date, ByVal _DT As DataTable, ByRef _ErrNo As Integer) As DataTable
        Try
            _ErrNo = 0
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_StockMovementSpecific]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BSID)
            BaseConn.cmd.Parameters.AddWithValue("@DT", _DT)
            BaseConn.cmd.CommandTimeout = 2000
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
    Public Function Get_LedgerStatementExcelExport(ByVal _strPath As String, ByVal _strpwd As String, ByVal _SiteID As String, ByVal _Ledger As Integer,
                                                   ByVal _FrmDate As Date, ByVal _ToDate As Date, ByVal _Condition As String, ByVal _LedgerType As String, ByVal _IncludePdc As Boolean) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strpwd)

            BaseConn.cmd = New SqlClient.SqlCommand("[sp_LedgerStatementDetailedExcelExport]", BaseConn.cnn)

            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Ledger", _Ledger)
            BaseConn.cmd.Parameters.AddWithValue("@FrmDate", _FrmDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            'BaseConn.cmd.Parameters.AddWithValue("@Condition", _Condition.ToUpper)
            'BaseConn.cmd.Parameters.AddWithValue("@LedgerType", _LedgerType.ToUpper)
            'BaseConn.cmd.Parameters.AddWithValue("@PDC", _IncludePdc)
            'BaseConn.cmd.Parameters.AddWithValue("@BSPeriod", _BusPeriod)
            'BaseConn.cmd.Parameters.AddWithValue("@BusStartDate", _BusStartDate)
            BaseConn.cmd.CommandTimeout = 1000

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function Get_ProductMasterView(ByVal SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal _ProductCode As String, ByVal _ItemCode As String, ByVal _ZeroSupress As Boolean, ByVal _Active As Boolean, ByVal _InActive As Boolean, ByVal _Flag As String, ByRef _ErrNo As Integer, ByRef _ErrStr As String) As DataTable
        Try
            _ErrNo = 0
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetProductMasterView]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@ProductCode", _ProductCode)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", _ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@ZeroSupress", _ZeroSupress)
            BaseConn.cmd.Parameters.AddWithValue("@Active", _Active)
            BaseConn.cmd.Parameters.AddWithValue("@InActive", _InActive)
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function Get_ItemDetails(ByVal SiteID As String, ByVal _strPath As String, ByVal _strPwd As String, ByVal _VoucherDT As DataTable, ByVal _MenuID As String, ByVal _Flag As String, ByRef _ErrNo As Integer, ByRef _ErrStr As String) As DataTable
        Try
            _ErrNo = 0
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetItemDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@DT", _VoucherDT)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", _MenuID)
            BaseConn.cmd.CommandTimeout = 2000
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function GetTaxGroupWithTaxCode(ByVal _strPath As String, ByVal _strPwd As String, ByVal SiteID As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetTaxGroupWithTaxCode]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", SiteID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
            BaseConn.da.Dispose()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function GetTaxForInventory(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _SiteID As String, ByVal _Flag As String, ByVal _DTVouNo As DataTable) As DataTable
        GetTaxForInventory = New DataTable

        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetTaxForInvoice]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@InventoryDT", _DTVouNo)

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            GetTaxForInventory = ds.Tables(0)
            Return GetTaxForInventory
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Function

    Public Sub UpdateCostItems(ByVal _SiteID As String, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _FromDate As Date, ByVal _BusinessPeriod As String, ByRef dt_ItemCol As DataTable)
        Try
            dt = New DataTable
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[UpdateCost]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriod)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCodeColl", dt_ItemCol)
            BaseConn.cmd.Parameters.AddWithValue("@UpdateCost", "")
            BaseConn.cmd.CommandTimeout = 2000
            BaseConn.cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function GetTaxDetailsGroupByHSN(ByVal _strPath As String, ByVal _strPwd As String, ByVal SiteID As String, ByVal FormName As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetTaxDetailsGroupByHSN]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@FormName", FormName)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
            BaseConn.da.Dispose()
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
    Public Function GetMerchantOutstanding(ByVal _SiteID As String, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _FromDate As Date,
                                    ByVal _ToDate As Date, ByVal dt_ItemArray As DataTable, ByVal _Flag As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetMerchantOutstandingDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@DT", dt_ItemArray)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
    Public Function GetStockMovementAnalysisReport(ByVal _SiteID As String, ByVal _BSID As Integer, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _FromDate As Date,
                                    ByVal _ToDate As Date, ByVal dt_ItemArray As DataTable) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[SP_GetStockMovementAnalysisReport]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BSID)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@DT", dt_ItemArray)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.cmd.CommandTimeout = 2000
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function GetDODetailsForInvoiceNo(ByVal _strPath As String, ByVal _strpwd As String, ByVal _SiteID As String, ByRef _VouNo As String, ByVal _LedgerID As Integer,
                                               ByRef _Currency As String) As DataSet
        Dim dset As New DataSet
        Try

            BaseConn.Open(_strPath, _strpwd)

            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetDODetailsForInvoiceNo]", BaseConn.cnn)

            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", _VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Currency", _Currency)
            BaseConn.cmd.CommandTimeout = 1000

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dset)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try

        Return dset
    End Function

    Public Function ClassToJSon(ByRef obj As Object) As String
        ClassToJSon = String.Empty
        Dim dic As New Dictionary(Of String, String)

        For Each f As Reflection.FieldInfo In obj.GetType.GetFields
            If f.FieldType IsNot GetType(DataTable) Then
                dic.Add(f.Name, f.GetValue(obj))
            End If
        Next
        Dim serializer As New JavaScriptSerializer
        ClassToJSon = serializer.Serialize(dic)

        Return ClassToJSon
    End Function

    Public Function DatatableToJSONString(ByRef _DT As DataTable) As String
        DatatableToJSONString = String.Empty
        DatatableToJSONString = JsonConvert.SerializeObject(_DT)
        Return DatatableToJSONString
    End Function

    Public Function GetDataTableFromJsonString(json As String) As DataTable
        Dim table As DataTable = JsonConvert.DeserializeObject(Of DataTable)(json)
        'table.Dump()
        Return table
    End Function

    Public Sub GetNextPreviousVoucherNo(_CID As Integer, _strDBPath As String, _strDBPwd As String, _MenuID As String, ByVal _Prefix As String, ByVal _VoucherNo As String, _Flag As String, ByRef _ErrNo As Integer, ByRef _RtnVouNo As String)
        Dim _ErrString As String = ""
        _RtnVouNo = ""
        _ErrNo = 0
        Try
            BaseConn.Open(_strDBPath, _strDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("GetNextPreviousVoucherNo", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", _MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Prefix", _Prefix)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", _VoucherNo)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            'BaseConn.cmd.Parameters.AddWithValue("@RtnVouNo", SqlDbType, VarChar(30)).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@RtnVouNo", SqlDbType.VarChar, 30).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            _RtnVouNo = BaseConn.cmd.Parameters("@RtnVouNo").Value.ToString
        Catch ex As Exception
            _ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub UpdateNotification(ByVal _CID As Integer, ByVal _strDBPath As String, ByVal _strDBPwd As String, ByVal _VouNo As String, ByVal _MenuID As String, ByVal _RevNo As Integer, ByVal _UserName As String, ByVal _Flag As String, ByVal _LedgerID As Integer, ByVal _SMSSuccess As Boolean, ByVal _EmailSuccess As Boolean, ByRef _ErrNo As Integer)
        _ErrNo = 0
        Try
            BaseConn.Open(_strDBPath, _strDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[UpdateNotification]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", _VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", _MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", _RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", _UserName)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@MerLedgerID", _LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@SMSSuccess", _SMSSuccess)
            BaseConn.cmd.Parameters.AddWithValue("@EmailSuccess", _EmailSuccess)

            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
        Catch ex As Exception
            _ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function GetConversionVouchers(ByVal _strPath As String, ByVal _strPwd As String, ByVal _CID As String, ByVal _Flag As String, _ApprovedStatus As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetConversionVouchers]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", _ApprovedStatus)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function GetConversionVouchersItems(ByVal _strPath As String, ByVal _strPwd As String, ByVal _CID As String, ByVal _Flag As String, _VouNo As String) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetConversionVouchersItems]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", _VouNo)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Sub Update_LableName(ByVal _int_CID As Integer, ByVal _DBPath As String, ByVal _DMPwd As String, ByVal _LanguageCode As Integer, ByVal _MenuID As String, ByVal _ControlName As String, ByVal _ControlText As String, ByVal _UserName As String, ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        _ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DMPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[UpdateLabelName]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", _MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@LanguageCode", _LanguageCode)
            BaseConn.cmd.Parameters.AddWithValue("@ControlName", _ControlName)
            BaseConn.cmd.Parameters.AddWithValue("@ControlText", _ControlText)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", _UserName)
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub


    Public Sub PCCUpdate(ByRef _strDBPath As String, ByRef _strDBPwd As String, ByRef _CID As Integer, ByRef _BusinessPeriodID As Integer, ByRef _VouNo As String,
                        ByRef _VouType As String, ByRef _CreatedBy As String, ByRef _DTPCC As DataTable, ByRef _OutSMS As String, ByRef _OutEmail As String, ByRef _ErrNo As Integer)
        Dim _ErrString As String = ""
        _ErrNo = 0
        Try
            BaseConn.Open(_strDBPath, _strDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[PCCUpdate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", _VouNo)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@VouType", _VouType)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", _CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@DTPCC", _DTPCC)

            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutSMSMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutEmailMsgText", SqlDbType.NVarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            _ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            _OutSMS = BaseConn.cmd.Parameters("@OutSMSMsgText").Value.ToString
            _OutEmail = BaseConn.cmd.Parameters("@OutEmailMsgText").Value.ToString
        Catch ex As Exception
            _ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetFinancialStatement(ByVal _strPath As String, ByVal _strpwd As String, ByVal _CID As String, ByVal _FrmDate As Date, ByVal _ToDate As Date,
                                          ByVal _ReportID As Integer, ByVal _ShowBeginingBalance As Boolean, ByVal _ShowEndingBalance As Boolean, ByVal _ShowDrCr As Boolean,
                                          ByVal _ShowQuaterly As Boolean, ByVal _ShowMonthly As Boolean, ByVal _ShowGroupTotalLine As Boolean, ByVal _PostClose As Integer,
                                          ByVal _ShowCostCentre As Boolean, ByVal _ShowGroupOnly As Boolean, ByRef _DTResult As DataTable, ByRef _DTFormatting As DataTable)
        Try
            Dim ds As New DataSet
            BaseConn.Open(_strPath, _strpwd)

            BaseConn.cmd = New SqlClient.SqlCommand("[GetFinancialReport]", BaseConn.cnn)

            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FrmDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@ReportID", _ReportID)
            BaseConn.cmd.Parameters.AddWithValue("@ShowBeginingBalance", _ShowBeginingBalance)
            BaseConn.cmd.Parameters.AddWithValue("@ShowEndingBalance", _ShowEndingBalance)
            BaseConn.cmd.Parameters.AddWithValue("@ShowDrCr", _ShowDrCr)
            BaseConn.cmd.Parameters.AddWithValue("@ShowQuaterly", _ShowQuaterly)
            BaseConn.cmd.Parameters.AddWithValue("@ShowMonthly", _ShowMonthly)
            BaseConn.cmd.Parameters.AddWithValue("@ShowGroupTotalLine", _ShowGroupTotalLine)
            BaseConn.cmd.Parameters.AddWithValue("@PostClose", _PostClose)
            BaseConn.cmd.Parameters.AddWithValue("@ShowCostCentre", _ShowCostCentre)
            BaseConn.cmd.Parameters.AddWithValue("@ShowGroupOnly", _ShowGroupOnly)
            BaseConn.cmd.CommandTimeout = 1000

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(ds)

            _DTResult = ds.Tables(0)
            _DTFormatting = ds.Tables(1)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetFinancialStatement_New(ByVal _strPath As String, ByVal _strpwd As String, ByVal _CID As String, ByVal _FrmDate As Date, ByVal _ToDate As Date,
                                          ByVal _ReportID As Integer, ByVal _ShowBeginingBalance As Boolean, ByVal _ShowEndingBalance As Boolean, ByVal _ShowDrCr As Boolean,
                                          ByVal _ShowQuaterly As Boolean, ByVal _ShowMonthly As Boolean, _ShowYearly As Boolean, ByVal _ShowGroupTotalLine As Boolean, _ShowActiveOnly As Boolean, ByVal _PostClose As Integer,
                                          ByVal _ShowCostCentre As Boolean, ByVal _ShowGroupOnly As Boolean, ByRef _DTResult As DataTable, ByRef _DTFormatting As DataTable)
        Try
            Dim ds As New DataSet
            BaseConn.Open(_strPath, _strpwd)

            BaseConn.cmd = New SqlClient.SqlCommand("[GetFinancialReport_New]", BaseConn.cnn)

            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FrmDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@ReportID", _ReportID)
            BaseConn.cmd.Parameters.AddWithValue("@ShowBeginingBalance", _ShowBeginingBalance)
            BaseConn.cmd.Parameters.AddWithValue("@ShowEndingBalance", _ShowEndingBalance)
            BaseConn.cmd.Parameters.AddWithValue("@ShowDrCr", _ShowDrCr)
            BaseConn.cmd.Parameters.AddWithValue("@ShowQuaterly", _ShowQuaterly)
            BaseConn.cmd.Parameters.AddWithValue("@ShowMonthly", _ShowMonthly)
            BaseConn.cmd.Parameters.AddWithValue("@ShowYearly", _ShowYearly)
            BaseConn.cmd.Parameters.AddWithValue("@ShowGroupTotalLine", _ShowGroupTotalLine)
            BaseConn.cmd.Parameters.AddWithValue("@ShowActiveOnly", _ShowActiveOnly)
            BaseConn.cmd.Parameters.AddWithValue("@PostClose", _PostClose)
            BaseConn.cmd.Parameters.AddWithValue("@ShowCostCentre", _ShowCostCentre)
            BaseConn.cmd.Parameters.AddWithValue("@ShowGroupOnly", _ShowGroupOnly)
            BaseConn.cmd.CommandTimeout = 1000

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(ds)

            _DTResult = ds.Tables(0)
            _DTFormatting = ds.Tables(1)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function GetCostCentreStatement(ByVal _strPath As String, ByVal _strpwd As String, ByVal _CID As String, ByVal _FrmDate As Date, ByVal _ToDate As Date, ByVal _CostCentreGroupID As Integer,
                                         ByVal _CostCentreID As Integer, ByVal _ShowDrCr As Boolean, ByVal _ShowQuaterly As Boolean, ByVal _ShowMonthly As Boolean) As DataTable
        Try
            GetCostCentreStatement = New DataTable
            Dim ds As New DataSet

            BaseConn.Open(_strPath, _strpwd)

            BaseConn.cmd = New SqlClient.SqlCommand("[GetCostCentreReport]", BaseConn.cnn)

            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FrmDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@CostCentreGroupID", _CostCentreGroupID)
            BaseConn.cmd.Parameters.AddWithValue("@CostCentreID", _CostCentreID)
            BaseConn.cmd.Parameters.AddWithValue("@ShowDrCr", _ShowDrCr)
            BaseConn.cmd.Parameters.AddWithValue("@ShowQuaterly", _ShowQuaterly)
            BaseConn.cmd.Parameters.AddWithValue("@ShowMonthly", _ShowMonthly)
            BaseConn.cmd.CommandTimeout = 1000

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(GetCostCentreStatement)
            'BaseConn.da.Fill(ds)
            'dt = ds.Tables(0)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return GetCostCentreStatement
    End Function

    Public Function ProjectGeneralReport(ByVal _SiteID As String, ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _ProjectID As String, ByVal _EmpLedgerID As String, ByVal _LocationID As String, ByVal _MerchantID As String, ByVal _DateType As String, ByVal _FromDate As Date,
                                  ByVal _ToDate As Date, ByVal _Flag As String, ByRef _ErrNo As Integer, ByRef _ErrStr As String, ByVal dt_ItemArray As DataTable, ByVal _ShowAll As Boolean, ByVal _Completed As Boolean, ByVal _NotCompleted As Boolean, ByVal _Invoice As Boolean, ByVal _UnInvoice As Boolean) As DataTable
        _ErrNo = 0
        _ErrStr = ""
        Try
            dt = New DataTable
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[FE_GetProjectGeneralReport]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@ProjectID", _ProjectID)
            BaseConn.cmd.Parameters.AddWithValue("@EmpLedgerID", _EmpLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@LocationID", _LocationID)
            BaseConn.cmd.Parameters.AddWithValue("@MerchantID", _MerchantID)
            BaseConn.cmd.Parameters.AddWithValue("@DateType", _DateType)
            BaseConn.cmd.Parameters.AddWithValue("@FromDate", _FromDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@Invoice", _Invoice)
            BaseConn.cmd.Parameters.AddWithValue("@UnInvoice", _UnInvoice)
            BaseConn.cmd.Parameters.AddWithValue("@DT", dt_ItemArray)
            BaseConn.cmd.Parameters.AddWithValue("@ShowAll", _ShowAll)
            BaseConn.cmd.Parameters.AddWithValue("@Completed", _Completed)
            BaseConn.cmd.Parameters.AddWithValue("@NotCompleted", _NotCompleted)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 2000
            BaseConn.da.Fill(dt)
            _ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrStr = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.ToString
        Finally
            BaseConn.Close()
        End Try
        Return dt

    End Function

    Public Function getItemDetails(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _CID As Integer, ByVal _ItemCode As String, MerchantLedger As Integer) As DataTable
        Try
            dt = New DataTable
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetItemDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", _ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", MerchantLedger)
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

    Public Function GetReportData(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _CID As Integer, ByVal _VouType As String, ByVal _VouNo As String) As DataSet
        Dim ds As New DataSet
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetReportData]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@VouType", _VouType)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", _VouNo)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(ds)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return ds
    End Function
    Public Function GetMainFormReportData(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _CID As Integer, ByVal _VouType As String, ByVal dt As DataTable, ByRef Status As String) As DataSet
        Dim ds As New DataSet
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetMainFormReportData]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@VouType", _VouType)
            BaseConn.cmd.Parameters.AddWithValue("@dt", dt)
            BaseConn.cmd.Parameters.Add("@Status", SqlDbType.VarChar, 100).Direction = ParameterDirection.Output
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(ds)
            ds.Tables(0).TableName = "SalesOrderMain"
            ds.Tables(1).TableName = "SalesOrder_Sub"

            Status = BaseConn.cmd.Parameters("@Status").Value.ToString
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return ds
    End Function

    Public Function FE_ProductionOrderTransmittalLedger(ByVal _strPath As String, ByVal _strpwd As String, ByVal _CID As String, ByVal _FrmDate As Date, ByVal _ToDate As Date) As DataTable
        FE_ProductionOrderTransmittalLedger = New DataTable
        Try
            BaseConn.Open(_strPath, _strpwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[FE_ProductionOrderTransmittalLedger]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@StartDate", _FrmDate)
            BaseConn.cmd.Parameters.AddWithValue("@EndDate", _ToDate)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            BaseConn.da.Fill(FE_ProductionOrderTransmittalLedger)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return FE_ProductionOrderTransmittalLedger
    End Function

    Public Function GetFinancialReportID(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, _CID As String) As DataTable
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetFinancialReportID]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            dt = New DataTable
            BaseConn.da.Fill(dt)

        Catch ex As Exception

        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Sub UpdateFinancialReportRights(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, _CID As String, _ReportID As String,
                                   ByVal _UserRights As DataTable, ByVal _GroupRights As DataTable, ByRef ErrNo As Int16)
        Dim objDalGeneral As New DAL_General(_CID)

        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("UpdateFinancialReportRights", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@ReportID", _ReportID)
            BaseConn.cmd.Parameters.AddWithValue("@UserRights", objDalGeneral.DatatableToJSONString(_UserRights))
            BaseConn.cmd.Parameters.AddWithValue("@GroupRights", objDalGeneral.DatatableToJSONString(_GroupRights))
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            'ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
        Catch ex As Exception
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub CheckNegativeStock(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, _CID As Integer, _VouNo As String, _ItemCode As String, ByRef _ItemType As String, ByRef _Qty As Double, ByRef _ErrNo As Integer)
        _ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[CheckNegativeStock]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@ItemCode", _ItemCode)
            BaseConn.cmd.Parameters.AddWithValue("@VouNo", _VouNo)
            BaseConn.cmd.Parameters.Add("@TypeOut", SqlDbType.VarChar, 30).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@QtyOUT", SqlDbType.Float).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
            _ItemType = BaseConn.cmd.Parameters("@TypeOut").Value.ToString
            _Qty = BaseConn.cmd.Parameters("@QtyOUT").Value.ToString
        Catch ex As Exception
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetDBList(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, _CID As String, ByRef _DBList As DataTable, _Flag As String, _Path As String)
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetDBList]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@Path", _Path)
            BaseConn.cmd.CommandTimeout = 5000
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            dt = New DataTable
            BaseConn.da.Fill(dt)
            _DBList = dt
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function GetCustomer(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, _CID As Integer) As DataTable
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetCustomer]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            dt = New DataTable
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function

    Public Function GetSalesmanList(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, _CID As Integer) As DataTable
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetSalesmanList]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            dt = New DataTable
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt
    End Function
End Class
