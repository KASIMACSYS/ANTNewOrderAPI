Imports Microsoft.Office.Interop
Imports System.Data.SqlClient
Imports System.Runtime.InteropServices

Public Class DAL_Excel
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Function ImportExcel2DataTable(ByVal Path As String, ByRef ErrNo As Integer, ByRef ErrString As String) As DataTable
        ErrNo = 0
        ErrString = String.Empty

        Dim MyConnection As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection
        Dim DtTable As New DataTable
        Dim strShetname As String
        Dim DtSet As System.Data.DataSet
        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter

        Dim oApp As Excel.Application = Nothing
        Dim oBooks As Excel.Workbooks = Nothing
        Dim oBook As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing

        Try
            'MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; " & "Data Source=" & Path & "; " & "Extended Properties=""Excel 8.0;HDR=YES""")
            'MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; " & "Data Source=" & Path & "; ")
            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Path & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=2""")
            oApp = New Excel.Application
            oBooks = oApp.Workbooks
            oBook = oBooks.Add(Path)
            oSheet = oApp.ActiveSheet
            strShetname = oSheet.Name

            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [" & strShetname & "$]", MyConnection)
            MyCommand.TableMappings.Add("Table", "Fusion")

            DtSet = New System.Data.DataSet

            MyCommand.Fill(DtSet)

            MyConnection.Close()

            DtTable = DtSet.Tables(0)
        Catch ex As Exception
            ErrNo = 1
            ErrString = ex.Message.ToString
        Finally

            Marshal.ReleaseComObject(oSheet)
            Marshal.ReleaseComObject(oBook)
            Marshal.ReleaseComObject(oBooks)
            Marshal.ReleaseComObject(oApp)

            oSheet = Nothing
            oBook = Nothing
            oBooks = Nothing
            oApp = Nothing

            GC.Collect()
            GC.WaitForPendingFinalizers()

        End Try


        ImportExcel2DataTable = DtTable
    End Function

    Public Sub UpdateItemMaster(ByVal str_SiteID As String, ByVal _strPath As String, ByVal _strPWD As String, ByVal _BusinessPeriodID As Integer, _
                                ByVal _LoggedUser As String, ByVal _WHID As Integer, ByVal DTItems As DataTable, ByRef _ErrNo As Integer, ByRef _ErrString As String)
        _ErrNo = 0
        _ErrString = ""
        Try
            BaseConn.Open(_strPath, _strPWD)
            BaseConn.cmd = New SqlClient.SqlCommand("[UpdateItemMasterFromExcel]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", _LoggedUser)
            BaseConn.cmd.Parameters.AddWithValue("@WHID", _WHID)
            BaseConn.cmd.Parameters.AddWithValue("@DTItems", DTItems)
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()

            _ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrNo = 1
            _ErrString = ex.Message.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Sub UpdateEmployeeMaster(ByVal str_SiteID As String, ByVal _strPath As String, ByVal _strPWD As String, ByVal _BusinessPeriodID As Integer, _
                               ByVal _LoggedUser As String, ByVal DTItems As DataTable, ByRef _ErrNo As Integer, ByRef _ErrString As String)
        _ErrNo = 0
        _ErrString = ""
        Try
            BaseConn.Open(_strPath, _strPWD)
            BaseConn.cmd = New SqlClient.SqlCommand("[UpdateEmployeeMasterFromExcel]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", _LoggedUser)
            BaseConn.cmd.Parameters.AddWithValue("@EmployeeDT", DTItems)
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            _ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrNo = 1
            _ErrString = ex.Message.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Sub UpdateAsset(ByVal str_SiteID As String, ByVal _strPath As String, ByVal _strPWD As String, ByVal _BusinessPeriodID As Integer, _
                               ByVal _LoggedUser As String, ByVal DTItems As DataTable, ByRef _ErrNo As Integer, ByRef _ErrString As String)
        _ErrNo = 0
        _ErrString = ""
        Try
            BaseConn.Open(_strPath, _strPWD)
            BaseConn.cmd = New SqlClient.SqlCommand("[UpdateAssetFromExcel]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", _LoggedUser)
            BaseConn.cmd.Parameters.AddWithValue("@AssetDT", DTItems)
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            _ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrNo = 1
            _ErrString = ex.Message.ToString
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub UpdateItemPrice(ByVal str_SiteID As String, ByVal _strPath As String, ByVal _strPWD As String, ByVal DTItems As DataTable, ByVal DTTax As DataTable, _
                                ByRef _ErrNo As Integer, ByRef _ErrString As String, ByVal _Flag As String)

        Try

            BaseConn.Open(_strPath, _strPWD)
            BaseConn.cmd = New SqlClient.SqlCommand("[UpdateItemPriceFromExcel]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@DTItems", DTItems)
            BaseConn.cmd.Parameters.AddWithValue("@DTTax", DTTax)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.ExecuteNonQuery()
        Catch ex As Exception
            _ErrNo = 1
            _ErrString = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Sub UpdateItemTax(ByVal str_SiteID As String, ByVal _strPath As String, ByVal _strPWD As String, ByVal DTTax As DataTable, _
                                ByRef _ErrNo As Integer, ByRef _ErrString As String, ByVal _Flag As String)

        Try

            BaseConn.Open(_strPath, _strPWD)
            BaseConn.cmd = New SqlClient.SqlCommand("[UpdateItemTaxFromExcel]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@DTTax", DTTax)
            'BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()
        Catch ex As Exception
            _ErrNo = 1
            _ErrString = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Sub UpdateEmpDocument(ByVal str_SiteID As String, ByVal _strPath As String, ByVal _strPWD As String, ByVal DTItems As DataTable, _
                                 ByVal UserName As String, ByRef _ErrNo As Integer, ByRef _ErrString As String)

        Try

            BaseConn.Open(_strPath, _strPWD)
            BaseConn.cmd = New SqlClient.SqlCommand("[UpdateEmployeeDocumentFromExcel]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@DT", DTItems)
            BaseConn.cmd.Parameters.AddWithValue("@UserName", UserName)
            BaseConn.cmd.ExecuteNonQuery()
        Catch ex As Exception
            _ErrNo = 1
            _ErrString = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Sub ImportSISfromExcel(ByVal _strPath As String, ByVal _strPwd As String, ByVal _SiteID As String, ByVal _BSID As Integer,
                                      ByVal _JVLedgerID As Integer, ByVal _SISMainDT As DataTable,
                             ByVal _CreatedBy As String, ByRef _ErrNo As Integer, ByRef _ErrDesc As String)
        Dim _ErrString As String = ""
        _ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("ImportSISfromExcel", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BSID)
            BaseConn.cmd.Parameters.AddWithValue("@JVLedgerID", _JVLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@SISMainDT", _SISMainDT)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", _CreatedBy)

            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()

            _ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            _ErrDesc = _ErrString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(_SiteID)
            ObjDalGeneral.Elog_Insert(_SiteID, _strPath, _strPwd, _BSID, _CreatedBy, Date.Now, "", "SIS", Err.Number, "Error in Import from Excel :", ex.Message, 5, 3, 1, _ErrNo)
            _ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub ImportPIPfromExcel(ByVal _strPath As String, ByVal _strPwd As String, ByVal _SiteID As String, ByVal _BSID As Integer,
                                      ByVal _JVLedgerID As Integer, ByVal _PIPMainDT As DataTable,
                             ByVal _CreatedBy As String, ByRef _ErrNo As Integer, ByRef _ErrDesc As String)
        Dim _ErrString As String = ""
        _ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("ImportPIPfromExcel", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BSID)
            BaseConn.cmd.Parameters.AddWithValue("@JVLedgerID", _JVLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@PIPMainDT", _PIPMainDT)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", _CreatedBy)

            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.cmd.ExecuteNonQuery()

            _ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            _ErrDesc = _ErrString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(_SiteID)
            ObjDalGeneral.Elog_Insert(_SiteID, _strPath, _strPwd, _BSID, _CreatedBy, Date.Now, "", "PIP", Err.Number, "Error in Import from Excel :", ex.Message, 5, 3, 1, _ErrNo)
            _ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub UpdateMerchant(ByVal str_SiteID As String, ByVal _strPath As String, ByVal _strPWD As String, ByVal _BusinessPeriodID As Integer, ByVal _LoggedUser As String, ByVal DTItems As DataTable, ByRef _ErrNo As Integer, ByRef _ErrDesc As String)
        _ErrNo = 0
        Dim _ErrString As String = ""
        Try

            BaseConn.Open(_strPath, _strPWD)
            BaseConn.cmd = New SqlClient.SqlCommand("[ImportMerchantFromExcel]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", _LoggedUser)
            BaseConn.cmd.Parameters.AddWithValue("@MerchantMainDT", DTItems)
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()

            _ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
            _ErrDesc = _ErrString
        Catch ex As Exception
            _ErrNo = 1
            _ErrString = ex.Message
            _ErrDesc = _ErrString
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Sub UpdateMinMaxQty(ByVal str_SiteID As String, ByVal _strPath As String, ByVal _strPWD As String, ByVal DTItems As DataTable, _
                                ByRef _ErrNo As Integer, ByRef _ErrString As String)

        Try

            BaseConn.Open(_strPath, _strPWD)
            BaseConn.cmd = New SqlClient.SqlCommand("[UpdateMinMaxQty]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@DTItems", DTItems)
            BaseConn.cmd.ExecuteNonQuery()
        Catch ex As Exception
            _ErrNo = 1
            _ErrString = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub UpdateBOMNo(ByVal str_SiteID As String, ByVal _strPath As String, ByVal _strPWD As String, ByVal DTItems As DataTable, _
                               ByRef _ErrNo As Integer, ByRef _ErrString As String)

        Try

            BaseConn.Open(_strPath, _strPWD)
            BaseConn.cmd = New SqlClient.SqlCommand("[UpdateBOMNo]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@DTItems", DTItems)
            BaseConn.cmd.ExecuteNonQuery()
        Catch ex As Exception
            _ErrNo = 1
            _ErrString = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Sub ImportItemBarcodeFromExcel(ByVal str_SiteID As String, ByVal _strPath As String, ByVal _strPWD As String, ByVal DTItems As DataTable, _
                             ByRef _ErrNo As Integer, ByRef _ErrString As String)

        Try

            BaseConn.Open(_strPath, _strPWD)
            BaseConn.cmd = New SqlClient.SqlCommand("[ImportItemBarCodeFromExcel]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@DTItems", DTItems)
            BaseConn.cmd.ExecuteNonQuery()
        Catch ex As Exception
            _ErrNo = 1
            _ErrString = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
End Class
