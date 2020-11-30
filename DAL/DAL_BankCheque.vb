'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Imports Classes

Public Class DAL_BankCheque
    Private BaseConn As New SQLConn()
    Private dt As DataTable
    Private ObjDalGeneral As DAL_General

    Public Function GetBankLedgers(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _SiteID As Integer, ByRef _ErrNo As Integer) As DataTable
        dt = New DataTable
        _ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetBankLedgers]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = ds.Tables(0)
        Catch ex As Exception
            _ErrNo = 1
        End Try
        Return dt
    End Function

    Public Function Get_Structure(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _SiteID As Integer, _
                                  ByVal _Flag As String, ByVal _ChqDate As Date, ByVal _ChqStatus As String, ByRef _ErrNo As Integer) As DataTable
        dt = New DataTable
        _ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetChequeDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            'BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@ChequeDate", _ChqDate)
            BaseConn.cmd.Parameters.AddWithValue("@ChequeStatus", _ChqStatus)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = ds.Tables(0)
        Catch ex As Exception
            _ErrNo = 1
        End Try
        Return dt
    End Function

    Public Function Put_Structure(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _SiteID As String, ByVal _Flag As String, ByVal ChqDepositDT As DataTable, ByRef _ErrNo As Integer) As String
        Put_Structure = String.Empty
        Dim _ErrString As String = ""
        _ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[ChequeDepositUpdated]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@ChequeDeposit", ChqDepositDT)
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            _ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(_SiteID)
            ObjDalGeneral.Elog_Insert(_SiteID, _DBPath, _DBPwd, 0, "", DateTime.Now, "", "ChequeDepositUpdated", Err.Number, "Error in " & _Flag & " : ChequeDepositUpdated", ex.Message, 5, 3, 1, _ErrNo)
            _ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
        Return _ErrString
    End Function
End Class

Public Class DAL_BankReconcilation
    Private BaseConn As New SQLConn()
    Private dt As DataTable
    Private ObjDalGeneral As DAL_General

    Public Function Get_Structure(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _SiteID As Integer, ByVal _BusinessPeriodID As Integer, ByVal _DstLedgerID As String, _
                                  ByVal _DateType As String, ByVal _FrmDate As Date, ByVal _ToDate As Date, ByVal _StatusAll As String, ByVal _StatusOpen As String, _
                                  ByVal _StatusClosed As String, ByVal _StatusBounced As String, ByRef _BankBal As Double, ByRef _OutDebits As Double, ByRef _OutCredits As Double, ByRef _BalBF As Double, ByRef _ErrNo As Integer) As DataTable
        dt = New DataTable
        _ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetCheque4ReconcilationSub]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@DstLedgerID", _DstLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@DateType", _DateType)
            BaseConn.cmd.Parameters.AddWithValue("@FrmDate", _FrmDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)
            BaseConn.cmd.Parameters.AddWithValue("@StatusAll", _StatusAll)
            BaseConn.cmd.Parameters.AddWithValue("@StatusOpen", _StatusOpen)
            BaseConn.cmd.Parameters.AddWithValue("@StatusClosed", _StatusClosed)
            BaseConn.cmd.Parameters.AddWithValue("@StatusBounced", _StatusBounced)

            BaseConn.cmd.Parameters.Add("@BankBal", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutDebits", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutCredits", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@BalBF", SqlDbType.Decimal).Direction = ParameterDirection.Output

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            _BankBal = BaseConn.cmd.Parameters("@BankBal").Value
            _OutDebits = BaseConn.cmd.Parameters("@OutDebits").Value
            _OutCredits = BaseConn.cmd.Parameters("@OutCredits").Value
            _BalBF = BaseConn.cmd.Parameters("@BalBF").Value

            dt = ds.Tables(0)

        Catch ex As Exception
            _ErrNo = 1

        End Try
        Return dt
    End Function

  Public Sub Get_BankBalDetails(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _SiteID As Integer, ByVal _BusinessPeriodID As Integer, ByVal _DstLedgerID As String, _
                                 ByVal _FrmDate As Date, ByVal _ToDate As Date, ByRef _BankBal As Double, ByRef _OutDebits As Double, ByRef _OutCredits As Double, ByRef _BalBF As Double, ByRef _ErrNo As Integer)

        _ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetBankBalanceDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@DstLedgerID", _DstLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@FrmDate", _FrmDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)

            BaseConn.cmd.Parameters.Add("@BankBal", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutDebits", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutCredits", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@BalBF", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()

            _BankBal = BaseConn.cmd.Parameters("@BankBal").Value
            _OutDebits = BaseConn.cmd.Parameters("@OutDebits").Value
            _OutCredits = BaseConn.cmd.Parameters("@OutCredits").Value
            _BalBF = BaseConn.cmd.Parameters("@BalBF").Value
        Catch ex As Exception
            _ErrNo = 1
        End Try
    End Sub

    Public Function Get_ReconcilationMain(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _SiteID As Integer, ByVal _BusinessPeriodID As Integer, ByVal _DstLedgerID As String, ByVal dtp_date As String,
                                ByVal _FrmDate As Date, ByVal _ToDate As Date,
                                 ByRef _BankBal As Double, ByRef _OutDebits As Double, ByRef _OutCredits As Double, ByRef _BalBF As Double, ByRef _ErrNo As Integer) As DataTable
        dt = New DataTable
        _ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetCheque4Reconcilation]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@DstLedgerID", _DstLedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Date1", dtp_date)
            BaseConn.cmd.Parameters.AddWithValue("@FrmDate", _FrmDate)
            BaseConn.cmd.Parameters.AddWithValue("@ToDate", _ToDate)

            BaseConn.cmd.Parameters.Add("@BankBal", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutDebits", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutCredits", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@BalBF", SqlDbType.Decimal).Direction = ParameterDirection.Output
            BaseConn.cmd.CommandTimeout = 500
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            _BankBal = BaseConn.cmd.Parameters("@BankBal").Value
            _OutDebits = BaseConn.cmd.Parameters("@OutDebits").Value
            _OutCredits = BaseConn.cmd.Parameters("@OutCredits").Value
            _BalBF = BaseConn.cmd.Parameters("@BalBF").Value

            dt = ds.Tables(0)

        Catch ex As Exception
            _ErrNo = 1
        End Try
        Return dt
    End Function

    Public Function Put_Structure(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _SiteID As String, ByVal _MenuID As String, ByVal ChqStatusDT As DataTable, ByRef _ErrNo As Integer) As String
        Put_Structure = String.Empty
        Dim _ErrString As String = ""
        _ErrNo = 0
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[ChequeStatusUpdated]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", _MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@ChequeStatus", ChqStatusDT)
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            _ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrNo = 1
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(_SiteID)
            ObjDalGeneral.Elog_Insert(_SiteID, _DBPath, _DBPwd, 0, "", DateTime.Now, "", "ChequeStatusUpdated", _ErrNo, "Error in ChequeStatusUpdated", ex.Message, 5, 3, 1, _ErrNo)
        Finally
            BaseConn.Close()
        End Try

        Return _ErrString
    End Function
End Class
