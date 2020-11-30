'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes
Public Class DAL_PayCert
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub Get_Approve(ByRef Obj As csPayCert, ByVal _strPath As String, ByVal _strPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetPayCertApprove]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.objPayCertMain.int_BusinessPeriod)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", Obj.objPayCertMain.int_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_CID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.objPayCertMain.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@PCNo", Obj.objPayCertMain.str_PCNo)
            BaseConn.cmd.Parameters.AddWithValue("@SetFlag", Obj.objPayCertMain.str_setFlag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            If Obj.objPayCertMain.str_setFlag = "GetmerchantDetail" Then
                Obj.objPayCertMain.int_LedgerID = ds.Tables(0).Rows(0)("LedgerID").ToString()
                Obj.objPayCertMain.str_MerchantName = ds.Tables(0).Rows(0)("MerchantName").ToString()
                Obj.objPayCertMain.int_Aging = ds.Tables(0).Rows(0)("Aging").ToString()
                Obj.objPayCertMain.str_PayTerm = ds.Tables(0).Rows(0)("PayTerm").ToString()
                Obj.objPayCertMain.str_Comment = ds.Tables(0).Rows(0)("Comment").ToString()
                Obj.objPayCertMain.dtp_PCDate = ds.Tables(0).Rows(0)("PCDate").ToString()
                Obj.objPayCertMain.dtp_INVDate = ds.Tables(0).Rows(0)("INVDate").ToString()
                Obj.objPayCertMain.dbl_TCAmount = ds.Tables(0).Rows(0)("TCAmount").ToString()

                Obj.objPayCertMain.dbl_AdvanceAmount = ds.Tables(0).Rows(0)("AdvanceAmount").ToString()

                Obj.objPayCertMain.str_TCCurrency = ds.Tables(0).Rows(0)("TCCurrency").ToString()
                Obj.objPayCertMain.dbl_ExchangeRate = ds.Tables(0).Rows(0)("ExchangeRate").ToString()
                Obj.objPayCertMain.str_createdby = ds.Tables(0).Rows(0)("CreatedBy").ToString
                Obj.objPayCertMain.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
                Obj.objPayCertMain.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
                Obj.objPayCertMain.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
                Obj.objPayCertMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
                Obj.objPayCertMain.str_Desc5 = ds.Tables(0).Rows(0)("Desc5").ToString()
                Obj.objPayCertMain.str_Desc6 = ds.Tables(0).Rows(0)("Desc6").ToString()
                Obj.objPayCertMain.str_Desc7 = ds.Tables(0).Rows(0)("Desc7").ToString()
                Obj.objPayCertMain.str_Desc8 = ds.Tables(0).Rows(0)("Desc8").ToString()
            End If
            If Obj.objPayCertMain.str_Flag = "Approve" Then
                Obj.objPayCertSub.dt_PayCert = ds.Tables(1)
            End If
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Get_Structure(ByVal _strPath As String, ByVal _strPwd As String, ByVal BusinessPeriodID As String, ByVal Flag As String, ByVal LedgerID As String, ByVal CID As String, ByVal str_PCNo As String) As DataTable
        dt = New DataTable
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetPayCert]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Flag)
            BaseConn.cmd.Parameters.AddWithValue("@CID", CID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@PCNo", str_PCNo)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            'Dim ds As New DataSet
            BaseConn.da.Fill(dt)

        Catch ex As Exception
            MsgBox("Error" & ex.Message)
        Finally
            BaseConn.Close()
        End Try
        Return dt

    End Function

    Public Function Update_PayCert(ByVal obj As csPayCert, ByRef PCNo As String, ByRef intRevNo As Integer, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("PayCertUpdated", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_CID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.objPayCertMain.int_BusinessPeriod)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.objPayCertMain.str_Flag)

            BaseConn.cmd.Parameters.AddWithValue("@MenuID", obj.objPayCertMain.str_MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@FormPrefix", obj.objPayCertMain.str_PayCertPrefix)
            BaseConn.cmd.Parameters.AddWithValue("@PCNo", obj.objPayCertMain.str_PCNo)
            BaseConn.cmd.Parameters.AddWithValue("@RevNo", obj.objPayCertMain.int_RevNo)
            BaseConn.cmd.Parameters.AddWithValue("@PCDate", obj.objPayCertMain.dtp_PCDate)
            BaseConn.cmd.Parameters.AddWithValue("@INVDate", obj.objPayCertMain.dtp_INVDate)
            BaseConn.cmd.Parameters.AddWithValue("@PCStatus", obj.objPayCertMain.int_PCStatus)

            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.objPayCertMain.int_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@MerchantName", obj.objPayCertMain.str_MerchantName)
            BaseConn.cmd.Parameters.AddWithValue("@Aging", obj.objPayCertMain.int_Aging)
            BaseConn.cmd.Parameters.AddWithValue("@PayTerm", obj.objPayCertMain.str_PayTerm)
            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.objPayCertMain.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@TotApproveAmt", obj.objPayCertMain.dbl_TotApproveAmt)
            BaseConn.cmd.Parameters.AddWithValue("@TCAmount", obj.objPayCertMain.dbl_TCAmount)
            BaseConn.cmd.Parameters.AddWithValue("@TCCurrency", obj.objPayCertMain.str_TCCurrency)
            BaseConn.cmd.Parameters.AddWithValue("@ExchangeRate", obj.objPayCertMain.dbl_ExchangeRate)
            BaseConn.cmd.Parameters.AddWithValue("@AdvanceAmount", obj.objPayCertMain.dbl_AdvanceAmount)

            ''AM Specific
            BaseConn.cmd.Parameters.AddWithValue("@Desc1", obj.objPayCertMain.str_Desc1)
            BaseConn.cmd.Parameters.AddWithValue("@Desc2", obj.objPayCertMain.str_Desc2)
            BaseConn.cmd.Parameters.AddWithValue("@Desc3", obj.objPayCertMain.str_Desc3)
            BaseConn.cmd.Parameters.AddWithValue("@Desc4", obj.objPayCertMain.str_Desc4)
            BaseConn.cmd.Parameters.AddWithValue("@Desc5", obj.objPayCertMain.str_Desc5)
            BaseConn.cmd.Parameters.AddWithValue("@Desc6", obj.objPayCertMain.str_Desc6)
            BaseConn.cmd.Parameters.AddWithValue("@Desc7", obj.objPayCertMain.str_Desc7)
            BaseConn.cmd.Parameters.AddWithValue("@Desc8", obj.objPayCertMain.str_Desc8)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedStatus", obj.objPayCertMain.bool_ApprovedStatus)
            BaseConn.cmd.Parameters.AddWithValue("@PayCertDetailsDT", obj.objPayCertSub.dt_PayCert)
            BaseConn.cmd.Parameters.Add("@PCNoOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@OutRevNo", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()

            PCNo = BaseConn.cmd.Parameters("@PCNoOut").Value.ToString
            intRevNo = BaseConn.cmd.Parameters("@OutRevNo").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_CID)
            ObjDalGeneral.Elog_Insert(obj.str_CID, _StrDBPath, _StrDBPwd, obj.objPayCertMain.int_BusinessPeriod, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "PaymentCertificate", Err.Number, "Error in " & obj.objPayCertMain.str_Flag & " : " & obj.objPayCertMain.str_PCNo & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try
        Update_PayCert = _ErrString
    End Function
End Class
