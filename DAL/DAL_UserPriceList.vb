'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports Classes

Public Class DAL_UserPriceList
    Private ObjDalGeneral As DAL_General
    Dim dt As DataTable
    'Dim objcsIndent As New csIndent
    Dim BaseConn As New SQLConn()

    Public Sub Get_Structure(ByRef Obj As csUserPriceList, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal ErrNo As String, ByVal ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetUserPriceList]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@UserID", Obj.Int_UserID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManID", Obj.Int_SalesMan)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            Obj.dt_UserPriceList = ds.Tables(0)
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub


    Public Function Update_UserPriceList(ByRef obj As csUserPriceList, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("UpdateUserPriceList", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.int_CID)
            BaseConn.cmd.Parameters.AddWithValue("@UserID", obj.Int_UserID)
            BaseConn.cmd.Parameters.AddWithValue("@SalesManID", obj.Int_SalesMan)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.Str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.Str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.dtp_LastUpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@PriceListDT", obj.dt_UserPriceList)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedBy", obj.Str_ApprovedBy)
            BaseConn.cmd.Parameters.AddWithValue("@ApprovedDate", obj.dtp_ApprovedDate)

            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.int_CID)
            ObjDalGeneral.Elog_Insert(obj.int_CID, _StrDBPath, _StrDBPwd, 0, obj.Int_UserID, obj.dtp_LastUpdatedDate, "", "UserPriceList", Err.Number, "Error in " & obj.str_Flag & " : " & obj.Int_UserID & "  ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_UserPriceList = _ErrString
    End Function


End Class
