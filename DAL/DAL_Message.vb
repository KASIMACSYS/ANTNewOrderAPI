Imports Classes

Public Class DAL_Message

    Private dt As DataTable
    Private BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Sub GetUserGroupDetails(ByVal _DBPath As String, ByVal _DBPwd As String, ByVal _User As String, ByRef _SiteID As String, ByRef _DTMCCB As DataTable, _
                                   ByRef _DTLocal As DataTable, ByRef _DTMessage As DataTable, ByRef ErrNo As Integer, ByRef ErrMsg As String)
        Try
            BaseConn.Open(_DBPath, _DBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetUserGroupDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@User", _User)

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            'If ds.Tables(0).Rows.Count > 0 Then
            _DTMCCB = ds.Tables(0)
            'End If

            'If ds.Tables(1).Rows.Count > 0 Then
            _DTLocal = ds.Tables(1)
            'End If

            'If ds.Tables(2).Rows.Count > 0 Then
            _DTMessage = ds.Tables(2)
            'End If

        Catch ex As Exception
            ErrNo = 1
            ErrMsg = ex.Message ' "Problem in Updating Invoice"
        Finally
            BaseConn.Close()
        End Try
    End Sub


    Public Function MessageUpdate(ByVal _strPath As String, ByVal _strPwd As String, ByVal obj As csMessage, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("sp_MessageUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", obj.str_SiteID)
            'BaseConn.cmd.Parameters.AddWithValue("@MsgID", obj.int_MsgID)
            BaseConn.cmd.Parameters.AddWithValue("@MsgDate", obj.Date_MsgDate)
            BaseConn.cmd.Parameters.AddWithValue("@FromUser", obj.str_FromUser)
            BaseConn.cmd.Parameters.AddWithValue("@Subject", obj.str_Subject)
            BaseConn.cmd.Parameters.AddWithValue("@Message", obj.str_Message)

            BaseConn.cmd.Parameters.AddWithValue("@Comment", obj.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@PreMsgID", obj.int_PreMsgID)
            BaseConn.cmd.Parameters.AddWithValue("@ToUserDT", obj.ToUsers)

            BaseConn.cmd.ExecuteNonQuery()

        Catch ex As Exception
            _ErrString = ex.Message
            'ObjDalGeneral = New DAL_General(obj.str_SiteID)
            'ObjDalGeneral.Elog_Insert(obj.str_SiteID, _strPath, _strPwd, obj.objBOMMain.int_BusinessPeriodID, obj.str_CreatedBy, obj.dtp_CreatedDate, "", "BOM", ErrNo, "Error in " & obj.objBOMMain.str_Flag & " : " & obj.objBOMMain.str_BOMNo & " ", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        MessageUpdate = _ErrString
    End Function


    Public Function MessageStatusUpdate(ByVal _strPath As String, ByVal _strPwd As String, ByVal _SiteID As String, _
                                        ByVal _MsgID As Integer, ByVal _ToUser As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("sp_MessageStatusUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@MsgID", _MsgID)
            BaseConn.cmd.Parameters.AddWithValue("@ToUser", _ToUser)
            BaseConn.cmd.ExecuteNonQuery()

        Catch ex As Exception
            _ErrString = ex.Message
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        MessageStatusUpdate = _ErrString
    End Function

    Public Sub GetUnreadMsgCount(ByVal _strPath As String, ByVal _strPwd As String, ByVal _SiteID As String, _
                                       ByVal _ToUser As String, ByRef _UnReadMsgCnt As Integer, ByRef ErrNo As Integer)
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_strPath, _strPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("sp_GetUnreadMsgCount", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@User", _ToUser)
            'BaseConn.cmd.Parameters.AddWithValue("@UnReadRcd", _UnReadMsgCnt)

            BaseConn.cmd.Parameters.Add("@UnReadRcd", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            _UnReadMsgCnt = BaseConn.cmd.Parameters("@UnReadRcd").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

    End Sub
End Class
