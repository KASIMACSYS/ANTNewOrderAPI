Imports Classes
Imports System.Data.SqlClient

Public Class DAL_MenuGrouping

    Dim dt As DataTable
    Dim BaseConn As New SQLConn()
    Private ObjDalGeneral As DAL_General

    Public Function Get_Structure(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef _CID As String, ByRef _UniqID As String, ByRef _Flag As String,
                                        ByRef _ErrNo As Integer, ByRef _ErrStr As String) As DataTable
        Get_Structure = New DataTable
        _ErrNo = 0
        _ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetMenuGrouping]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@ID", _UniqID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            'BaseConn.cmd.Parameters.AddWithValue("@ApplicationType", 1)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            Get_Structure = ds.Tables(0)
            Return Get_Structure
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message.ToString
        End Try
    End Function

    'Public Function Get_GeneralParameter(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef _SiteID As String, ByRef _UniqID As String, ByRef _Flag As String,
    '                                    ByRef _ErrNo As Integer, ByRef _ErrStr As String) As DataTable
    '    Get_GeneralParameter = New DataTable
    '    _ErrNo = 0
    '    _ErrStr = ""
    '    Try
    '        BaseConn.Open(_StrDBPath, _StrDBPwd)
    '        BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetMenuGrouping]", BaseConn.cnn)
    '        BaseConn.cmd.CommandType = CommandType.StoredProcedure
    '        BaseConn.cmd.Parameters.AddWithValue("@SiteID", _SiteID)
    '        BaseConn.cmd.Parameters.AddWithValue("@ID", _UniqID)
    '        BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
    '        BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
    '        Dim ds As New DataSet
    '        BaseConn.da.Fill(ds)
    '        Get_GeneralParameter = ds.Tables(0)
    '        Return Get_GeneralParameter
    '    Catch ex As Exception
    '        _ErrNo = 1
    '        _ErrStr = ex.Message.ToString
    '    End Try
    'End Function

    Public Sub MenuGroupingUpdate(ByRef _StrDBPath As String, ByRef _StrDBPwd As String, ByRef _CID As String, ByRef _Flag As String, ByRef _UniqID As Integer,
                                   ByRef _Description As String, ByRef _ParentID As Integer, ByRef _UpdatedBy As String, ByRef _UpdatedDate As Date, ByRef _DTForm As DataTable,
                                  ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[MenuGroupingUpdate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@UniqID", _UniqID)
            BaseConn.cmd.Parameters.AddWithValue("@Description", _Description)
            BaseConn.cmd.Parameters.AddWithValue("@ParentID", _ParentID)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", _UpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", _UpdatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@FormDT", _DTForm)
            BaseConn.cmd.ExecuteNonQuery()

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub MenuFormUpdate(ByRef _StrDBPath As String, ByRef _StrDBPwd As String, ByRef _CID As String, ByRef _Flag As String, ByRef _UniqID As Integer, ByRef _SortID As Integer,
                              ByRef objMenuForm As csNewForm, ByRef _ErrStr As String)
        ObjDalGeneral = New DAL_General(_CID)
        _ErrStr = String.Empty
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[MenuFormUpdate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@UniqID", _UniqID)
            BaseConn.cmd.Parameters.AddWithValue("@SortID", _SortID)
            BaseConn.cmd.Parameters.AddWithValue("@MenuID", objMenuForm.MenuID)
            BaseConn.cmd.Parameters.AddWithValue("@Description", objMenuForm.Description)
            BaseConn.cmd.Parameters.AddWithValue("@ShortcutKey", objMenuForm.ShortcutKey)
            BaseConn.cmd.Parameters.AddWithValue("@Color", objMenuForm.Color)
            BaseConn.cmd.Parameters.AddWithValue("@Reserved", objMenuForm.Reserved)
            BaseConn.cmd.Parameters.AddWithValue("@LoadGroupMgt", objMenuForm.LoadGroupMgt)
            BaseConn.cmd.Parameters.AddWithValue("@ApplicationType", objMenuForm.ApplicationType)
            BaseConn.cmd.Parameters.AddWithValue("@Parameters", ObjDalGeneral.DatatableToJSONString(objMenuForm.Parameters)) 'objMenuForm.Parameters)
            BaseConn.cmd.Parameters.AddWithValue("@Options", ObjDalGeneral.DatatableToJSONString(objMenuForm.Options)) 'objMenuForm.Options)

            If objMenuForm.img_Photo Is Nothing Then
                Dim photoParam As New SqlParameter("@Image", SqlDbType.Image)
                photoParam.Value = DBNull.Value
                BaseConn.cmd.Parameters.Add(photoParam)
            Else
                BaseConn.cmd.Parameters.AddWithValue("@Image", objMenuForm.img_Photo) 'PHOTO
            End If

            BaseConn.cmd.ExecuteNonQuery()

        Catch ex As Exception
            _ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Sub GetMenuFormDetails(ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef _CID As String, ByRef _objNewForm As csNewForm, ByRef _UniqID As String, ByRef _SortID As Integer,
                                       ByRef _ErrNo As Integer, ByRef _ErrStr As String)
        ObjDalGeneral = New DAL_General(_CID)
        Dim dt As New DataTable
        _ErrNo = 0
        _ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetMenuFormDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@UniqID", _UniqID)
            BaseConn.cmd.Parameters.AddWithValue("@SortID", _SortID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)
            dt = ds.Tables(0)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                _objNewForm.MenuID = dt.Rows(0)("MenuID").ToString
                _objNewForm.Description = dt.Rows(0)("Description").ToString
                _objNewForm.ShortcutKey = dt.Rows(0)("ShortcutKey").ToString
                _objNewForm.Color = dt.Rows(0)("Color").ToString
                _objNewForm.Reserved = dt.Rows(0)("Reserved")
                _objNewForm.LoadGroupMgt = dt.Rows(0)("LoadGroupMgt")
                _objNewForm.ApplicationType = dt.Rows(0)("ApplicationType")
                _objNewForm.Parameters = ObjDalGeneral.GetDataTableFromJsonString(dt.Rows(0)("Parameters").ToString) 'dt.Rows(0)("Parameters").ToString
                _objNewForm.Options = ObjDalGeneral.GetDataTableFromJsonString(dt.Rows(0)("Options").ToString)

                Dim str As String = ds.Tables(0).Rows(0)("Icon").ToString()
                If str.Length > 0 Then
                    _objNewForm.img_Photo = CType(ds.Tables(0).Rows(0)("Icon"), Byte())
                Else
                    _objNewForm.img_Photo = Nothing
                End If

                'Dim dt_Options As New DataTable
                'dt_Options.Columns.Add("Options")
                'Dim drow As DataRow
                'drow = dt_Options.NewRow
                'drow("Options") = "Add"
                'dt_Options.Rows.Add(drow)
                'drow = dt_Options.NewRow
                'drow("Options") = "Edit"
                'dt_Options.Rows.Add(drow)
                'drow = dt_Options.NewRow
                'drow("Options") = "View"
                'dt_Options.Rows.Add(drow)
                '_objNewForm.Options = dt_Options
            End If
        Catch ex As Exception
            _ErrNo = 1
            _ErrStr = ex.Message.ToString
        End Try
    End Sub

    Public Sub MenuMovesUpdate(ByRef _StrDBPath As String, ByRef _StrDBPwd As String, ByRef _CID As String, ByRef _Menus As DataTable, ByRef _ErrStr As String)
        _ErrStr = String.Empty
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[MenuMovesUpdate]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _CID)
            BaseConn.cmd.Parameters.AddWithValue("@Menus", _Menus)
            BaseConn.cmd.ExecuteNonQuery()

        Catch ex As Exception
            _ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
End Class
