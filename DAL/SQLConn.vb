'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================

Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports AES_Cryptography

Public Class SQLConn
    Private connetionString As String '= "Data Source=ACSYS-MF;Initial Catalog=Falcon;Integrated Security=True"
    Private objCryptography As New AES()
    Public cnn As SqlConnection
    Public cmd As New SqlCommand
    Public ds As New DataSet
    Public da As New SqlDataAdapter
    Public dr As SqlDataReader
    Public _dtAllSite As DataTable

    Private access_Conn As OleDbConnection
    Private access_Cmd As OleDbCommand
    Private access_DA As OleDbDataAdapter
    Private dt As DataTable

    'Public Sub Open() ' Open Connection
    '    'connetionString = "Data Source=ACSYS-MF;Initial Catalog=Falcon;Integrated Security=True"
    '    'connetionString = "Data Source=ACSYS1007\SQLEXPRESS;Initial Catalog=Falcon;Integrated Security=True"
    '    cnn = New SqlConnection(connetionString)
    '    Try
    '        cnn.Open()
    '    Catch ex As Exception
    '        MsgBox("Can not open connection ! " & ex.Message)
    '    End Try
    'End Sub

    Public Sub Open(ByVal strPath As String, ByVal strPwd As String) ' Open Connection
        connetionString = strPath ' 
        'connetionString = "Data Source=ACSYS-E17\ACSYSERP2017;Initial Catalog=V3.0b1;Integrated Security=False; User ID=sa;Connect Timeout=300;"
        'strPwd = "acSys@123"
        cnn = New SqlConnection(connetionString + "Password=" + strPwd)

        Try
            'Using cnn = New SqlConnection(connetionString + "Password=" + strPwd)
            cnn.Open()
            'End Using

        Catch ex As Exception
            MsgBox("Can not open connection ! " & ex.Message)
        End Try
    End Sub

    Public Sub Close() ' Close Connection
        Try
            cnn.Close()
            cnn.Dispose()
        Catch ex As Exception
            MsgBox("Can not close connection ! " & ex.Message)
        End Try
    End Sub

    Public Function getAllSite_Access(ByRef dt_AllSite As DataTable, ByRef _DTPOSConfig As DataTable, ByRef errNo As Integer) As String
        Dim ErrStr As String = String.Empty
        Try
            Dim cur_appl_dir As String
            cur_appl_dir = System.IO.Directory.GetCurrentDirectory
            'Dim connetionString As String = "Provider=Microsoft.Jet.Oledb.4.0;Data Source=" & cur_appl_dir & "\acSysERP.mdb;Jet OLEDB:Database Password=acSysERP"
            Dim connetionString As String = "Provider=Microsoft.Jet.Oledb.4.0;Data Source=" & cur_appl_dir & "\acSysERP.acs;Jet OLEDB:Database Password=acSysERP"
            access_Conn = New OleDbConnection(connetionString)
            access_Conn.Open()
            dt = New DataTable
            access_Cmd = New OleDbCommand("Select SiteID,SiteName,DBPath,Password,Default,'False' as LoginSiteID,'False' as IsAllowSite,ShowForm2,ShowOutLook from SiteMaster order by Default ASC,SiteID ASC", access_Conn)
            access_DA = New OleDbDataAdapter(access_Cmd)
            access_DA.Fill(dt)
            dt_AllSite = dt

            Dim i As Integer = dt_AllSite.Rows.Count
            For i = 0 To dt_AllSite.Rows.Count - 1
                Dim _pwd As String = objCryptography.AES_Decrypt(dt_AllSite.Rows(i)("Password").ToString)
                dt_AllSite.Rows(i)("Password") = _pwd
            Next
            dt_AllSite.AcceptChanges()

            dt = New DataTable
            access_Cmd = New OleDbCommand("Select TagID,TagValue from POSConfig", access_Conn)
            access_DA = New OleDbDataAdapter(access_Cmd)
            access_DA.Fill(_DTPOSConfig)
        Catch ex As Exception
            errNo = 1
            ErrStr = "Can not open connection ! " & ex.Message.ToString '  MsgBox("Can not open connection ! " & ex.Message)
        Finally
            access_Conn.Close()
        End Try
        Return ErrStr
    End Function
    ''======================================================================================
    ''  Description : Encryption function (copied from internet and modified)
    ''  Author: R. Mohamed Faizal 
    ''  Date  : 30/03/08
    ''======================================================================================
    'Friend Function EncryptText(ByRef strText As String) As Object
    '    Dim i, c As Short
    '    Dim strBuff As String = ""
    '    Dim strpwd As String = "acsysit"
    '    If Len(strpwd) Then
    '        For i = 1 To Len(strText)
    '            c = Asc(Mid(strText, i, 1))
    '            c = c + Asc(Mid(strpwd, (i Mod Len(strpwd)) + 1, 1))
    '            strBuff = strBuff & Chr(c And &HFFS)
    '        Next i
    '    Else
    '        strBuff = strText
    '    End If
    '    EncryptText = strBuff
    'End Function

    ''======================================================================================
    ''  Description : Decrypt text encrypted with EncryptText
    ''  Author: R. Mohamed Faizal 
    ''  Date  : 30/03/08
    ''======================================================================================
    'Friend Function DecryptText(ByRef strText As String) As Object
    '    Dim i, c As Short
    '    Dim strBuff As String = ""
    '    Dim strpwd As String = "acsysit"
    '    If Len(strpwd) Then
    '        For i = 1 To Len(strText)
    '            c = Asc(Mid(strText, i, 1))
    '            c = c - Asc(Mid(strpwd, (i Mod Len(strpwd)) + 1, 1))
    '            strBuff = strBuff & Chr(c And &HFFS)
    '        Next i
    '    Else
    '        strBuff = strText
    '    End If
    '    DecryptText = strBuff
    'End Function
End Class
