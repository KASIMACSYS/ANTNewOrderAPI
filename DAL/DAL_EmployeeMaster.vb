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

Public Class DAL_EmployeeMaster
    Private ObjDalGeneral As DAL_General
    Dim dt As DataTable
    Dim BaseConn As New SQLConn()

    Public Sub Get_Structure(ByRef Obj As Classes.csEmployeeMaster, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetEmployeeMasterDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@EmpID", Obj.ObjEmpMain.str_EmpID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.ObjEmpCommon.str_Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            Obj.ObjEmpMain.str_EmpID = ds.Tables(0).Rows(0)("EmpID").ToString()
            Obj.ObjEmpMain.str_FirstName = ds.Tables(0).Rows(0)("FirstName").ToString()
            Obj.ObjEmpMain.str_LastName = ds.Tables(0).Rows(0)("LastName").ToString()
            Obj.ObjEmpMain.str_AliasName1 = ds.Tables(0).Rows(0)("AliasName1").ToString()
            Obj.ObjEmpMain.str_AliasName2 = ds.Tables(0).Rows(0)("AliasName2").ToString()
            Obj.ObjEmpMain.dtp_JoiningDate = ds.Tables(0).Rows(0)("JoiningDate").ToString()
            Obj.ObjEmpMain.str_Designation = ds.Tables(0).Rows(0)("Designation").ToString()
            Obj.ObjEmpMain.str_Category = ds.Tables(0).Rows(0)("Category").ToString()
            Obj.ObjEmpMain.str_SubCategory = ds.Tables(0).Rows(0)("SubCategory").ToString()
            Obj.ObjEmpMain.str_Department = ds.Tables(0).Rows(0)("Department").ToString()
            Obj.ObjEmpMain.str_Nationality = ds.Tables(0).Rows(0)("Nationality").ToString()
            Obj.ObjEmpMain.str_Language = ds.Tables(0).Rows(0)("Language").ToString()

            Obj.ObjEmpMain.bool_SellFlag = ds.Tables(0).Rows(0)("SellFlag").ToString()

            Obj.ObjEmpMain.str_Comment = ds.Tables(0).Rows(0)("EmpComment").ToString()
            Obj.ObjEmpMain.dtp_DOB = ds.Tables(0).Rows(0)("DOB").ToString()
            Obj.ObjEmpMain.str_ICE1Name = ds.Tables(0).Rows(0)("ICE1Name").ToString()
            Obj.ObjEmpMain.str_ICE1No = ds.Tables(0).Rows(0)("ICE1No").ToString()
            Obj.ObjEmpMain.str_ICE1Comment = ds.Tables(0).Rows(0)("ICE1Comment").ToString()
            Obj.ObjEmpMain.str_ICE2Name = ds.Tables(0).Rows(0)("ICE2Name").ToString()
            Obj.ObjEmpMain.str_ICE2No = ds.Tables(0).Rows(0)("ICE2No").ToString()
            Obj.ObjEmpMain.str_ICE2Comment = ds.Tables(0).Rows(0)("ICE2Comment").ToString()
            Obj.ObjEmpLedger.bool_InActive = ds.Tables(0).Rows(0)("InActive").ToString
            Obj.ObjEmpMain.dbl_PassageAmount = ds.Tables(0).Rows(0)("PassageAmount").ToString()
            Obj.ObjEmpMain.str_syncID1 = ds.Tables(0).Rows(0)("SyncID1").ToString()
            Obj.ObjEmpMain.intGender = ds.Tables(0).Rows(0)("Gender").ToString()
            Obj.ObjEmpMain.int_EOSType = ds.Tables(0).Rows(0)("EOSType").ToString()
            Obj.ObjEmpMain.str_FamilyName = ds.Tables(0).Rows(0)("FamilyName").ToString()
            Obj.ObjEmpMain.str_MaritalStatus = ds.Tables(0).Rows(0)("MaritalStatus").ToString()
            Obj.ObjEmpMain.str_BloodGroup = ds.Tables(0).Rows(0)("BloodGroup").ToString()
            Obj.ObjEmpMain.str_Desc1 = ds.Tables(0).Rows(0)("Desc1").ToString()
            Obj.ObjEmpMain.str_Desc2 = ds.Tables(0).Rows(0)("Desc2").ToString()
            Obj.ObjEmpMain.str_Desc3 = ds.Tables(0).Rows(0)("Desc3").ToString()
            Obj.ObjEmpMain.str_Desc4 = ds.Tables(0).Rows(0)("Desc4").ToString()
            Obj.ObjEmpMain.bool_SendEmail = ds.Tables(0).Rows(0)("SendEmail").ToString()
            Obj.ObjEmpMain.bool_SendSMS = ds.Tables(0).Rows(0)("SendSMS").ToString()

            Dim str As String = ds.Tables(0).Rows(0)("Photo").ToString()
            If str.Length > 0 Then
                Obj.ObjEmpMain.img_Photo = CType(ds.Tables(0).Rows(0)("Photo"), Byte())
            Else
                Obj.ObjEmpMain.img_Photo = Nothing
            End If

            Obj.ObjEmpLedger.int_LedgerID = ds.Tables(1).Rows(0)("LedgerID").ToString()
            Obj.ObjEmpLedger.str_Class = ds.Tables(1).Rows(0)("Class").ToString()

            Obj.ObjEmpLedger.str_Description = ds.Tables(1).Rows(0)("Description").ToString()
            Obj.ObjEmpLedger.str_Type = ds.Tables(1).Rows(0)("LedgerType").ToString()
            Obj.ObjEmpLedger.str_AccountCode1 = ds.Tables(1).Rows(0)("AccountCode1").ToString()
            Obj.ObjEmpLedger.str_AccountCode2 = ds.Tables(1).Rows(0)("AccountCode2").ToString()
            Obj.ObjEmpLedger.str_LedgerComment = ds.Tables(1).Rows(0)("Comment").ToString()
            Obj.ObjEmpLedger.str_ParentAccount = ds.Tables(1).Rows(0)("ParentAccount").ToString()
            Obj.ObjEmpMain.dt_SpouseDatails = ds.Tables(2)

            If ds.Tables(3).Rows.Count > 0 Then
                Obj.ObjEmpMain.str_PreStreet = ds.Tables(3).Rows(0)("PreStreet").ToString()
                Obj.ObjEmpMain.str_PreArea = ds.Tables(3).Rows(0)("PreArea").ToString()
                Obj.ObjEmpMain.str_PreCity = ds.Tables(3).Rows(0)("PreTown").ToString()
                Obj.ObjEmpMain.str_PreState = ds.Tables(3).Rows(0)("PreState").ToString()
                Obj.ObjEmpMain.str_PrePincode = ds.Tables(3).Rows(0)("PrePincode").ToString()
                Obj.ObjEmpMain.str_PreCountry = ds.Tables(3).Rows(0)("PreCountry").ToString()
                Obj.ObjEmpMain.str_PreTel = ds.Tables(3).Rows(0)("PreTel").ToString()
                Obj.ObjEmpMain.str_PreExtension = ds.Tables(3).Rows(0)("PreExtension").ToString()
                Obj.ObjEmpMain.str_PerMobile1 = ds.Tables(3).Rows(0)("PreMobile1").ToString()
                Obj.ObjEmpMain.str_PerEmail1 = ds.Tables(3).Rows(0)("PreEmail1").ToString()

                Obj.ObjEmpMain.str_PerStreet = ds.Tables(3).Rows(0)("PerStreet").ToString()
                Obj.ObjEmpMain.str_PerArea = ds.Tables(3).Rows(0)("PerArea").ToString()
                Obj.ObjEmpMain.str_PerCity = ds.Tables(3).Rows(0)("PerTown").ToString()
                Obj.ObjEmpMain.str_PerState = ds.Tables(3).Rows(0)("PerState").ToString()
                Obj.ObjEmpMain.str_PerPincode = ds.Tables(3).Rows(0)("PerPincode").ToString()
                Obj.ObjEmpMain.str_PerCountry = ds.Tables(3).Rows(0)("PerCountry").ToString()
                Obj.ObjEmpMain.str_PerTel = ds.Tables(3).Rows(0)("PerTel").ToString()
                Obj.ObjEmpMain.str_PerExtension = ds.Tables(3).Rows(0)("PerExtension").ToString()
                Obj.ObjEmpMain.str_PerMobile2 = ds.Tables(3).Rows(0)("PreMobile2").ToString()
                Obj.ObjEmpMain.str_PerEmail2 = ds.Tables(3).Rows(0)("PreEmail2").ToString()

                Obj.objSTSEmpDetails.TimeZone = ds.Tables(3).Rows(0)("TimeZoneID").ToString()
                Obj.objSTSEmpDetails.GradeID = ds.Tables(3).Rows(0)("GradeID").ToString()
            End If
            If ds.Tables(4).Rows.Count > 0 Then
                Obj.objSTSEmpDetails.AuthendicationID = ds.Tables(4).Rows(0)("AuthGroupID").ToString()
            End If

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Sub Get_EmployeeDetails(ByVal _StrSiteID As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByVal _LedgerID As String, ByRef _dtEmployeeDetails As DataTable, ByVal _Flag As String, ByVal int_BusinessPeroidID As Integer)
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetEmployeeMasterDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", _StrSiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", int_BusinessPeroidID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", _Flag)
            BaseConn.cmd.Parameters.AddWithValue("@EmpID", _LedgerID)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            dt = New DataTable
            BaseConn.da.Fill(dt)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            BaseConn.Close()
        End Try
        _dtEmployeeDetails = dt
    End Sub
    Public Function Update_EmpMaster(ByVal obj As Classes.csEmployeeMaster, ByRef EmpID As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("EmployeeMasterUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID) 'obj.int_BusinessPeriodID)

            BaseConn.cmd.Parameters.AddWithValue("@EmpID", obj.ObjEmpMain.str_EmpID)
            BaseConn.cmd.Parameters.AddWithValue("@FirstName", obj.ObjEmpMain.str_FirstName)
            BaseConn.cmd.Parameters.AddWithValue("@LastName", obj.ObjEmpMain.str_LastName)
            BaseConn.cmd.Parameters.AddWithValue("@AliasName1", obj.ObjEmpMain.str_AliasName1)
            BaseConn.cmd.Parameters.AddWithValue("@AliasName2", obj.ObjEmpMain.str_AliasName2)
            BaseConn.cmd.Parameters.AddWithValue("@JoiningDate", obj.ObjEmpMain.dtp_JoiningDate)
            BaseConn.cmd.Parameters.AddWithValue("@Designation", obj.ObjEmpMain.str_Designation)
            BaseConn.cmd.Parameters.AddWithValue("@Category", obj.ObjEmpMain.str_Category)
            BaseConn.cmd.Parameters.AddWithValue("@SubCategory", obj.ObjEmpMain.str_SubCategory)
            BaseConn.cmd.Parameters.AddWithValue("@Department", obj.ObjEmpMain.str_Department)
            BaseConn.cmd.Parameters.AddWithValue("@Nationality", obj.ObjEmpMain.str_Nationality)
            BaseConn.cmd.Parameters.AddWithValue("@Language", obj.ObjEmpMain.str_Language)
            BaseConn.cmd.Parameters.AddWithValue("@SellFlag", obj.ObjEmpMain.bool_SellFlag)
            BaseConn.cmd.Parameters.AddWithValue("@StatusTech", obj.ObjEmpMain.bool_StatusTech)
            BaseConn.cmd.Parameters.AddWithValue("@EmpComment", obj.ObjEmpMain.str_Comment)
            BaseConn.cmd.Parameters.AddWithValue("@DOB", obj.ObjEmpMain.dtp_DOB)
            BaseConn.cmd.Parameters.AddWithValue("@ICE1Name", obj.ObjEmpMain.str_ICE1Name)
            BaseConn.cmd.Parameters.AddWithValue("@ICE1No", obj.ObjEmpMain.str_ICE1No)
            BaseConn.cmd.Parameters.AddWithValue("@ICE1Comment", obj.ObjEmpMain.str_ICE1Comment)
            BaseConn.cmd.Parameters.AddWithValue("@ICE2Name", obj.ObjEmpMain.str_ICE2Name)
            BaseConn.cmd.Parameters.AddWithValue("@ICE2No", obj.ObjEmpMain.str_ICE2No)
            BaseConn.cmd.Parameters.AddWithValue("@ICE2Comment", obj.ObjEmpMain.str_ICE2Comment)
            BaseConn.cmd.Parameters.AddWithValue("@PassageAmount", obj.ObjEmpMain.dbl_PassageAmount)
            BaseConn.cmd.Parameters.AddWithValue("@SyncID1", obj.ObjEmpMain.str_syncID1)
            BaseConn.cmd.Parameters.AddWithValue("@Gender", obj.ObjEmpMain.intGender)
            BaseConn.cmd.Parameters.AddWithValue("@EOSType", obj.ObjEmpMain.int_EOSType)
            BaseConn.cmd.Parameters.AddWithValue("@FamilyName", obj.ObjEmpMain.str_FamilyName)
            BaseConn.cmd.Parameters.AddWithValue("@MaritalStatus", obj.ObjEmpMain.str_MaritalStatus)
            BaseConn.cmd.Parameters.AddWithValue("@BloodGroup", obj.ObjEmpMain.str_BloodGroup)
            BaseConn.cmd.Parameters.AddWithValue("@Desc1", obj.ObjEmpMain.str_Desc1)
            BaseConn.cmd.Parameters.AddWithValue("@Desc2", obj.ObjEmpMain.str_Desc2)
            BaseConn.cmd.Parameters.AddWithValue("@Desc3", obj.ObjEmpMain.str_Desc3)
            BaseConn.cmd.Parameters.AddWithValue("@Desc4", obj.ObjEmpMain.str_Desc4)
            BaseConn.cmd.Parameters.AddWithValue("@Payable", obj.ObjEmpMain.bool_Payable)

            If obj.ObjEmpMain.img_Photo Is Nothing Then
                Dim photoParam As New SqlParameter("@Photo", SqlDbType.Image)
                photoParam.Value = DBNull.Value
                BaseConn.cmd.Parameters.Add(photoParam)
            Else
                BaseConn.cmd.Parameters.AddWithValue("@Photo", obj.ObjEmpMain.img_Photo) 'PHOTO
            End If
          
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.ObjEmpMain.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.ObjEmpMain.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.ObjEmpMain.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.ObjEmpMain.dtp_LastUpdatedDate)
            '-----------------------DT
            '----------Ledger---------------
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.ObjEmpLedger.int_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@Class", obj.ObjEmpLedger.str_Catagory)
            BaseConn.cmd.Parameters.AddWithValue("@Description", obj.ObjEmpLedger.str_Description)
            BaseConn.cmd.Parameters.AddWithValue("@AccountType1", obj.ObjEmpLedger.str_AccountCode1)
            BaseConn.cmd.Parameters.AddWithValue("@AccountType2", obj.ObjEmpLedger.str_AccountCode2)
            BaseConn.cmd.Parameters.AddWithValue("@Type", obj.ObjEmpLedger.str_Type)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerComment", obj.ObjEmpLedger.str_LedgerComment)
            BaseConn.cmd.Parameters.AddWithValue("@ParentAccount", obj.ObjEmpLedger.str_ParentAccount)
            BaseConn.cmd.Parameters.AddWithValue("@Status", obj.ObjEmpLedger.bool_InActive)
            BaseConn.cmd.Parameters.AddWithValue("@ReadOnly", obj.ObjEmpLedger.int_ReadOnly)

            BaseConn.cmd.Parameters.AddWithValue("@SpouseDetails", obj.ObjEmpMain.dt_SpouseDatails)

            '----------EmployeeMasterSub---------------
            BaseConn.cmd.Parameters.AddWithValue("@PreStreet", obj.ObjEmpMain.str_PreStreet)
            BaseConn.cmd.Parameters.AddWithValue("@PreArea", obj.ObjEmpMain.str_PreArea)
            BaseConn.cmd.Parameters.AddWithValue("@PreCity", obj.ObjEmpMain.str_PreCity)
            BaseConn.cmd.Parameters.AddWithValue("@PreState", obj.ObjEmpMain.str_PreState)
            BaseConn.cmd.Parameters.AddWithValue("@PrePincode", obj.ObjEmpMain.str_PrePincode)
            BaseConn.cmd.Parameters.AddWithValue("@PreCountry", obj.ObjEmpMain.str_PreCountry)
            BaseConn.cmd.Parameters.AddWithValue("@PreTel", obj.ObjEmpMain.str_PreTel)
            BaseConn.cmd.Parameters.AddWithValue("@PreExtension", obj.ObjEmpMain.str_PreExtension)
            BaseConn.cmd.Parameters.AddWithValue("@PerStreet", obj.ObjEmpMain.str_PerStreet)
            BaseConn.cmd.Parameters.AddWithValue("@PerArea", obj.ObjEmpMain.str_PerArea)
            BaseConn.cmd.Parameters.AddWithValue("@PerCity", obj.ObjEmpMain.str_PerCity)
            BaseConn.cmd.Parameters.AddWithValue("@PerPincode", obj.ObjEmpMain.str_PerPincode)
            BaseConn.cmd.Parameters.AddWithValue("@PerCountry", obj.ObjEmpMain.str_PerCountry)
            BaseConn.cmd.Parameters.AddWithValue("@PerTel", obj.ObjEmpMain.str_PerTel)
            BaseConn.cmd.Parameters.AddWithValue("@PerExtension", obj.ObjEmpMain.str_PerExtension)
            BaseConn.cmd.Parameters.AddWithValue("@PerMobile1", obj.ObjEmpMain.str_PerMobile1)
            BaseConn.cmd.Parameters.AddWithValue("@PerMobile2", obj.ObjEmpMain.str_PerMobile2)
            BaseConn.cmd.Parameters.AddWithValue("@PerEmail1", obj.ObjEmpMain.str_PerEmail1)
            BaseConn.cmd.Parameters.AddWithValue("@PerEmail2", obj.ObjEmpMain.str_PerEmail2)
            BaseConn.cmd.Parameters.AddWithValue("@PerState", obj.ObjEmpMain.str_PerState)
            BaseConn.cmd.Parameters.AddWithValue("@TimeZone", obj.objSTSEmpDetails.TimeZone)
            BaseConn.cmd.Parameters.AddWithValue("@GradeID", obj.objSTSEmpDetails.GradeID)
            BaseConn.cmd.Parameters.AddWithValue("@BranchName", obj.ObjEmpMain.str_BranchName)
            BaseConn.cmd.Parameters.AddWithValue("@BankAccName", obj.ObjEmpMain.str_BankAccName)
            BaseConn.cmd.Parameters.AddWithValue("@BankAccType", obj.ObjEmpMain.str_BankAccType)
            BaseConn.cmd.Parameters.AddWithValue("@Town", obj.ObjEmpMain.str_Town)
            BaseConn.cmd.Parameters.AddWithValue("@CostPerHr", obj.ObjEmpMain.int_CostPerHr)
            BaseConn.cmd.Parameters.AddWithValue("@Password", obj.ObjEmpMain.str_Password)
            BaseConn.cmd.Parameters.AddWithValue("@AuthGrp", obj.objSTSEmpDetails.AuthendicationID)
            BaseConn.cmd.Parameters.AddWithValue("@GroupID", obj.ObjEmpMain.str_GroupID)
            BaseConn.cmd.Parameters.AddWithValue("@CreateLogin", obj.ObjEmpMain.str_CreateLogin)
            BaseConn.cmd.Parameters.AddWithValue("@SMSFlag", obj.ObjEmpMain.bool_SendSMS)
            BaseConn.cmd.Parameters.AddWithValue("@EmailFlag", obj.ObjEmpMain.bool_SendEmail)

            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjEmpCommon.str_Flag)
            BaseConn.cmd.Parameters.Add("@EmpIDOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 500).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            EmpID = BaseConn.cmd.Parameters("@EmpIDOut").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.int_BusinessPeriodID, obj.ObjEmpMain.str_CreatedBy, obj.ObjEmpMain.dtp_CreatedDate, "", "EmployeeMaster", Err.Number, "Error in " & obj.ObjEmpCommon.str_Flag & " : " & obj.ObjEmpMain.str_EmpID & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = 1
        Finally
            BaseConn.Close()
        End Try

        Update_EmpMaster = _ErrString
    End Function
    Public Function Update_EmpAccounts(ByVal obj As Classes.csEmployeeMaster, ByRef EmpID As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("EmpAccountsUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", obj.str_SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID) 'obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.ObjEmpMain.str_EmpID)
            BaseConn.cmd.Parameters.AddWithValue("@JoiningDate", obj.ObjEmpMain.dtp_JoiningDate)
            BaseConn.cmd.Parameters.AddWithValue("@BankACNo", obj.ObjEmpMain.str_BankACNo)
            BaseConn.cmd.Parameters.AddWithValue("@IBan", obj.ObjEmpMain.str_IBan)
            BaseConn.cmd.Parameters.AddWithValue("@BankName", obj.ObjEmpMain.str_BankName)
            BaseConn.cmd.Parameters.AddWithValue("@BeneficiaryCode", obj.ObjEmpMain.str_BeneficiaryCode)
            BaseConn.cmd.Parameters.AddWithValue("@ChequePrintName", obj.ObjEmpMain.str_ChequePrintName)
            BaseConn.cmd.Parameters.AddWithValue("@MonthlyDeductable", obj.ObjEmpMain.dbl_MonthlyDeductable)
            BaseConn.cmd.Parameters.AddWithValue("@WPSID", obj.ObjEmpWPSDetails.str_WPSID)
            BaseConn.cmd.Parameters.AddWithValue("@WPSType", obj.ObjEmpWPSDetails.str_WPSType)
            BaseConn.cmd.Parameters.AddWithValue("@WPSRoutingCode", obj.ObjEmpWPSDetails.str_WPSRoutingCode)
            BaseConn.cmd.Parameters.AddWithValue("@DeductMonth", obj.ObjEmpMain.int_DeductMonth)
            BaseConn.cmd.Parameters.AddWithValue("@EmpHRDetailsDT", obj.ObjEmpHR.dt_EmpHR)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.ObjEmpMain.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.ObjEmpMain.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.ObjEmpMain.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.ObjEmpMain.dtp_LastUpdatedDate)

            BaseConn.cmd.Parameters.Add("@EmpIDOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            EmpID = BaseConn.cmd.Parameters("@EmpIDOut").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.int_BusinessPeriodID, obj.ObjEmpMain.str_CreatedBy, obj.ObjEmpMain.dtp_CreatedDate, "", "EmpAccounts", Err.Number, "Error in " & obj.ObjEmpCommon.str_Flag & " : " & obj.ObjEmpMain.str_EmpID & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
        Update_EmpAccounts = _ErrString
    End Function
    Public Sub UpdateEmpAccountsFromExcel(ByVal str_SiteID As String, ByVal _strPath As String, ByVal _strPWD As String, ByVal _BusinessPeriodID As Integer, _
                              ByVal _LoggedUser As String, ByVal dt_EmpHR As DataTable, ByRef _ErrNo As Integer, ByRef _ErrString As String)
        _ErrNo = 0
        _ErrString = ""
        Try
            BaseConn.Open(_strPath, _strPWD)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_ImportEmpAccountsFromExcel]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", _BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@EmpHRDetailsDT", dt_EmpHR)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", _LoggedUser)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", Date.Now)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", _LoggedUser)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", Date.Now)
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
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
    Public Sub GetEmpAccounts(ByRef Obj As Classes.csEmployeeMaster, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[GetEmpAccountsDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@CID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", Obj.ObjEmpMain.str_EmpID)

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            Obj.ObjEmpMain.str_BankACNo = ds.Tables(0).Rows(0)("BankACNo").ToString()
            Obj.ObjEmpMain.str_IBan = ds.Tables(0).Rows(0)("IBan").ToString()
            Obj.ObjEmpMain.str_BankName = ds.Tables(0).Rows(0)("BankName").ToString()
            Obj.ObjEmpMain.str_BeneficiaryCode = ds.Tables(0).Rows(0)("BeneficiaryCode").ToString()
            Obj.ObjEmpMain.str_ChequePrintName = ds.Tables(0).Rows(0)("ChequeName").ToString()
            Obj.ObjEmpMain.dbl_MonthlyDeductable = ds.Tables(0).Rows(0)("MonthlyDeductable").ToString()

            Obj.ObjEmpWPSDetails.str_WPSType = ds.Tables(0).Rows(0)("WPSType").ToString()
            Obj.ObjEmpWPSDetails.str_WPSRoutingCode = ds.Tables(0).Rows(0)("WPSRoutingCode").ToString()
            Obj.ObjEmpWPSDetails.str_WPSID = ds.Tables(0).Rows(0)("WPSID").ToString()

            'If ds.Tables(1).Rows.Count > 0 Then
            Obj.ObjEmpHR.dt_EmpHR = ds.Tables(1)
            'Else
            '    Obj.ObjEmpHR.dt_EmpHR = ds.Tables(2)
            'End If
            Obj.ObjEmpCommon.int_count = ds.Tables(2).Rows(0)("HikeCount")
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Function Update_EmpDocumentControl(ByVal obj As Classes.csEmployeeMaster, ByRef EmpID As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("sp_EmpDocumentControlUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", obj.str_SiteID) 'obj.str_SiteID
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID) 'obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.ObjEmpMain.str_EmpID)
            BaseConn.cmd.Parameters.AddWithValue("@EmpDocumentDetailsDT", obj.ObjEmpDocument.dt_EmpDocumnet)
            

            BaseConn.cmd.Parameters.Add("@EmpIDOut", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            EmpID = BaseConn.cmd.Parameters("@EmpIDOut").Value.ToString
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.int_BusinessPeriodID, obj.ObjEmpMain.str_CreatedBy, obj.ObjEmpMain.dtp_CreatedDate, "", "EmpDocumentControl", Err.Number, "Error in " & obj.ObjEmpCommon.str_Flag & " : " & obj.ObjEmpMain.str_EmpID & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try
        Update_EmpDocumentControl = _ErrString
    End Function

    Public Sub GetEmpDocumentControl(ByRef Obj As Classes.csEmployeeMaster, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetEmpDocumentControlDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", Obj.ObjEmpMain.str_EmpID)

            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                Obj.ObjEmpDocument.dt_EmpDocumnet = ds.Tables(0)
            End If
        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub
    Public Sub Get_StructureEmpHrDetils(ByRef Obj As Classes.csEmployeeMaster, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer, ByRef ErrStr As String)
        ErrNo = 0
        ErrStr = ""
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("[sp_GetSalesmanSettingsDetails]", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            'BaseConn.cmd.Parameters.AddWithValue("@EmpID", Obj.ObjEmpMain.str_EmpID)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", Obj.ObjEmpLedger.int_LedgerID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", Obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", Obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", Obj.ObjEmpCommon.str_Flag)
            BaseConn.da = New SqlClient.SqlDataAdapter(BaseConn.cmd)
            Dim ds As New DataSet
            BaseConn.da.Fill(ds)

            Obj.ObjEmpMain.str_EmpID = ds.Tables(0).Rows(0)("EmpID").ToString()
            Obj.ObjEmpMain.dbl_CarryFwd = ds.Tables(0).Rows(0)("CarryFwd").ToString()
            Obj.ObjEmpMain.dbl_AnnualLeave = ds.Tables(0).Rows(0)("AnnualLeave").ToString()
            Obj.ObjEmpMain.dbl_AnnualSickLeave = ds.Tables(0).Rows(0)("AnnualSickLeave").ToString()
            Obj.ObjEmpMain.dbl_Lieu = ds.Tables(0).Rows(0)("Lieu").ToString()
            Obj.ObjEmpMain.dbl_TakenLeave = ds.Tables(0).Rows(0)("TakenLeave").ToString()
            Obj.ObjEmpMain.dbl_SickLeavePaid = ds.Tables(0).Rows(0)("SickLeavePaid").ToString()
            Obj.ObjEmpMain.dbl_OtherPaidLeave = ds.Tables(0).Rows(0)("OtherPaidLeave").ToString()
            Obj.ObjEmpMain.dbl_UnPaidLeave = ds.Tables(0).Rows(0)("UnPaidLeave").ToString()
            Obj.ObjEmpMain.dbl_TotalAvailableLeave = ds.Tables(0).Rows(0)("TotalAvailableLeave").ToString()
            Obj.ObjEmpMain.dbl_TotalTakenLeave = ds.Tables(0).Rows(0)("TotalTakenLeave").ToString()
            Obj.ObjEmpMain.dbl_RemainingLeave = ds.Tables(0).Rows(0)("RemainingLeave").ToString()

        Catch ex As Exception
            ErrNo = 1
            ErrStr = ex.Message
        Finally
            BaseConn.Close()
        End Try
    End Sub

    Public Function Update_EmpHr(ByVal obj As Classes.csEmployeeMaster, ByRef EmpID As String, ByVal _StrDBPath As String, ByVal _StrDBPwd As String, ByRef ErrNo As Integer) As String
        Dim _ErrString As String = ""
        ErrNo = 0
        Try
            BaseConn.Open(_StrDBPath, _StrDBPwd)
            BaseConn.cmd = New SqlClient.SqlCommand("sp_EmpHrUpdate", BaseConn.cnn)
            BaseConn.cmd.CommandType = CommandType.StoredProcedure
            BaseConn.cmd.Parameters.AddWithValue("@SiteID", obj.str_SiteID)
            BaseConn.cmd.Parameters.AddWithValue("@BusinessPeriodID", obj.int_BusinessPeriodID)
            BaseConn.cmd.Parameters.AddWithValue("@Flag", obj.ObjEmpCommon.str_Flag)
            BaseConn.cmd.Parameters.AddWithValue("@LedgerID", obj.ObjEmpLedger.int_LedgerID)

            BaseConn.cmd.Parameters.AddWithValue("@CarryFwd", obj.ObjEmpMain.dbl_CarryFwd)
            BaseConn.cmd.Parameters.AddWithValue("@AnnualLeave", obj.ObjEmpMain.dbl_AnnualLeave)
            BaseConn.cmd.Parameters.AddWithValue("@AnnualSickLeave", obj.ObjEmpMain.dbl_AnnualSickLeave)
            BaseConn.cmd.Parameters.AddWithValue("@Lieu", obj.ObjEmpMain.dbl_Lieu)
            BaseConn.cmd.Parameters.AddWithValue("@TakenLeave", obj.ObjEmpMain.dbl_TakenLeave)
            BaseConn.cmd.Parameters.AddWithValue("@SickLeavePaid", obj.ObjEmpMain.dbl_SickLeavePaid)
            BaseConn.cmd.Parameters.AddWithValue("@OtherPaidLeave", obj.ObjEmpMain.dbl_OtherPaidLeave)
            BaseConn.cmd.Parameters.AddWithValue("@UnPaidLeave", obj.ObjEmpMain.dbl_UnPaidLeave)
            BaseConn.cmd.Parameters.AddWithValue("@TotalAvailableLeave", obj.ObjEmpMain.dbl_TotalAvailableLeave)
            BaseConn.cmd.Parameters.AddWithValue("@TotalTakenLeave", obj.ObjEmpMain.dbl_TotalTakenLeave)
            BaseConn.cmd.Parameters.AddWithValue("@RemainingLeave", obj.ObjEmpMain.dbl_RemainingLeave)

            BaseConn.cmd.Parameters.AddWithValue("@CreatedBy", obj.ObjEmpMain.str_CreatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@CreatedDate", obj.ObjEmpMain.dtp_CreatedDate)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedBy", obj.ObjEmpMain.str_LastUpdatedBy)
            BaseConn.cmd.Parameters.AddWithValue("@LastUpdatedDate", obj.ObjEmpMain.dtp_LastUpdatedDate)

            BaseConn.cmd.Parameters.AddWithValue("@ERRORNO", SqlDbType.Int).Direction = ParameterDirection.Output
            BaseConn.cmd.Parameters.Add("@ERRORDESC", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output
            BaseConn.cmd.ExecuteNonQuery()
            ErrNo = BaseConn.cmd.Parameters("@ERRORNO").Value.ToString
            _ErrString = BaseConn.cmd.Parameters("@ERRORDESC").Value.ToString
        Catch ex As Exception
            _ErrString = ex.Message
            ObjDalGeneral = New DAL_General(obj.str_SiteID)
            ObjDalGeneral.Elog_Insert(obj.str_SiteID, _StrDBPath, _StrDBPwd, obj.int_BusinessPeriodID, obj.ObjEmpMain.str_CreatedBy, obj.ObjEmpMain.dtp_CreatedDate, "", "Hr", Err.Number, "Error in " & obj.ObjEmpCommon.str_Flag & " : " & obj.ObjEmpMain.str_EmpID & "", ex.Message, 5, 3, 1, ErrNo)
            ErrNo = Err.Number
        Finally
            BaseConn.Close()
        End Try

        Update_EmpHr = _ErrString
    End Function

End Class
