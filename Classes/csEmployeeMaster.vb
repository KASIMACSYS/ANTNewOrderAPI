'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Public Class csEmployeeMaster
    Inherits csSignature
    Public str_SiteID As String
    Public int_BusinessPeriodID As Integer
    Public ObjEmpMain As New csEmpMain
    Public ObjEmpLedger As New csEmpLedger
    Public ObjEmpDocument As New csEmpDocument
    Public ObjEmpHR As New csEmpHR
    Public ObjEmpCommon As New csEmpCommon
    Public ObjEmpWPSDetails As New csEmpWPS
    Public objproject As csProjectDetail
    Public objSTSEmpDetails As New csSTSEmpDetails
End Class

Public Class csEmpMain
    Inherits csEmployeeMasterSub
    Public str_EmpID As String
    Public str_FirstName As String
    Public str_LastName As String
    Public str_AliasName1 As String
    Public str_AliasName2 As String
    Public dtp_JoiningDate As Date
    Public str_Designation As String
    Public str_Category As String
    Public str_SubCategory As String
    Public str_Department As String
    Public str_Nationality As String
    Public str_Language As String
    Public bool_StatusTech As Boolean
    Public dbl_CarryFwd As Double
    Public dbl_AnnualLeave As Double
    Public dbl_AnnualSickLeave As Double
    Public dbl_Lieu As Double
    Public dbl_TakenLeave As Double
    Public dbl_SickLeavePaid As Double
    Public dbl_OtherPaidLeave As Double
    Public dbl_UnPaidLeave As Double
    Public dbl_TotalAvailableLeave As Double
    Public dbl_TotalTakenLeave As Double
    Public dbl_RemainingLeave As Double
    Public dtp_DOB As Date
    Public str_ICE1Name As String
    Public str_ICE1No As String
    Public str_ICE1Comment As String
    Public str_ICE2Name As String
    Public str_ICE2No As String
    Public str_ICE2Comment As String
    Public str_BankACNo As String = String.Empty
    Public str_BeneficiaryCode As String = String.Empty
    Public str_IBan As String = String.Empty
    Public str_BankName As String = String.Empty
    Public str_ChequePrintName As String = String.Empty
    Public dbl_MonthlyDeductable As Double
    Public bool_SellFlag As Boolean
    Public int_DeductMonth As Integer
    Public str_FamilyName As String
    Public str_MaritalStatus As String
    Public str_BloodGroup As String

    Public dbl_PaymentLimit As Double
    Public int_LimitStatus As Integer
    Public dbl_Commission As Double
    Public str_SalesManComment As String
    Public chk_Technician As Boolean
    Public str_Comment As String
    Public str_Desc1 As String
    Public str_Desc2 As String
    Public str_Desc3 As String
    Public str_Desc4 As String
    Public str_CreatedBy As String
    Public dtp_CreatedDate As Date
    Public str_LastUpdatedBy As String
    Public dtp_LastUpdatedDate As Date

    Public bool_Payable As Boolean
    Public img_Photo() As Byte
    Public dbl_PassageAmount As Double
    Public str_syncID1 As String

    Public intGender As Integer
    Public int_EOSType As Integer
    Public bool_SendEmail As Boolean
    Public bool_SendSMS As Boolean
    Public dt_SpouseDatails As DataTable
End Class

Public Class csEmpLedger
    Public int_LedgerID As Integer
    Public str_Class As String
    Public str_Catagory As String = "Employee"
    Public str_COA As String
    Public str_Description As String
    Public str_Type As String
    Public dbl_Amount As Double
    Public dbl_Advance As Double
    Public str_LedgerComment As String
    Public str_ParentAccount As String
    Public str_AccountCode1 As String
    Public str_AccountCode2 As String
    Public bool_InActive As Boolean
    Public int_ReadOnly As Integer
End Class

Public Class csEmpDocument
    Public dt_EmpDocumnet As DataTable
End Class

Public Class csEmpHR
    Public dt_EmpHR As DataTable
End Class

Public Class csEmpWPS
    Public str_WPSID As String
    Public str_WPSType As String
    Public str_WPSRoutingCode As String
End Class

Public Class csEmpCommon
    Public str_Flag As String
    Public dt_empadvdetails As DataTable
    Public int_count As Integer
End Class

Public Class csSTSEmpDetails
    Public TimeZone As String
    Public GradeID As String
    Public GroupName As String
    Public GroupID As String
    Public AuthendicationID As Integer

End Class