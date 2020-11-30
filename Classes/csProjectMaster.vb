'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Public Class csProjectMaster
    Public str_SiteID As String
    Public int_BusinessPeriodID As Integer
    Public ObjProject As New csProject
    Public ObjLocation As New csLocation
    Public ObjProjCommon As New csProjCommon
End Class

Public Class csProject
    Public str_ProjectID As String
    Public str_Description As String
    Public str_MerchantID As String
    Public bool_Status As Boolean
    Public dtp_StartDate As Date
    Public dtp_EndDate As Date
    Public dbl_BudgetAmount As Double
    Public int_EstimatedManDay As Integer
    Public str_Country As String
    Public str_State As String
    Public str_City As String
    Public str_ContactPerson As String
    Public str_ContactDesignatin As String
    Public str_ContactEmail As String
    Public str_ContactTelephone As String
    Public str_ContactMobile As String
    Public int_ProductID As Integer
    Public int_DayHours As Integer
    Public dt_LocationSub As New DataTable
    Public str_DstLedgerID As String
    Public int_PCCID As Integer
End Class

Public Class csLocation
    Public str_ProjectID As String
    Public str_ProjectLocation As String
    Public str_EditLocation As String
End Class

Public Class csProjCommon
    Public str_Flag As String
    Public str_ProjOrLocation As String
End Class