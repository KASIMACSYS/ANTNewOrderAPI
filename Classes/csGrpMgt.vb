'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================


Public Class csGrpMgt
    Public str_SiteID As String
    Public str_Restriction As String
    Public objcsGrpMgtMain As New csGrpMgtMain
    Public objcsGrpMgtSub As New csGrpMgtSub
    Public objcsMenuMgtMain As New csMenuMgtMain
    Public objcsMenuMgtSub As New csMenuMgtSub
    Public objcsGrpMgtGS As New csGroupMgtGeneralSettings
    Public objcsGrpMgtCommon As New csGrpMgtCommon
End Class

Public Class csGrpMgtMain
    Public GroupID As String
    Public GroupName As String
    Public CreatedDate As Date
    Public CreatedBy As String
    Public ModifiedDate As Date
    Public ModifiedBy As String
    Public GroupSiteID As String
    Public GroupLevel As String
End Class

Public Class csGrpMgtSub
    Public GroupID As String
    Public MenuID As String
    Public Options As String
    Public Favorite As Boolean
End Class

Public Class csMenuMgtMain
    Public MenuID As String
    Public MenuGroup As String
    Public DefaultText As String
    Public CustomText As String
End Class

Public Class csMenuMgtSub
    Public MenuId As String
    Public Options As String
    Public Id As Integer
End Class

Public Class csGroupMgtGeneralSettings
    Public dt_GroupMgtgeneralSettings
End Class

Public Class csGrpMgtCommon
    Public str_Flag As String
    Public dt_GrpMgtfilemain
    Public dt_GrpMgtfilesub
    Public dt_GrpMgtMastermain
    Public dt_GrpMgtMastersub
    Public dt_GrpMgtSalesMain
    Public dt_GrpMgtSalesSub
    Public dt_GrpMgtPurchaseMain
    Public dt_GrpMgtPurchaseSub
    Public dt_GrpMgtManufacturingMain
    Public dt_GrpMgtManufacturingSub
    Public dt_GrpMgtAccountsMain
    Public dt_GrpMgtAccountsSub
    Public dt_GrpMgtHRPayrollMain
    Public dt_GrpMgtHRPayrollSub
    Public dt_GrpMgtReportMain
    Public dt_GrpMgtReportSub
    Public dt_GrpMgtInventoryMain
    Public dt_GrpMgtInventorySub
    Public dt_GrpMgtAssetMain
    Public dt_GrpMgtAssetSub

    Public dt_GrpMgtfilemain_Report
    Public dt_GrpMgtfilesub_Report
    Public dt_GrpMgtMastermain_Report
    Public dt_GrpMgtMastersub_Report
    Public dt_GrpMgtSalesMain_Report
    Public dt_GrpMgtSalesSub_Report
    Public dt_GrpMgtPurchaseMain_Report
    Public dt_GrpMgtPurchaseSub_Report
    Public dt_GrpMgtManufacturingMain_Report
    Public dt_GrpMgtManufacturingSub_Report
    Public dt_GrpMgtAccountsMain_Report
    Public dt_GrpMgtAccountsSub_Report
    Public dt_GrpMgtHRPayrollMain_Report
    Public dt_GrpMgtHRPayrollSub_Report
    Public dt_GrpMgtReportMain_Report
    Public dt_GrpMgtReportSub_Report
    Public dt_GrpMgtInventoryMain_Report
    Public dt_GrpMgtInventorySub_Report
    Public dt_GrpMgtAssetMain_Report
    Public dt_GrpMgtAssetSub_Report

    Public str_FileFrom As String
    Public str_FileTo As String
    Public dt_GrpMgtall
    Public dt_GrpMgtall_Report
    Public dt_AllSiteID
    Public dt_groupaccpermission
    Public dt_GrpMgtSalesMan As DataTable
End Class
