
Public Class csAsset
    Public str_SiteID As String
    Public int_BusinessPeriodID As Integer
    Public str_Flag As String
    Public str_Type As String

    Public str_AssetID As String
    Public str_Description As String
    Public str_AssetRefID As String
    Public int_RevNo As Integer
    Public int_AssetGroupCategory As Integer
    Public str_Category As String
    Public str_Status As String
    Public dbl_Qty As Double
    Public img_Photo() As Byte
    Public int_AssetLedgerID As Integer
    Public str_BarCodeNo As String
    Public int_AccDepLedgerID As Integer

    Public str_Manufacturer As String
    Public str_Model As String
    Public str_PartNo As String
    Public str_SerialNo As String
    Public str_Desc1 As String
    Public str_Desc2 As String
    Public str_Desc3 As String
    Public str_Desc4 As String
    Public str_Comment As String

    'Public dtp_Manufacturing As Date
    'Public dtp_Installation As Date
    'Public dtp_ExpiryDate As Date

    'Public int_LedgerID As Integer
    'Public str_ContactName As String
    'Public str_Name As String
    'Public str_Address As String
    'Public str_Mobile As String
    'Public str_Tel As String
    'Public str_Fax As String
    'Public str_Email As String

    'Public dtp_CapitalizedOn As Date

    Public dtp_AquisitionOn As Date
    Public dbl_AmountPosted As Double
    Public dbl_SalvageorScrapValues As Double
    Public str_AcqInvRef As String
    Public int_AcqLedgerID As Integer
    Public str_AcqComment As String

    Public str_DepriciationType As String
    Public dtp_StartDate As Date
    Public int_NoofYears As Integer
    Public dbl_DepriciationPercentage As Double
    'Public dtp_EndDate As Date
    Public str_DepComment As String
    Public int_DepLedgerID As Integer

    Public str_DisposalType As String
    Public str_DisposalInvRef As String
    Public dtp_DisposalDate As Date
    Public int_DisposalLedgerID As Integer
    Public int_SalesLedgerID As Integer
    Public dbl_SellingAmount As Double
    Public str_DisposalComment As String

    Public str_CreatedBy As String
    Public dtp_CreatedDate As Date
    Public str_LastUpdatedBy As String
    Public dtp_LastUpdatedDate As Date

    Public dt_Transaction As DataTable

    Public dt_FileUpload As DataTable
    Public str_MenuID As String
    Public str_FormPrefix As String

    Public XMLData As String


    Public dtp_ReValDate As Date
    Public int_ReValLedger As Integer
    Public dbl_ReValAmount As Double
    Public str_ReValComment As String

End Class

Public Class csTask
    Public int_CID As Integer
    Public int_BusinessPeriodID As Integer
    Public str_Flag As String

    Public str_TaskID As String
    Public str_AssetID As String
    Public str_TaskCategory As String
    Public str_Type As String
    Public str_Desc1 As String
    Public str_Desc2 As String
    Public str_Desc3 As String
    Public str_Desc4 As String
    Public dtp_StartDate As Date
    Public dtp_DueDate As Date
    Public str_Status As String
    Public int_PopUpdays As Integer
    Public bool_NotifyFlag As Boolean
    Public str_Comment As String
    Public str_MenuID As String
    Public dt_Task As DataTable
    Public bool_Option As Boolean
    Public dtp_PostingDate As Date

    Public dbl_Amount1 As Decimal
    Public dbl_Amount2 As Decimal

    Public str_CreatedBy As String
    Public dtp_CreatedDate As Date
    Public str_LastUpdatedBy As String
    Public dtp_LastUpdatedDate As Date
End Class
