'======================================================================================
'$Author: Meeran $
'$Rev: 674 $
'$Date: 2012-05-29 18:06:08 +0530 (Tue, 29 May 2012) $ 
'======================================================================================

'==================================================================================
'Slno   ChangeBy    Date        Description
'==================================================================================




''' <summary>
''' have created 2 single ton class
''' </summary>
''' <auther>KM1007</auther>
''' <remarks></remarks>

Public Class csUserDefaultsSingleTon
    Public Shared ObjUserDefaults As csUserDefaultsSingleTon
    Public Shared isObjCreated As Boolean

#Region "ClassVariables"
    Dim _UserID As Integer, _UserName As String = String.Empty, _PWD As String, _GroupID As String = String.Empty, _GroupName As String = String.Empty, _UserLoginSiteID As String = String.Empty, _SalesManID As String = String.Empty
    Dim _UserMainTheme As String, _UserSecondaryTheme As String
    Private _DTRemoteSiteByUser As New DataTable
    Private _DTRemoteSiteWithGroup As New DataTable
    Dim _LanguageCode As Integer

    Public Property LanguageCode() As Integer
        Get
            Return _LanguageCode
        End Get
        Set(ByVal value As Integer)
            _LanguageCode = value
        End Set
    End Property

    Public Property SalesManID() As String
        Get
            Return _SalesManID
        End Get
        Set(ByVal value As String)
            _SalesManID = value
        End Set
    End Property

    Public Property UserID() As Integer
        Get
            Return _UserID
        End Get
        Set(ByVal value As Integer)
            _UserID = value
        End Set
    End Property

    Public Property UserMainTheme() As String
        Get
            Return _UserMainTheme
        End Get
        Set(ByVal value As String)
            _UserMainTheme = value
        End Set
    End Property

    Public Property UserSecondaryTheme() As String
        Get
            Return _UserSecondaryTheme
        End Get
        Set(ByVal value As String)
            _UserSecondaryTheme = value
        End Set
    End Property

    Public Property UserName() As String
        Get
            Return _UserName
        End Get
        Set(ByVal value As String)
            _UserName = value
        End Set
    End Property

    Public Property GroupID() As String
        Get
            Return _GroupID
        End Get
        Set(ByVal value As String)
            _GroupID = value
        End Set
    End Property

    Public Property GroupName() As String
        Get
            Return _GroupName
        End Get
        Set(ByVal value As String)
            _GroupName = value
        End Set
    End Property

    Public Property UserLoginSiteID() As String
        Get
            Return _UserLoginSiteID
        End Get
        Set(ByVal value As String)
            _UserLoginSiteID = value
        End Set
    End Property

    Public Property DTRemoteSiteByUser() As DataTable
        Get
            Return _DTRemoteSiteByUser
        End Get
        Set(ByVal value As DataTable)
            _DTRemoteSiteByUser = value
        End Set
    End Property

    Public Property DTRemoteSiteWithGroup() As DataTable
        Get
            Return _DTRemoteSiteWithGroup
        End Get
        Set(ByVal value As DataTable)
            _DTRemoteSiteWithGroup = value
        End Set
    End Property

    Public Property Password() As String
        Get
            Return _PWD
        End Get
        Set(ByVal value As String)
            _PWD = value
        End Set
    End Property
#End Region

    Private Sub New()
        'Private constructor wont allow other user to create object directly
    End Sub

    Public Shared Function getLoggedUserDetails() As csUserDefaultsSingleTon
        If isObjCreated = False Then
            'First instance
            'return a new object
            ObjUserDefaults = New csUserDefaultsSingleTon
            isObjCreated = True
            Return ObjUserDefaults
        Else
            'future object request
            'return an already created object
            Return ObjUserDefaults
        End If
    End Function

    Public Shared Sub Reset()
        If isObjCreated Then
            isObjCreated = False
            ObjUserDefaults = Nothing
        End If
    End Sub

End Class


Public Class csSiteDefaultsSingleTon
    Public Shared ObjSiteDefaultsSingleTon As csSiteDefaultsSingleTon
    Public Shared isObjCreated As Boolean

    Private _SiteDefaults As New Dictionary(Of String, csSiteDefaults)

    Public Property SiteDefaults() As Dictionary(Of String, csSiteDefaults)
        Get
            Return _SiteDefaults
        End Get
        Set(ByVal value As Dictionary(Of String, csSiteDefaults))
            _SiteDefaults = value
        End Set
    End Property

    Public Enum Action
        ADD
        EDIT
        DELETE
        APPROVED
        PRICERCVD
    End Enum


    Private Sub New()
        'Private constructor wont allow other user to create object directly
    End Sub

    Public Shared Function getSiteDefaultsSingleTon() As csSiteDefaultsSingleTon
        If isObjCreated = False Then
            'First instance
            'return a new object
            ObjSiteDefaultsSingleTon = New csSiteDefaultsSingleTon
            isObjCreated = True
            Return ObjSiteDefaultsSingleTon
        Else
            'future object request
            'return an already created object
            Return ObjSiteDefaultsSingleTon
        End If
    End Function

    Public Shared Sub DestroySiteDefaultSingleTon()
        If isObjCreated Then
            isObjCreated = False
            'ObjSiteDefaultsSingleTon = Nothing
        End If
    End Sub
End Class


Public Class csSiteDefaults
    Private _BusinessStartDate As Date
    Private _CustomerSettings As New Dictionary(Of String, String)
    Private _MenuSettings As New Dictionary(Of String, node_struct)
    Private _DashboardMenuSettings As New Dictionary(Of String, node_struct)
    Private _LicenseDetails As New Dictionary(Of String, String)
    Private _SalesMan As New DataTable
    Private _DTBSPeriod As New DataTable
    Private _EligibleMenu As New DataTable

    Public Property BusinessStartDate() As Date
        Get
            Return _BusinessStartDate
        End Get
        Set(ByVal value As Date)
            _BusinessStartDate = value
        End Set
    End Property

    Public Property DTBSPeriod() As DataTable
        Get
            Return _DTBSPeriod
        End Get
        Set(ByVal value As DataTable)
            _DTBSPeriod = value
        End Set
    End Property


    Public Property LicenseDetails() As Dictionary(Of String, String)
        Get
            Return _LicenseDetails
        End Get
        Set(ByVal value As Dictionary(Of String, String))
            _LicenseDetails = value
        End Set
    End Property

    Public Property DashboardMenuSettings() As Dictionary(Of String, node_struct)
        Get
            Return _DashboardMenuSettings
        End Get
        Set(ByVal value As Dictionary(Of String, node_struct))
            _DashboardMenuSettings = value
        End Set
    End Property

    Public Property MenuSettings() As Dictionary(Of String, node_struct)
        Get
            Return _MenuSettings
        End Get
        Set(ByVal value As Dictionary(Of String, node_struct))
            _MenuSettings = value
        End Set
    End Property

    Public Property CustomerSettings() As Dictionary(Of String, String)
        Get
            Return _CustomerSettings
        End Get
        Set(ByVal value As Dictionary(Of String, String))
            _CustomerSettings = value
        End Set
    End Property

    Public Property SalesMan() As DataTable
        Get
            Return _SalesMan
        End Get
        Set(ByVal value As DataTable)
            _SalesMan = value
        End Set
    End Property

    Public Property EligibleMenu() As DataTable
        Get
            Return _EligibleMenu
        End Get
        Set(ByVal value As DataTable)
            _EligibleMenu = value
        End Set
    End Property





    Structure node_structLngToken
        Public MenuID As String
        Public Options1 As Hashtable
    End Structure

    Public _DictLanguageTokens As New Dictionary(Of String, node_structLngToken)

    Public Property LanguageTokens() As Dictionary(Of String, node_structLngToken)
        Get
            Return _DictLanguageTokens
        End Get
        Set(ByVal value As Dictionary(Of String, node_structLngToken))
            _DictLanguageTokens = value
        End Set
    End Property

    Structure node_struct
        Public MenuID As String
        Public MenuText As String
        Public ShortKey As String
        Public _Color As String
        Public Reserved As Boolean
        Public LoadGrpMgt As Boolean
        Public Parameters As DataTable
        Public Options1 As Hashtable
        Public Options As Hashtable
        Public MenuGroup As String
        Public Favorites As Boolean
    End Structure
End Class

Public Class csFormDefaults
    Private _VouPrefix As String = String.Empty
    Private _IsVouFrcFlag As Boolean = False
    Private _IsVouApproveFlag As Boolean = False
    Private _GridRowFoldCount As Integer = 0
    Private _StartSeq As Integer = 0
    Private _FormName As String = String.Empty
    Private _FormLabelSettings As DataTable
    Private _FormLabelCustomData As DataTable
    Private _ProjectEnable As Boolean = False
    Private _SizeMultiplyFactor As Double = 0.0
    Private _IsSearchByItemDesc As Integer
    Private _IsPrimaryQtyEditable As Boolean = False
    Private _ConvertFrom As String = String.Empty
    Private _IsStatusCancel As Boolean = False
    Private _UpdateInventory As Boolean = False
    Private _GeneralLedger As Boolean = False
    Private _CheckDirInv As Boolean = False
    Private _DefalutTab As String = String.Empty
    Private _CashTransaction As Boolean = False
    Private _CheckMinMaxQty As Boolean = False
    Private _DefaultAccLedger As Integer = 0
    Private _NegativeStockRule As String = String.Empty
    Private _CheckCreditLimitFlag As Boolean = False
    Private _MCCBSerachByName As Boolean = False
    Private _MCCBWildSearch As Boolean = False
    Private _LedgerDepartment As String = String.Empty
    Private _isRetentionFlagEnable As String = String.Empty
    Private _CheckCreditDaysFlag As Boolean = False
    Private _AllowEnquiryPrice As Boolean = False
    Private _CashLedger As String = String.Empty
    Private _UseBarCode As Boolean = False
    Private _NegativeStockConfig As String = String.Empty
    Private _ShowAllMerchantTxnInItemDialog As Boolean = False
    Private _PrintOnApprove As Boolean = False
    Private _PrintPermissionOnlyApprovedVoucher As Boolean = False
    Private _AllowExcessQty As Boolean = False
    Private _AllowUnderCost As Boolean = False
    Private _AllowPriceChange As Boolean = False
    Private _AllowAdditionalItem As Boolean = False
    Private _AllowLedgerGroup As Boolean = False
    Private _RestrictDateSelection As Boolean = False
    Private _AllowLessthanCostPrice As String = String.Empty
    Private _AllowLessthanMinPrice As String = String.Empty
    Private _AllowGreaterthanMaxPrice As String = String.Empty
    Private _TaxEnable As Boolean = False
    Private _FormType As String = String.Empty
    Private _TaxLedgerType As String = String.Empty
    Private _DiscountInServiceItem As Boolean = False
    Private _DisplayItemsTotal As Boolean = False
    Private _InvoiceLevelTax As Boolean
    Private _TaxAfterDiscount As Boolean
    Private _ShowItemTaxDetails As Boolean
    Private _ShowInvTaxDetails As Boolean
    Private _ShowTradeDiscDetails As Boolean
    Private _SalesmanVisibility As Boolean
    Private _CostCentre As Boolean = False
    Private _CheckMerchantExpiry As Boolean = False

    Public Property ShowAllMerchantTxnInItemDialog() As Boolean
        Get
            Return _ShowAllMerchantTxnInItemDialog
        End Get
        Set(ByVal value As Boolean)
            _ShowAllMerchantTxnInItemDialog = value
        End Set
    End Property

    Public Property isRetentionFlagEnable() As String
        Get
            Return _isRetentionFlagEnable
        End Get
        Set(ByVal value As String)
            _isRetentionFlagEnable = value
        End Set
    End Property

    Public Property FormName() As String
        Get
            Return _FormName
        End Get
        Set(ByVal value As String)
            _FormName = value
        End Set
    End Property

    Public Property Prefix() As String
        Get
            Return _VouPrefix
        End Get
        Set(ByVal value As String)
            _VouPrefix = value
        End Set
    End Property

    Public Property IsFrcFlag() As Boolean
        Get
            Return _IsVouFrcFlag
        End Get
        Set(ByVal value As Boolean)
            _IsVouFrcFlag = value
        End Set
    End Property

    Public Property IsApproveFlag() As Boolean
        Get
            Return _IsVouApproveFlag
        End Get
        Set(ByVal value As Boolean)
            _IsVouApproveFlag = value
        End Set
    End Property

    Public Property GridRowFoldCount() As Integer
        Get
            Return _GridRowFoldCount
        End Get
        Set(ByVal value As Integer)
            _GridRowFoldCount = value
        End Set
    End Property

    'Public Property StartSeq() As Integer
    '    Get
    '        Return _StartSeq
    '    End Get
    '    Set(ByVal value As Integer)
    '        _StartSeq = value
    '    End Set
    'End Property

    Public Property FormLabelSettings() As DataTable
        Get
            Return _FormLabelSettings
        End Get
        Set(ByVal value As DataTable)
            _FormLabelSettings = value
        End Set
    End Property

    Public Property FormLabelCustomData() As DataTable
        Get
            Return _FormLabelCustomData
        End Get
        Set(ByVal value As DataTable)
            _FormLabelCustomData = value
        End Set
    End Property

    Public Property ProjectEnable() As Boolean
        Get
            Return _ProjectEnable
        End Get
        Set(ByVal value As Boolean)
            _ProjectEnable = value
        End Set
    End Property


    Public Property SizeMultiplyFactor() As Double
        Get
            Return _SizeMultiplyFactor
        End Get
        Set(ByVal value As Double)
            _SizeMultiplyFactor = value
        End Set
    End Property

    Public Property IsSearchByItemDesc() As Integer
        Get
            Return _IsSearchByItemDesc
        End Get
        Set(ByVal value As Integer)
            _IsSearchByItemDesc = value
        End Set
    End Property

    Public Property IsPrimaryQtyEditable() As Boolean
        Get
            Return _IsPrimaryQtyEditable
        End Get
        Set(ByVal value As Boolean)
            _IsPrimaryQtyEditable = value
        End Set
    End Property
    Public Property ConvertFrom() As String
        Get
            Return _ConvertFrom
        End Get
        Set(ByVal value As String)
            _ConvertFrom = value
        End Set
    End Property
    Public Property IsStatusCancel() As Boolean
        Get
            Return _IsStatusCancel
        End Get
        Set(ByVal value As Boolean)
            _IsStatusCancel = value
        End Set
    End Property

    Public Property UpdateInventory() As Boolean
        Get
            Return _UpdateInventory
        End Get
        Set(ByVal value As Boolean)
            _UpdateInventory = value
        End Set
    End Property

    Public Property GeneralLedger() As Boolean
        Get
            Return _GeneralLedger
        End Get
        Set(ByVal value As Boolean)
            _GeneralLedger = value
        End Set
    End Property

    Public Property CheckDirInv() As Boolean
        Get
            Return _CheckDirInv
        End Get
        Set(ByVal value As Boolean)
            _CheckDirInv = value
        End Set
    End Property

    Public Property DefalutTab() As String
        Get
            Return _DefalutTab
        End Get
        Set(ByVal value As String)
            _DefalutTab = value
        End Set
    End Property

    Public Property CashTransaction() As Boolean
        Get
            Return _CashTransaction
        End Get
        Set(ByVal value As Boolean)
            _CashTransaction = value
        End Set
    End Property

    Public Property DefaultAccLedger() As Integer
        Get
            Return _DefaultAccLedger
        End Get
        Set(ByVal value As Integer)
            _DefaultAccLedger = value
        End Set
    End Property

    Public Property MCCBWildSearch() As Boolean
        Get
            Return _MCCBWildSearch
        End Get
        Set(ByVal value As Boolean)
            _MCCBWildSearch = value
        End Set
    End Property

    Public Property MCCBSerachByName() As Boolean
        Get
            Return _MCCBSerachByName
        End Get
        Set(ByVal value As Boolean)
            _MCCBSerachByName = value
        End Set
    End Property

    Public Property CheckCreditLimitFlag() As Boolean
        Get
            Return _CheckCreditLimitFlag
        End Get
        Set(ByVal value As Boolean)
            _CheckCreditLimitFlag = value
        End Set
    End Property
    Public Property CheckCreditDaysFlag() As Boolean
        Get
            Return _CheckCreditDaysFlag
        End Get
        Set(ByVal value As Boolean)
            _CheckCreditDaysFlag = value
        End Set
    End Property

    Public Property NegativeStockRule() As String
        Get
            Return _NegativeStockRule
        End Get
        Set(ByVal value As String)
            _NegativeStockRule = value
        End Set
    End Property

    Public Property CheckMinMaxQty() As Boolean
        Get
            Return _CheckMinMaxQty
        End Get
        Set(ByVal value As Boolean)
            _CheckMinMaxQty = value
        End Set
    End Property

    Public Property LedgerDepartment() As String
        Get
            Return _LedgerDepartment
        End Get
        Set(ByVal value As String)
            _LedgerDepartment = value
        End Set
    End Property

    Public Property AllowEnquiryPrice() As Boolean
        Get
            Return _AllowEnquiryPrice
        End Get
        Set(ByVal value As Boolean)
            _AllowEnquiryPrice = value
        End Set
    End Property

    Public Property CashLedger() As String
        Get
            Return _CashLedger
        End Get
        Set(ByVal value As String)
            _CashLedger = value
        End Set
    End Property

    Public Property UseBarCode() As Boolean
        Get
            Return _UseBarCode
        End Get
        Set(ByVal value As Boolean)
            _UseBarCode = value
        End Set
    End Property

    Public Property NegativeStockConfig() As String
        Get
            Return _NegativeStockConfig
        End Get
        Set(ByVal value As String)
            _NegativeStockConfig = value
        End Set
    End Property
    Public Property PrintOnApprove() As String
        Get
            Return _PrintOnApprove
        End Get
        Set(ByVal value As String)
            _PrintOnApprove = value
        End Set
    End Property

    Public Property PrintPermissionOnlyApprovedVoucher() As Boolean
        Get
            Return _PrintPermissionOnlyApprovedVoucher
        End Get
        Set(ByVal value As Boolean)
            _PrintPermissionOnlyApprovedVoucher = value
        End Set
    End Property
    Public Property AllowExcessQty() As Boolean
        Get
            Return _AllowExcessQty
        End Get
        Set(ByVal value As Boolean)
            _AllowExcessQty = value
        End Set
    End Property
    Public Property AllowUnderCost() As Boolean
        Get
            Return _AllowUnderCost
        End Get
        Set(ByVal value As Boolean)
            _AllowUnderCost = value
        End Set
    End Property
    Public Property AllowPriceChange() As Boolean
        Get
            Return _AllowPriceChange
        End Get
        Set(ByVal value As Boolean)
            _AllowPriceChange = value
        End Set
    End Property
    Public Property AllowAdditionalItem() As Boolean
        Get
            Return _AllowAdditionalItem
        End Get
        Set(ByVal value As Boolean)
            _AllowAdditionalItem = value
        End Set
    End Property
    Public Property AllowLessthanMinPrice() As String
        Get
            Return _AllowLessthanMinPrice
        End Get
        Set(ByVal value As String)
            _AllowLessthanMinPrice = value
        End Set
    End Property
    Public Property AllowGreaterthanMaxPrice() As String
        Get
            Return _AllowGreaterthanMaxPrice
        End Get
        Set(ByVal value As String)
            _AllowGreaterthanMaxPrice = value
        End Set
    End Property
    Public Property AllowLessthanCostPrice() As String
        Get
            Return _AllowLessthanCostPrice
        End Get
        Set(ByVal value As String)
            _AllowLessthanCostPrice = value
        End Set
    End Property
    Public Property TaxEnable() As Boolean
        Get
            Return _TaxEnable
        End Get
        Set(ByVal value As Boolean)
            _TaxEnable = value
        End Set
    End Property

    Public Property TaxLedgerType() As String
        Get
            Return _TaxLedgerType
        End Get
        Set(ByVal value As String)
            _TaxLedgerType = value
        End Set
    End Property

    Public Property FormType() As String
        Get
            Return _FormType
        End Get
        Set(ByVal value As String)
            _FormType = value
        End Set
    End Property

    Public Property DiscountInServiceItem() As Boolean
        Get
            Return _DiscountInServiceItem
        End Get
        Set(ByVal value As Boolean)
            _DiscountInServiceItem = value
        End Set
    End Property

    Public Property DisplayItemsTotal() As Boolean
        Get
            Return _DisplayItemsTotal
        End Get
        Set(ByVal value As Boolean)
            _DisplayItemsTotal = value
        End Set
    End Property

    Public Property AllowInvLevelTax() As Boolean
        Get
            Return _InvoiceLevelTax
        End Get
        Set(ByVal value As Boolean)
            _InvoiceLevelTax = value
        End Set
    End Property

    Public Property InvLevelTaxAfterDiscount() As Boolean
        Get
            Return _TaxAfterDiscount
        End Get
        Set(ByVal value As Boolean)
            _TaxAfterDiscount = value
        End Set
    End Property

    Public Property ShowItemTaxDetails() As Boolean
        Get
            Return _ShowItemTaxDetails
        End Get
        Set(ByVal value As Boolean)
            _ShowItemTaxDetails = value
        End Set
    End Property

    Public Property ShowInvTaxDetails() As Boolean
        Get
            Return _ShowInvTaxDetails
        End Get
        Set(ByVal value As Boolean)
            _ShowInvTaxDetails = value
        End Set
    End Property

    Public Property ShowTradeDiscDetails() As Boolean
        Get
            Return _ShowTradeDiscDetails
        End Get
        Set(ByVal value As Boolean)
            _ShowTradeDiscDetails = value
        End Set
    End Property

    Public Property SalesmanVisibility() As Boolean
        Get
            Return _SalesmanVisibility
        End Get
        Set(ByVal value As Boolean)
            _SalesmanVisibility = value
        End Set
    End Property
    Public Property CostCentre() As Boolean
        Get
            Return _CostCentre
        End Get
        Set(ByVal value As Boolean)
            _CostCentre = value
        End Set
    End Property

    Public Property CheckMerchantExpiry() As Boolean
        Get
            Return _CheckMerchantExpiry
        End Get
        Set(ByVal value As Boolean)
            _CheckMerchantExpiry = value
        End Set
    End Property

    Public Property AllowLedgerGroup() As Boolean
        Get
            Return _AllowLedgerGroup
        End Get
        Set(ByVal value As Boolean)
            _AllowLedgerGroup = value
        End Set
    End Property

    Public Property RestrictDateSelection() As Boolean
        Get
            Return _RestrictDateSelection
        End Get
        Set(ByVal value As Boolean)
            _RestrictDateSelection = value
        End Set
    End Property

End Class

Public Class csGlobalDefaults
    Public Shared ObjGlobalDefault As csGlobalDefaults
    Public Shared isObjCreated As Boolean
    Public DTAllSiteDetails As New DataTable
    Public DTAllowedSiteDetails As New DataTable
    'Public DTPOSConfig As New DataTable
    Public ObjPOSConfig As New csPOSConfig
    Public DTMenuMgt As New DataTable
    'Public g_DualPwdFlag As Boolean = False

    '---------------------------------------------------------------
    Public str_company_name As String = "AcSys-IT"
    Public str_version As String = "V3.0"
    Public str_website As String = "www.acSys-IT.com"
    Public str_Product_Name As String = "acSysERP"
    Public copyrights As String = "All Rights Reserved For AcSys-IT Solutions"
    Public g_CurrentDirectoryPath As String = System.IO.Directory.GetCurrentDirectory
    '-------------------------------------------------------------

    Public ClipboardSource As String
    Public ClipboardSourceNo As String
    Public CopyCurrencyCodeFrom As String
    Public CopySiteID As String
    Public dt_clipboard As DataTable
    Public dt_clipboardItemExtraDetails As DataTable
    Public bool_ShowOutLook As Boolean = False

    Public _strSessionID As String = String.Empty

    Public Shared Function getGlobalDefaults() As csGlobalDefaults
        If isObjCreated = False Then
            'First instance
            'return a new object
            ObjGlobalDefault = New csGlobalDefaults
            isObjCreated = True
            Return ObjGlobalDefault
        Else
            'future object request
            'return an already created object
            Return ObjGlobalDefault
        End If
    End Function


    Public Sub Clipboard_Copy(ByVal dt As DataTable)
        If dt.Rows.Count > 0 Then
            dt_clipboard = New DataTable
            Dim dl_clipboard As DataRow
            dt_clipboard.Columns.Add(New DataColumn("SortNo", GetType(Integer)))
            dt_clipboard.Columns.Add(New DataColumn("Slno", GetType(Integer)))
            dt_clipboard.Columns.Add(New DataColumn("BarCodeNo"))
            dt_clipboard.Columns.Add(New DataColumn("ItemCode"))
            dt_clipboard.Columns.Add(New DataColumn("ItemDesc1"))
            dt_clipboard.Columns.Add(New DataColumn("Alias1"))
            dt_clipboard.Columns.Add(New DataColumn("ItemDesc2"))
            dt_clipboard.Columns.Add(New DataColumn("Alias2"))
            dt_clipboard.Columns.Add(New DataColumn("Unit"))
            '' dt_clipboard.Columns.Add(New DataColumn("BaseUnit"))
            dt_clipboard.Columns.Add(New DataColumn("BaseUnit", System.Type.GetType("System.Double")))
            dt_clipboard.Columns.Add(New DataColumn("VouQty", System.Type.GetType("System.Double")))
            dt_clipboard.Columns.Add(New DataColumn("PrimaryQty", System.Type.GetType("System.Double")))
            dt_clipboard.Columns.Add(New DataColumn("Price", System.Type.GetType("System.Double")))
            dt_clipboard.Columns.Add(New DataColumn("BaseUnitPrice", System.Type.GetType("System.Double")))
            dt_clipboard.Columns.Add(New DataColumn("TCDiscountAmount", System.Type.GetType("System.Double")))
            dt_clipboard.Columns.Add(New DataColumn("Amount", System.Type.GetType("System.Double")))
            dt_clipboard.Columns.Add(New DataColumn("LCAmount", System.Type.GetType("System.Double")))
            dt_clipboard.Columns.Add(New DataColumn("Package", System.Type.GetType("System.Double")))
            dt_clipboard.Columns.Add(New DataColumn("Pieces", System.Type.GetType("System.Double")))
            dt_clipboard.Columns.Add(New DataColumn("DeliveredTotQty", System.Type.GetType("System.Double")))
            dt_clipboard.Columns.Add(New DataColumn("PartNo"))
            dt_clipboard.Columns.Add(New DataColumn("Comment"))
            dt_clipboard.Columns.Add(New DataColumn("Desc1"))
            dt_clipboard.Columns.Add(New DataColumn("Desc2"))
            dt_clipboard.Columns.Add(New DataColumn("Desc3"))
            dt_clipboard.Columns.Add(New DataColumn("Desc4"))
            dt_clipboard.Columns.Add(New DataColumn("Desc5"))
            dt_clipboard.Columns.Add(New DataColumn("Desc6"))
            dt_clipboard.Columns.Add(New DataColumn("Desc7"))
            dt_clipboard.Columns.Add(New DataColumn("Desc8"))
            dt_clipboard.Columns.Add(New DataColumn("Type"))
            dt_clipboard.Columns.Add(New DataColumn("Batch", System.Type.GetType("System.Boolean")))
            dt_clipboard.Columns.Add(New DataColumn("Bin", System.Type.GetType("System.Boolean")))
            dt_clipboard.Columns.Add(New DataColumn("Serial", System.Type.GetType("System.Boolean")))
            dt_clipboard.Columns.Add(New DataColumn("MinSellPrice", System.Type.GetType("System.Double")))
            dt_clipboard.Columns.Add(New DataColumn("MaxPurPrice", System.Type.GetType("System.Double")))
            dt_clipboard.Columns.Add(New DataColumn("Tax"))
            dt_clipboard.Columns.Add(New DataColumn("TaxAmount", System.Type.GetType("System.Double")))
            dt_clipboard.Columns.Add(New DataColumn("ItemTaxDetails"))
            dt_clipboard.Columns.Add(New DataColumn("TaxPercentage", System.Type.GetType("System.Double")))
            dt_clipboard.Columns.Add(New DataColumn("NonClaimableTaxAmount"))
            dt_clipboard.Columns.Add(New DataColumn("DiscType"))
            dt_clipboard.Columns.Add(New DataColumn("DiscPercentage", System.Type.GetType("System.Double")))
            dt_clipboard.Columns.Add(New DataColumn("PriceType"))
            dt_clipboard.Columns.Add(New DataColumn("NetAmount", System.Type.GetType("System.Double")))
            'PriceType()

            dt_clipboard.Columns("BaseUnit").DefaultValue = 1
            dt_clipboard.Columns("BaseUnitPrice").DefaultValue = 0
            dt_clipboard.Columns("TCDiscountAmount").DefaultValue = 0
            dt_clipboard.Columns("Price").DefaultValue = 0
            dt_clipboard.Columns("Amount").DefaultValue = 0
            dt_clipboard.Columns("LCAmount").DefaultValue = 0
            dt_clipboard.Columns("Pieces").DefaultValue = 0
            dt_clipboard.Columns("DeliveredTotQty").DefaultValue = 0
            dt_clipboard.Columns("Package").DefaultValue = 0
            dt_clipboard.Columns("MinSellPrice").DefaultValue = 0
            dt_clipboard.Columns("MaxPurPrice").DefaultValue = 0
            dt_clipboard.Columns("DiscPercentage").DefaultValue = 0

            dt_clipboard.Columns("Tax").DefaultValue = 0
            dt_clipboard.Columns("TaxAmount").DefaultValue = 0
            dt_clipboard.Columns("TaxPercentage").DefaultValue = 0
            dt_clipboard.Columns("ItemTaxDetails").DefaultValue = 0
            dt_clipboard.Columns("NonClaimableTaxAmount").DefaultValue = 0
            dt_clipboard.Columns("NetAmount").DefaultValue = 0



            Dim i As Integer
            For i = 0 To dt.Rows.Count - 1
                dl_clipboard = dt_clipboard.NewRow
                dl_clipboard("SortNo") = i + 1
                dl_clipboard("Slno") = dt.Rows(i)("Slno")
                dl_clipboard("BarCodeNo") = dt.Rows(i)("BarCodeNo").ToString
                dl_clipboard("ItemCode") = dt.Rows(i)("ItemCode").ToString
                dl_clipboard("ItemDesc1") = dt.Rows(i)("ItemDesc1").ToString
                dl_clipboard("Alias1") = dt.Rows(i)("Alias1").ToString
                dl_clipboard("ItemDesc2") = dt.Rows(i)("ItemDesc2").ToString
                dl_clipboard("Alias2") = dt.Rows(i)("Alias2").ToString
                dl_clipboard("Unit") = dt.Rows(i)("Unit").ToString
                dl_clipboard("BaseUnit") = dt.Rows(i)("BaseUnit").ToString
                dl_clipboard("VouQty") = dt.Rows(i)("VouQty").ToString
                dl_clipboard("PrimaryQty") = dt.Rows(i)("PrimaryQty").ToString
                dl_clipboard("Package") = dt.Rows(i)("Package").ToString
                'dl_clipboard("ItemDesc1") = dt.Rows(i)("ItemDesc1").ToString
                If dt.Columns.Contains("DiscType") Then
                    dl_clipboard("DiscType") = dt.Rows(i)("DiscType").ToString
                End If
                dl_clipboard("PriceType") = dt.Rows(i)("PriceType").ToString

                If dt.Columns.Contains("Price") Then
                    dl_clipboard("Price") = dt.Rows(i)("Price").ToString
                End If

                If dt.Columns.Contains("BaseUnitPrice") Then
                    dl_clipboard("BaseUnitPrice") = dt.Rows(i)("BaseUnitPrice").ToString
                End If

                If dt.Columns.Contains("TCDiscountAmount") Then
                    dl_clipboard("TCDiscountAmount") = dt.Rows(i)("TCDiscountAmount").ToString
                End If

                If dt.Columns.Contains("Amount") Then
                    dl_clipboard("Amount") = dt.Rows(i)("Amount").ToString
                End If

                If dt.Columns.Contains("LCAmount") Then
                    dl_clipboard("LCAmount") = 0
                End If
                If dt.Columns.Contains("PartNo") Then
                    dl_clipboard("PartNo") = dt.Rows(i)("PartNo").ToString
                End If
                dl_clipboard("Comment") = dt.Rows(i)("Comment").ToString
                dl_clipboard("Desc1") = dt.Rows(i)("Desc1").ToString
                dl_clipboard("Desc2") = dt.Rows(i)("Desc2").ToString
                dl_clipboard("Desc3") = dt.Rows(i)("Desc3").ToString
                dl_clipboard("Desc4") = dt.Rows(i)("Desc4").ToString
                dl_clipboard("Desc5") = dt.Rows(i)("Desc5").ToString
                dl_clipboard("Desc6") = dt.Rows(i)("Desc6").ToString
                dl_clipboard("Desc7") = dt.Rows(i)("Desc7").ToString
                dl_clipboard("Desc8") = dt.Rows(i)("Desc8").ToString
                If dt.Columns.Contains("Type") Then
                    dl_clipboard("Type") = dt.Rows(i)("Type").ToString
                End If
                dl_clipboard("Batch") = dt.Rows(i)("Batch").ToString
                dl_clipboard("Bin") = dt.Rows(i)("Bin").ToString
                dl_clipboard("Serial") = dt.Rows(i)("Serial").ToString

                If dt.Columns.Contains("DiscPercentage") Then
                    dl_clipboard("DiscPercentage") = dt.Rows(i)("DiscPercentage").ToString
                End If

                If dt.Columns.Contains("Tax") Then
                    dl_clipboard("Tax") = dt.Rows(i)("Tax").ToString
                End If
                If dt.Columns.Contains("TaxPercentage") Then
                    dl_clipboard("TaxPercentage") = dt.Rows(i)("TaxPercentage").ToString
                End If
                If dt.Columns.Contains("TaxAmount") Then
                    dl_clipboard("TaxAmount") = dt.Rows(i)("TaxAmount").ToString
                End If
                If dt.Columns.Contains("ItemTaxDetails") Then
                    dl_clipboard("ItemTaxDetails") = dt.Rows(i)("ItemTaxDetails").ToString
                End If
                If dt.Columns.Contains("NonClaimableTaxAmount") Then
                    dl_clipboard("NonClaimableTaxAmount") = dt.Rows(i)("NonClaimableTaxAmount").ToString
                End If

                If dt.Columns.Contains("NetAmount") Then
                    dl_clipboard("NetAmount") = dt.Rows(i)("NetAmount").ToString
                End If

                'If Not strFormName.ToUpper = "INDENT" Then
                'dl_clipboard("Type") = dt.Rows(i)("Type").ToString
                'End If

                dt_clipboard.Rows.Add(dl_clipboard)
            Next
        End If
    End Sub
    Public Sub Clipboard_ItemExtraDetailsCopy(ByVal dt As DataTable)
        If dt.Rows.Count > 0 Then
            dt_clipboardItemExtraDetails = New DataTable
            Dim dl_clipboardItemExtraDetails As DataRow
            dt_clipboardItemExtraDetails.Columns.Add("SortNo")
            dt_clipboardItemExtraDetails.Columns.Add("SlNo")
            dt_clipboardItemExtraDetails.Columns.Add("ItemCode")
            dt_clipboardItemExtraDetails.Columns.Add("ImagePath")
            dt_clipboardItemExtraDetails.Columns.Add("Description")

            Dim i As Integer
            For i = 0 To dt.Rows.Count - 1
                dl_clipboardItemExtraDetails = dt_clipboardItemExtraDetails.NewRow
                dl_clipboardItemExtraDetails("SortNo") = dt.Rows(i)("SortNo")
                dl_clipboardItemExtraDetails("SlNo") = dt.Rows(i)("SortNo")
                dl_clipboardItemExtraDetails("ItemCode") = dt.Rows(i)("ItemCode").ToString
                dl_clipboardItemExtraDetails("ImagePath") = dt.Rows(i)("ImagePath").ToString
                dl_clipboardItemExtraDetails("Description") = dt.Rows(i)("Description").ToString
                dt_clipboardItemExtraDetails.Rows.Add(dl_clipboardItemExtraDetails)
            Next
        End If
    End Sub

    Public Function Clipboard_Paste() As DataTable
        Clipboard_Paste = dt_clipboard
    End Function
    Public Function ClipboardItemExtraDetails_Paste() As DataTable
        ClipboardItemExtraDetails_Paste = dt_clipboardItemExtraDetails
    End Function
End Class