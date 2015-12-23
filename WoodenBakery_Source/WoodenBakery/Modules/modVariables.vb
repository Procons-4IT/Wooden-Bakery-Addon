Public Module modVariables

    Public oApplication As clsListener
    Public strSQL As String
    Public cfl_Text As String
    Public cfl_Btn As String

    Public CompanyDecimalSeprator As String
    Public CompanyThousandSeprator As String
    Public strCardCode As String = ""
    Public blnDraft As Boolean = False
    Public blnError As Boolean = False
    Public frmSourceForm As SAPbouiCOM.Form
    Public strDocEntry As String

    Public MatrixId As String
    Public InvForConsumedItems, count As Integer
    Public RowtoDelete As Integer
    Public sPath, strSelectedFilepath, strSelectedFolderPath As String
    Public blnFlag As Boolean = False
    Dim oDataSrc_Line, oDataSrc_Line3 As SAPbouiCOM.DBDataSource
    Dim oDataSrc_Line1, oDataSrc_Line2, oDataSrc_Line4, oDataSrc_Line5 As SAPbouiCOM.DBDataSource

    Public frmSourceMatrix As SAPbouiCOM.Matrix
    Public strItemSelectionQuery As String = ""
    Public frmSourcePaymentform As SAPbouiCOM.Form
    Public frmApprovalWOrksheetForm As SAPbouiCOM.Form
    Public frm_InvoiceForm As SAPbouiCOM.Form
    Public blnInvoiceForm As Boolean = False

    Public intSelectedMatrixrow As Integer = 0
    Public strFilepath As String
    Public blnIsHanaDB As Boolean


    Public Enum ValidationResult As Integer
        CANCEL = 0
        OK = 1
    End Enum

    Public Enum DocumentType As Integer
        RENTAL_QUOTATION = 1
        RENTAL_ORDER
        RENTAL_RETURN
    End Enum
    Public Enum HeaderDoctype
        Fix
        Spl
    End Enum

    Public Const frm_WAREHOUSES As Integer = 62
    Public Const frm_FixedAsset As String = "1473000075"
    Public Const frm_ARInvoicePayment As String = "60090"
    Public Const mnu_PaymentMeans As String = "5892"
    Public Const frm_PaymentMeans As String = "146"
    Public Const frm_BatchSelect As String = "42"
    Public Const frm_BatchSetup As String = "41"

    Public Const frm_StockRequest As String = "frm_StRequest"
    Public Const frm_InvSO As String = "frm_InvSO"
    Public Const frm_Warehouse As String = "62"
    Public Const frm_PurchaseOrder As String = "142"
    Public Const frm_SalesOrder As String = "139"
    Public Const frm_Invoice As String = "133"
    Public Const frm_BPMaster As String = "134"
    Public Const frm_ItemMaster As String = "150"
    Public Const frm_Customer As String = "134"

    Public Const frm_APInvoice As String = "141"
    Public Const frm_ARCreditNote As String = "179"
  
    Public Const mnu_FIND As String = "1281"
    Public Const mnu_ADD As String = "1282"
    Public Const mnu_CLOSE As String = "1286"
    Public Const mnu_NEXT As String = "1288"
    Public Const mnu_PREVIOUS As String = "1289"
    Public Const mnu_Cancel As String = "1284"
    Public Const mnu_FIRST As String = "1290"
    Public Const mnu_LAST As String = "1291"
    Public Const mnu_ADD_ROW As String = "1292"
    Public Const mnu_DELETE_ROW As String = "1293"
    Public Const mnu_TAX_GROUP_SETUP As String = "8458"
    Public Const mnu_DEFINE_ALTERNATIVE_ITEMS As String = "11531"
    Public Const mnu_CloseOrderLines As String = "DABT_910"
    Public Const mnu_InvSO As String = "DABT_911"

    Public Const mnu_CPRL_O As String = "CPRL_O"
    Public Const mnu_CPRL_C As String = "CPRL_C"
    Public Const mnu_ICCancel As String = "mnu_ICCancel"

    
    Public Const xml_MENU As String = "Menu.xml"
    Public Const xml_MENU_REMOVE As String = "RemoveMenus.xml"
    Public Const xml_StRequest As String = "StRequest.xml"
    Public Const xml_InvSO As String = "frm_InvSO.xml"

    Public Const frm_SalePriceBP As String = "333"


    Public Const frm_CustComDef As String = "frm_CustComDef"
    Public Const xml_CustRebate As String = "frm_CustComDef.xml"
    Public Const mnu_CustRebate As String = "Mnu_CustRebate"

    Public Const frm_SubComDef As String = "frm_SubComDef"
    Public Const xml_SubRebate As String = "frm_SubComDef.xml"
    Public Const mnu_SubRebate As String = "Mnu_SubRebate"

    Public Const frm_DisRule As String = "frm_DisRule"
    Public Const xml_DisRule As String = "frm_DisRule.xml"


    Public Const frm_ReasonCode As String = "frm_ReasonCode"
    Public Const xml_ReasonCode As String = "xml_ReasonCode.xml"
    Public Const mnu_ReasonCode As String = "W_1001"

  
    Public Const frm_FATransaction As String = "frm_FATrans"
    Public Const xml_FATransaction As String = "xml_FATrans.xml"
    Public Const mnu_FATransaction As String = "W_1002"

    Public Const frm_SOClosing As String = "frm_CLS"
    Public Const xml_SOClosing As String = "xml_SOClosing.xml"
    Public Const mnu_SOClosing As String = "W_1003"

    Public Const frm_SOClosingRepot As String = "frm_CLSR"
    Public Const xmm_SOClosingReport As String = "xml_SOCReport.xml"
    Public Const mnu_SOClosingReprot As String = "W_1004"

    Public Const frm_FATransactionApp As String = "frm_FATransApp"
    Public Const xml_FATransactionApp As String = "xml_FATransApp.xml"
    Public Const mnu_FATransactionApp As String = "W_1005"

    Public Const frm_WeekEndMaster As String = "frm_WeekEnd"
    Public Const xml_WeekEndMaster As String = "xml_WeekEnd.xml"
    Public Const mnu_WeekEndMaster As String = "W_1006"

    Public Const frm_UnLock As String = "frm_UnLock"
    Public Const xml_UnLock As String = "xml_UnLock.xml"
    Public Const mnu_UnLock As String = "W_1007"

    Public Const frm_ApprovalTemplate As String = "frm_AppTemplate"
    Public Const xml_ApprovalTemplate As String = "xml_ApprovalTemplate.xml"
    Public Const mnu_ApprovalTemplate As String = "W_1008"

    Public Const frm_ChoosefromList_Leave As String = "frm_CFLLeave"
    Public Const frm_ChoosefromList_UOM As String = "frm_CFLUOM"

    Public Const frm_AppHistory As String = "frm_AppHistory"
    Public Const xm_AppHistory As String = "xm_AppHistory.xml"

    Public Const frm_DocumentType As String = "frm_DocType"
    Public Const mnu_DocumentType As String = "W_1009"
    Public Const xml_DocumentType As String = "xml_DocType.xml"


    Public Const frm_ItemCagetory As String = "frm_ItmCr"
    Public Const mnu_ItemCagetory As String = "W_1010"
    Public Const xml_ItemCagetory As String = "xml_ItmCr.xml"

    Public Const frm_BPmCagetory As String = "frm_BPCr"
    Public Const mnu_BPCagetory As String = "W_1011"
    Public Const xml_BPCagetory As String = "xml_BPCr.xml"

    Public Const frm_WhsCagetory As String = "frm_WhsCr"
    Public Const mnu_WhsCagetory As String = "W_1012"
    Public Const xml_WhsCagetory As String = "xml_WhsCr.xml"

    Public Const frm_Rebate As String = "frm_Rebate"
    Public Const mnu_Rebate As String = "W_1013"
    Public Const xml_Rebate As String = "xml_Rebate.xml"

    Public Const frm_Z_OVPL As String = "frm_Z_OVPL"
    Public Const xml_Z_OVPL As String = "xml_Z_OVPL.xml"
    Public Const mnu_Z_OVPL As String = "W_1014"

    Public Const frm_Z_OVPL_A As String = "frm_Z_OVPL_A"
    Public Const xml_Z_OVPL_A As String = "xml_Z_OVPL_A.xml"
    Public Const mnu_Z_OVPL_A As String = "W_1015"

    Public Const frm_Z_ODEL As String = "frm_Z_ODEL"
    Public Const xml_Z_ODEL As String = "xml_Z_ODEL.xml"
    Public Const mnu_Z_ODEL As String = "W_1016"

    Public Const frm_Z_ODEL_R As String = "frm_Z_ODEL_R"
    Public Const xml_Z_ODEL_R As String = "xml_Z_ODEL_R.xml"
    Public Const mnu_Z_ODEL_R As String = "W_1017"

    Public Const frm_Z_OPRM As String = "frm_Z_OPRM"
    Public Const xml_Z_OPRM As String = "xml_Z_OPRM.xml"
    Public Const mnu_Z_OPRM As String = "W_1018"

    Public Const frm_Z_OCPR As String = "frm_Z_OCPR"
    Public Const xml_Z_OCPR As String = "xml_Z_OCPR.xml"
    Public Const mnu_Z_OCPR As String = "W_1019"

    Public Const frm_Z_CPRL As String = "frm_Z_CPRL"
    Public Const xml_Z_CPRL As String = "xml_Z_CPRL.xml"

    Public Const frm_Z_OICT As String = "frm_Z_OICT"
    Public Const xml_Z_OICT As String = "xml_Z_OICT.xml"
    Public Const mnu_Z_OICT As String = "W_1020"

End Module
