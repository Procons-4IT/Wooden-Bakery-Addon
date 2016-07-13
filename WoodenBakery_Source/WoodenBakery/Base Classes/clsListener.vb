Public Class clsListener
    Inherits Object
    Private ThreadClose As New Threading.Thread(AddressOf CloseApp)
    Private WithEvents _SBO_Application As SAPbouiCOM.Application
    Private _Company As SAPbobsCOM.Company
    Private _Utilities As clsUtilities
    Private _Collection As Hashtable
    Private _LookUpCollection As Hashtable
    Private _FormUID As String
    Private _AppProcedure As clsApprovalProcedure
    Private _Log As clsLog_Error
    Private oMenuObject As Object
    Private oItemObject As Object
    Private oSystemForms As Object
    Dim objFilters As SAPbouiCOM.EventFilters
    Dim objFilter As SAPbouiCOM.EventFilter
    Private _blnShowBatchSelection As Boolean = False

#Region "New"
    Public Sub New()
        MyBase.New()
        Try
            _Company = New SAPbobsCOM.Company
            _Utilities = New clsUtilities
            _Collection = New Hashtable(10, 0.5)
            _LookUpCollection = New Hashtable(10, 0.5)
            oSystemForms = New clsSystemForms
            _Log = New clsLog_Error
            _AppProcedure = New clsApprovalProcedure
            SetApplication()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Public Properties"

    Public ReadOnly Property SBO_Application() As SAPbouiCOM.Application
        Get
            Return _SBO_Application
        End Get
    End Property

    Public ReadOnly Property Company() As SAPbobsCOM.Company
        Get
            Return _Company
        End Get
    End Property

    Public ReadOnly Property Utilities() As clsUtilities
        Get
            Return _Utilities
        End Get
    End Property

    Public ReadOnly Property Collection() As Hashtable
        Get
            Return _Collection
        End Get
    End Property

    Public ReadOnly Property LookUpCollection() As Hashtable
        Get
            Return _LookUpCollection
        End Get
    End Property

    Public ReadOnly Property Log() As clsLog_Error
        Get
            Return _Log
        End Get
    End Property
    Public ReadOnly Property ApplProcedure() As clsApprovalProcedure
        Get
            Return _AppProcedure
        End Get
    End Property

    Public Property ShowBatchSelection() As Boolean
        Get
            Return _blnShowBatchSelection
        End Get
        Set(value As Boolean)
            _blnShowBatchSelection = value
        End Set
    End Property

#Region "Filter"

    Public Sub SetFilter(ByVal Filters As SAPbouiCOM.EventFilters)
        oApplication.SetFilter(Filters)
    End Sub
    Public Sub SetFilter()
        Try
            ''Form Load
            objFilters = New SAPbouiCOM.EventFilters

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
            objFilter.AddEx(frm_SalesOrder)
            objFilter.AddEx(frm_Z_ODEL)

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN)

            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
            objFilter.AddEx(frm_Z_ODEL)
            objFilter.AddEx(frm_SalesOrder)
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
#End Region

#End Region

#Region "Data/Menu Event"

    Private Sub _SBO_Application_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles _SBO_Application.FormDataEvent
        Select Case BusinessObjectInfo.FormTypeEx
            Case frm_Invoice, frm_ARCreditNote
                Dim objInvoice As clsRebatePosting
                objInvoice = New clsRebatePosting
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case frm_ItemCagetory
                Dim objInvoice As clsItemCategory
                objInvoice = New clsItemCategory
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)

            Case frm_BPmCagetory
                Dim objInvoice As clsBPCategory
                objInvoice = New clsBPCategory
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)

            Case frm_WhsCagetory
                Dim objInvoice As clsWarehouseCategory
                objInvoice = New clsWarehouseCategory
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)



            Case frm_DocumentType
                Dim objInvoice As clsDocType
                objInvoice = New clsDocType
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case frm_ReasonCode
                Dim objInvoice As clsReasonCode
                objInvoice = New clsReasonCode
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)

            Case frm_FATransaction
                Dim objInvoice As clsFixedAssetTransaction
                objInvoice = New clsFixedAssetTransaction
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)

            Case frm_WeekEndMaster
                Dim objInvoice As clsWeekEndMaster
                objInvoice = New clsWeekEndMaster
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)

            Case frm_UnLock
                Dim objInvoice As clsUnlockPostingDate
                objInvoice = New clsUnlockPostingDate
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case frm_ApprovalTemplate
                Dim objInvoice As clsAppTemplate
                objInvoice = New clsAppTemplate
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case frm_Z_OVPL
                Dim objSupplierPrice As clsSupplierPrice
                objSupplierPrice = New clsSupplierPrice
                objSupplierPrice.FormDataEvent(BusinessObjectInfo, BubbleEvent)

            Case frm_Z_ODEL
                Dim objDeliveryDoc As clsDeliveryDoc
                objDeliveryDoc = New clsDeliveryDoc
                objDeliveryDoc.FormDataEvent(BusinessObjectInfo, BubbleEvent)
            Case frm_Z_OPRM
                Dim objProDoc As clsPromotion
                objProDoc = New clsPromotion
                objProDoc.FormDataEvent(BusinessObjectInfo, BubbleEvent)

            Case frm_Z_OICT
                Dim objIncDoc As clsInventoryCount
                objIncDoc = New clsInventoryCount
                objIncDoc.FormDataEvent(BusinessObjectInfo, BubbleEvent)

            Case frm_BPMaster, frm_ItemMaster, frm_Warehouse, frm_FixedAsset
                Dim objInvoice As clsMasters
                objInvoice = New clsMasters
                objInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
        End Select
        '  End If
    End Sub

    Private Sub SBO_Application_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles _SBO_Application.MenuEvent
        Try
            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID
                    Case mnu_Rebate
                        oMenuObject = New clsRebate
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_ItemCagetory
                        oMenuObject = New clsItemCategory
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_BPCagetory
                        oMenuObject = New clsBPCategory
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_WhsCagetory
                        oMenuObject = New clsWarehouseCategory
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_DocumentType
                        oMenuObject = New clsDocType
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_FATransactionApp
                        oMenuObject = New clsFixedAssetApproval
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_ApprovalTemplate
                        oMenuObject = New clsAppTemplate
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_ReasonCode
                        oMenuObject = New clsReasonCode
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_SOClosing
                        oMenuObject = New clsSOPOClosing
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_SOClosingReprot
                        oMenuObject = New clsSOPOClosingReport
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_FATransaction
                        oMenuObject = New clsFixedAssetTransaction
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_SOClosing
                        oMenuObject = New clsSOPOClosing
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_SOClosingReprot
                        oMenuObject = New clsSOPOClosingReport
                        oMenuObject.MenuEvent(pVal, BubbleEvent)

                    Case mnu_WeekEndMaster
                        oMenuObject = New clsWeekEndMaster
                        oMenuObject.MenuEvent(pVal, BubbleEvent)

                    Case mnu_UnLock
                        oMenuObject = New clsUnlockPostingDate
                        oMenuObject.MenuEvent(pVal, BubbleEvent)

                    Case mnu_Z_OVPL
                        oMenuObject = New clsSupplierPrice
                        oMenuObject.MenuEvent(pVal, BubbleEvent)

                    Case mnu_Z_OVPL_A
                        oMenuObject = New clsSupplierPrice_Approval
                        oMenuObject.MenuEvent(pVal, BubbleEvent)

                    Case mnu_Z_ODEL
                        oMenuObject = New clsDeliveryDoc
                        oMenuObject.MenuEvent(pVal, BubbleEvent)

                    Case mnu_Z_ODEL_R
                        oMenuObject = New clsDeliveryDocReport
                        oMenuObject.MenuEvent(pVal, BubbleEvent)

                    Case mnu_Z_OPRM, mnu_CPRL_IP
                        oMenuObject = New clsPromotion
                        oMenuObject.MenuEvent(pVal, BubbleEvent)

                    Case mnu_Z_OCPR
                        oMenuObject = New clsPromotionMapping
                        oMenuObject.MenuEvent(pVal, BubbleEvent)

                    Case mnu_Z_OCPRS
                        oMenuObject = New clsPromotionMappingSupplier
                        oMenuObject.MenuEvent(pVal, BubbleEvent)

                    Case mnu_Z_OICT
                        oMenuObject = New clsInventoryCount
                        oMenuObject.MenuEvent(pVal, BubbleEvent)


                    Case mnu_ADD, mnu_FIND, mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS, mnu_ADD_ROW, mnu_DELETE_ROW, mnu_Cancel
                        If _Collection.ContainsKey(_FormUID) Then
                            oMenuObject = _Collection.Item(_FormUID)
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        End If
                    Case mnu_PaymentMeans
                        Dim oform As SAPbouiCOM.Form
                        oform = oApplication.SBO_Application.Forms.ActiveForm()
                        ' oApplication.Utilities.PopulateDocTotaltoPaymentMeans(frm_InvoiceForm, oform)

                    Case mnu_CPRL_C
                        oMenuObject = New clsBusinessPartner
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                        'Case mnu_BatchSelection, mnu_BatchSelection1
                        '    Try
                        '        oApplication.ShowBatchSelection = True
                        '        'oApplication._SBO_Application.ActivateMenuItem(mnu_BatchSelection)
                        '    Catch ex As Exception

                        '    End Try
                    Case mnu_CPRL_O
                        oMenuObject = New clsDocuments
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                    Case mnu_CPRL_I
                        oMenuObject = New clsMasters
                        oMenuObject.MenuEvent(pVal, BubbleEvent)
                End Select

            Else
                Select Case pVal.MenuUID
                    Case mnu_BatchSelection, mnu_BatchSelection1
                        Try
                            oApplication.ShowBatchSelection = True
                            'oApplication._SBO_Application.ActivateMenuItem(mnu_BatchSelection)
                        Catch ex As Exception

                        End Try
                    Case mnu_PaymentMeans
                        frm_InvoiceForm = oApplication.SBO_Application.Forms.ActiveForm()

                    Case mnu_CLOSE
                        If _Collection.ContainsKey(_FormUID) Then
                            oMenuObject = _Collection.Item(_FormUID)
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        End If
                    Case mnu_ADD, mnu_FIND, mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS, mnu_ADD_ROW, mnu_DELETE_ROW, mnu_Cancel, mnu_CPRL_O, mnu_CPRL_I
                        If _Collection.ContainsKey(_FormUID) Then
                            oMenuObject = _Collection.Item(_FormUID)
                            oMenuObject.MenuEvent(pVal, BubbleEvent)
                        End If

                        If pVal.MenuUID = mnu_DELETE_ROW Then
                            Dim oForm As SAPbouiCOM.Form
                            If IsNothing(_FormUID) Then
                            Else
                                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                                If (oForm.TypeEx = frm_SalesOrder.ToString) Then
                                    'Dim oMatrix As SAPbouiCOM.Matrix
                                    'oMatrix = oForm.Items.Item("38").Specific
                                    'Dim intRow = oMatrix.GetCellFocus().rowIndex
                                    'If oMatrix.Columns.Item("U_PrRef").Cells().Item(intRow).Specific.value <> "" Then
                                    '    If oApplication.SBO_Application.MessageBox("Promotion Items Link for row about to delete?", , "Yes", "No") = 1 Then
                                    '        Dim strRef As String = oMatrix.Columns.Item("U_PrRef").Cells().Item(intRow).Specific.value

                                    '        'Delete Row
                                    '        Dim intRowCount As Integer = oMatrix.RowCount
                                    '        While intRowCount > 0
                                    '            If strRef = oMatrix.Columns.Item("U_PrRef").Cells().Item(intRowCount).Specific.value And CType(oMatrix.Columns.Item("U_IType").Cells().Item(intRowCount).Specific, SAPbouiCOM.ComboBox).Selected.Value = "F" Then
                                    '                oMatrix.DeleteRow(intRowCount)
                                    '            End If
                                    '            intRowCount -= 1
                                    '        End While

                                    '    Else
                                    '        BubbleEvent = False
                                    '    End If
                                    'End If
                                End If
                            End If
                        End If

                End Select

            End If

        Catch ex As Exception
            Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            oMenuObject = Nothing
        End Try
    End Sub

#End Region

#Region "Item Event"
    Private Sub SBO_Application_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles _SBO_Application.ItemEvent
        Try
            _FormUID = FormUID

            If pVal.Before_Action = True Then
                Dim oform As SAPbouiCOM.Form
                oform = oApplication.SBO_Application.Forms.Item(FormUID)
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        oform = oApplication.SBO_Application.Forms.Item(FormUID)
                        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
                        Dim oCons As SAPbouiCOM.Conditions
                        Dim oCon As SAPbouiCOM.Condition
                        Dim oCFL As SAPbouiCOM.ChooseFromList
                        Dim OItem As SAPbouiCOM.Item
                        Dim oEdittext As SAPbouiCOM.EditText
                        Dim oMatrix As SAPbouiCOM.Matrix
                        Dim strCFLID As String = ""
                        Try
                            If oform.TypeEx <> "0" And pVal.ItemUID <> "" Then
                                OItem = oform.Items.Item(pVal.ItemUID)
                                If OItem.Type = SAPbouiCOM.BoFormItemTypes.it_MATRIX Then
                                    oMatrix = OItem.Specific
                                    strCFLID = oMatrix.Columns.Item(pVal.ColUID).ChooseFromListUID
                                ElseIf OItem.Type = SAPbouiCOM.BoFormItemTypes.it_EDIT Then
                                    oEdittext = OItem.Specific
                                    strCFLID = oEdittext.ChooseFromListUID
                                End If
                                oApplication.Utilities.filterProjectChooseFromList(oform, strCFLID)
                            End If
                        Catch ex As Exception
                            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        End Try
                End Select
                If oform.TypeEx = frm_FixedAsset Then
                    If pVal.ItemUID = "1470002156" Or pVal.ItemUID = "1470002158" Then
                        If oform.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.ItemUID = "1" And (oform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                    If oform.TypeEx <> "0" Then
                        If oApplication.Utilities.UnlockSpecificDate(oform) = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If
                End If
                If pVal.FormTypeEx = "60092" Or pVal.FormTypeEx = "141" Or pVal.FormTypeEx = "143" Then
                    Dim oMatrix As SAPbouiCOM.Matrix
                    Select Case pVal.EventType
                        Case SAPbouiCOM.BoEventTypes.et_CLICK
                            oform = oApplication.SBO_Application.Forms.Item(FormUID)
                            If pVal.ItemUID = "38" And pVal.ColUID = "14" Then
                                oMatrix = oform.Items.Item(pVal.ItemUID).Specific
                                If oApplication.Utilities.ValidateItemIdentifier(oApplication.Utilities.getMatrixValues(oMatrix, "1", pVal.Row)) = False Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If
                        Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                            oform = oApplication.SBO_Application.Forms.Item(FormUID)
                            If pVal.ItemUID = "38" And pVal.ColUID = "14" And pVal.CharPressed <> 9 Then
                                oMatrix = oform.Items.Item(pVal.ItemUID).Specific
                                If oApplication.Utilities.ValidateItemIdentifier(oApplication.Utilities.getMatrixValues(oMatrix, "1", pVal.Row)) = False Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If
                    End Select
                End If
            End If
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                Select Case pVal.FormType
                End Select
            End If
            If pVal.EventType <> SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                Select Case pVal.FormTypeEx
                    Case frm_Rebate
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsRebate
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_ItemCagetory
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsItemCategory
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_BPmCagetory
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsBPCategory
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_WhsCagetory
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsWarehouseCategory
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_DocumentType
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsDocType
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_BPMaster, frm_ItemMaster, frm_Warehouse, frm_FixedAsset
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsMasters
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_BatchSelect
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsBatchSelection
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_ARInvoicePayment
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsDocuments
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_PaymentMeans
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsDocuments
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_SalesOrder
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsDocuments
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_FATransactionApp
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsFixedAssetApproval
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_ApprovalTemplate
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsAppTemplate
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_ReasonCode
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsReasonCode
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_SOClosing
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsSOPOClosing
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_SOClosingRepot
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsSOPOClosingReport
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_FATransaction
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsFixedAssetTransaction
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If
                    Case frm_WeekEndMaster
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsWeekEndMaster
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_BPMaster
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsBusinessPartner
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_UnLock
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsUnlockPostingDate
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_ChoosefromList_Leave
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsChooseFromList_Leave
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_ChoosefromList_UOM
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsChooseFromList_UOM
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_DisRule
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsDisRule
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_Z_OVPL
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsSupplierPrice
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_Z_OVPL_A
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsSupplierPrice_Approval
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_Z_ODEL
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsDeliveryDoc
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_Z_ODEL_R
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsDeliveryDocReport
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If


                    Case frm_PurchaseOrder
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsPurchaseOrder
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_Z_OPRM
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsPromotion
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_Z_OCPR
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsPromotionMapping
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_Z_OCPRS
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsPromotionMappingSupplier
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                    Case frm_Z_OICT
                        If Not _Collection.ContainsKey(FormUID) Then
                            oItemObject = New clsInventoryCount
                            oItemObject.FrmUID = FormUID
                            _Collection.Add(FormUID, oItemObject)
                        End If

                End Select
            End If

            If _Collection.ContainsKey(FormUID) Then
                oItemObject = _Collection.Item(FormUID)
                If oItemObject.IsLookUpOpen And pVal.BeforeAction = True Then
                    _SBO_Application.Forms.Item(oItemObject.LookUpFormUID).Select()
                    BubbleEvent = False
                    Exit Sub
                End If
                _Collection.Item(FormUID).ItemEvent(FormUID, pVal, BubbleEvent)
            End If
            If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD And pVal.BeforeAction = False Then
                If _LookUpCollection.ContainsKey(FormUID) Then
                    oItemObject = _Collection.Item(_LookUpCollection.Item(FormUID))
                    If Not oItemObject Is Nothing Then
                        oItemObject.IsLookUpOpen = False
                    End If
                    _LookUpCollection.Remove(FormUID)
                End If

                If _Collection.ContainsKey(FormUID) Then
                    _Collection.Item(FormUID) = Nothing
                    _Collection.Remove(FormUID)
                End If

            End If

        Catch ex As Exception
            Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub
#End Region

#Region "Right Click Event"

    Private Sub _SBO_Application_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles _SBO_Application.RightClickEvent
        Try
            Dim oForm As SAPbouiCOM.Form
            oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
            If oForm.TypeEx = frm_SalesOrder Then
                oMenuObject = New clsDocuments
                oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            ElseIf (oForm.TypeEx = frm_Customer) Then
                oMenuObject = New clsBusinessPartner
                oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            ElseIf (oForm.TypeEx = frm_Z_OPRM) Then
                oMenuObject = New clsPromotion
                oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            ElseIf (oForm.TypeEx = frm_ItemMaster) Then
                oMenuObject = New clsMasters
                oMenuObject.RightClickEvent(eventInfo, BubbleEvent)
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

#Region "Application Event"
    Private Sub SBO_Application_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles _SBO_Application.AppEvent
        Try
            Select Case EventType
                Case SAPbouiCOM.BoAppEventTypes.aet_ShutDown, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged
                    _Utilities.AddRemoveMenus("RemoveMenus.xml")
                    CloseApp()
            End Select
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Termination Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
        End Try
    End Sub
#End Region

#Region "Close Application"
    Private Sub CloseApp()
        Try
            If Not _SBO_Application Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_SBO_Application)
            End If

            If Not _Company Is Nothing Then
                If _Company.Connected Then
                    _Company.Disconnect()
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_Company)
            End If

            _Utilities = Nothing
            _Collection = Nothing
            _LookUpCollection = Nothing

            ThreadClose.Sleep(10)
            System.Windows.Forms.Application.Exit()
        Catch ex As Exception
            Throw ex
        Finally
            oApplication = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub
#End Region

#Region "Set Application"
    Private Sub SetApplication()
        Dim SboGuiApi As SAPbouiCOM.SboGuiApi
        Dim sConnectionString As String

        Try
            If Environment.GetCommandLineArgs.Length > 1 Then
                sConnectionString = Environment.GetCommandLineArgs.GetValue(1)
                SboGuiApi = New SAPbouiCOM.SboGuiApi
                SboGuiApi.Connect(sConnectionString)
                _SBO_Application = SboGuiApi.GetApplication()
            Else
                Throw New Exception("Connection string missing.")
            End If

        Catch ex As Exception
            Throw ex
        Finally
            SboGuiApi = Nothing
        End Try
    End Sub
#End Region

#Region "Finalize"
    Protected Overrides Sub Finalize()
        Try
            MyBase.Finalize()
            '            CloseApp()

            oMenuObject = Nothing
            oItemObject = Nothing
            oSystemForms = Nothing

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Addon Termination Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
        Finally
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub
#End Region

    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window

        Private _hwnd As IntPtr

        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle() As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property

    End Class
    
End Class
