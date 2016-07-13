Imports SAPbobsCOM

Public Class clsSupplierPrice
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBDataSource As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Z_OVPL Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If (pVal.ItemUID <> "1" And pVal.ItemUID <> "2") And (oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                                    Dim oDoc As SAPbouiCOM.DBDataSource
                                    oDoc = oForm.DataSources.DBDataSources.Item(0)
                                    Dim oItem As SAPbouiCOM.Item
                                    If pVal.ItemUID <> "" Then
                                        oItem = oForm.Items.Item(pVal.ItemUID)
                                        If oItem.Type <> SAPbouiCOM.BoFormItemTypes.it_BUTTON And oItem.Type <> SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON Then
                                            If (oDoc.GetValue("U_Z_DocStatus", 0).Trim <> "D") Then
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                                    If Validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                ElseIf pVal.ItemUID = "1_" Then
                                    Dim objHistory As New clsAppHistory
                                    objHistory.LoadForm(oForm, oApplication.Utilities.getEdittextvalue(oForm, "_8_"), HeaderDoctype.Spl)
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim sCHFL_ID As String
                                Dim intChoice As Integer
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OVPL")
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    Dim val, val1 As String
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        intChoice = 0
                                        oForm.Freeze(True)

                                        If oCFL.ObjectType = "2" Then
                                            val = oDataTable.GetValue("CardCode", 0)
                                            val1 = oDataTable.GetValue("CardName", 0)
                                            'oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                            'oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID & "_", val1)
                                            oDBDataSource.SetValue("U_Z_CardCode", 0, val)
                                            oDBDataSource.SetValue("U_Z_CardName", 0, val1)
                                        ElseIf oCFL.ObjectType = "4" Then
                                            val = oDataTable.GetValue("ItemCode", 0)
                                            val1 = oDataTable.GetValue("ItemName", 0)
                                            'oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                            'oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID & "_", val1)
                                            oDBDataSource.SetValue("U_Z_ItemCode", 0, val)
                                            oDBDataSource.SetValue("U_Z_ItemName", 0, val1)
                                            Dim strCCurrency As String
                                            Dim dblCPrice As Double
                                            oApplication.Utilities.GetItemPriceSupplierCatelog(val, oDBDataSource.GetValue("U_Z_CardCode", 0).ToString, strCCurrency, dblCPrice)
                                            oDBDataSource.SetValue("U_Z_CCurrency", 0, strCCurrency)
                                            oDBDataSource.SetValue("U_Z_CPrice", 0, dblCPrice)
                                        End If

                                        oForm.Freeze(False)
                                    End If
                                Catch ex As Exception
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If
                                    oForm.Freeze(False)
                                End Try

                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"

    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_Z_OVPL
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If
                Case mnu_ADD
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    initialize(oForm)
                Case mnu_FIND
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        EnableControls(oForm)
                    End If
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

#End Region

#Region "Right Click Event"

    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
        If oForm.TypeEx = frm_Z_OVPL Then
            If (eventInfo.BeforeAction = True) Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            Else
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then

                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            End If
        End If
    End Sub

#End Region

#Region "Form Data Event"

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
            If BusinessObjectInfo.FormTypeEx = frm_Z_OVPL Then
                Select Case BusinessObjectInfo.BeforeAction
                    Case True
                        Select Case BusinessObjectInfo.EventType
                            
                        End Select
                    Case False
                        Select Case BusinessObjectInfo.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                If BusinessObjectInfo.ActionSuccess = True Then
                                    oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OVPL")
                                    If oDBDataSource.GetValue("U_Z_DocStatus", 0).ToString.Trim <> "D" Then
                                        oForm.Items.Item("1").Enabled = False
                                    Else
                                        oForm.Items.Item("1").Enabled = True
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                If BusinessObjectInfo.ActionSuccess = True Then
                                    Dim s As String = oApplication.Company.GetNewObjectType()

                                    Dim stXML As String = BusinessObjectInfo.ObjectKey
                                    stXML = stXML.Replace("<?xml version=""1.0"" encoding=""UTF-16"" ?><Vendor_PriceParams><DocEntry>", "")
                                    stXML = stXML.Replace("</DocEntry></Vendor_PriceParams>", "")
                                    stXML = stXML.Replace("<?xml version=""1.0"" encoding=""UTF-16"" ?><Vendor PriceParams><DocEntry>", "")
                                    stXML = stXML.Replace("</DocEntry></Vendor PriceParams>", "")
                                    Dim otest, oItem As SAPbobsCOM.Recordset
                                    otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oItem = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
                                    If stXML <> "" Then
                                        otest.DoQuery("select * from ""@Z_OVPL""  where ""DocEntry""=" & stXML)
                                        If otest.RecordCount > 0 Then
                                            Dim strDocEntry As String = otest.Fields.Item("DocEntry").Value
                                            If 1 = 1 Then
                                                If blnIsHanaDB = True Then
                                                    oItem.DoQuery("Select ifnull(""U_Z_Identifier"",'F') from OITM where ""ItemCode""='" & otest.Fields.Item("U_Z_ItemCode").Value & "'")
                                                Else
                                                    oItem.DoQuery("Select isnull(""U_Z_Identifier"",'F') from OITM where ""ItemCode""='" & otest.Fields.Item("U_Z_ItemCode").Value & "'")
                                                End If
                                                If 1 = 1 Then ' oItem.Fields.Item(0).Value = "F" Then 'Validate the Item Identifier whether Fixed or Variable
                                                    Dim intTempID As String = oApplication.ApplProcedure.GetTemplateID(HeaderDoctype.Spl, otest.Fields.Item("U_Z_CardCode").Value, "@Z_APPT1", "U_Z_EmpId") 'oApplication.Utilities.GetTemplateID(oForm, HeaderDoctype.Train, otest.Fields.Item("U_Z_HREmpID").Value)
                                                    If intTempID <> "0" Then
                                                        oApplication.ApplProcedure.UpdateApprovalRequired("@Z_OVPL", "DocEntry", otest.Fields.Item("DocEntry").Value, "Y", intTempID, "P")
                                                        Dim strMessage As String = "Supplier Price transaction need approval for the transaction id is : " & otest.Fields.Item("DocEntry").Value
                                                        oApplication.ApplProcedure.InitialCurNextApprover("@Z_OVPL", "DocEntry", otest.Fields.Item("DocEntry").Value, intTempID, strMessage, "Supplier Price Approval Notification")
                                                        otest.DoQuery("Update ""@Z_OVPL"" set ""U_Z_DocStatus""='P' where ""DocEntry""=" & otest.Fields.Item("DocEntry").Value & "")
                                                    Else
                                                        oApplication.ApplProcedure.UpdateApprovalRequired("@Z_OVPL", "DocEntry", otest.Fields.Item("DocEntry").Value, "N", intTempID, "A")
                                                        otest.DoQuery("Update ""@Z_OVPL"" set ""U_Z_DocStatus""='A' where ""DocEntry""=" & otest.Fields.Item("DocEntry").Value & "")
                                                        oApplication.Utilities.UpdateSupplierCatelog(strDocEntry)
                                                    End If
                                                Else 'Validate the Item Identifier whether Fixed or Variable
                                                    Dim intTempID As String = oApplication.ApplProcedure.GetTemplateID(HeaderDoctype.Spl, otest.Fields.Item("U_Z_CardCode").Value, "@Z_APPT1", "U_Z_EmpId") 'oApplication.Utilities.GetTemplateID(oForm, HeaderDoctype.Train, otest.Fields.Item("U_Z_HREmpID").Value)
                                                    oApplication.ApplProcedure.UpdateApprovalRequired("@Z_OVPL", "DocEntry", otest.Fields.Item("DocEntry").Value, "N", intTempID, "A")
                                                    otest.DoQuery("Update ""@Z_OVPL"" set ""U_Z_DocStatus""='A' where ""DocEntry""=" & otest.Fields.Item("DocEntry").Value & "")
                                                    oApplication.Utilities.UpdateSupplierCatelog(strDocEntry)
                                                End If

                                            End If
                                        End If
                                    End If
                                End If
                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

#Region "Methods"

    Private Sub loadForm()
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Z_OVPL) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_Z_OVPL, frm_Z_OVPL)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.DataSources.UserDataSources.Add("series", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 100)
        initialize(oForm)
        addChooseFromList(oForm)
        fillCombo(oForm)
        oForm.DataBrowser.BrowseBy = "_8_"

        oForm.Items.Item("3").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        oForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        oForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        oForm.Items.Item("6_").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
        'oForm.Items.Item("24").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        oForm.Items.Item("1").Enabled = True
        oForm.Freeze(False)
    End Sub

    Public Sub loadFormbykey(ByVal DocEntry As String)
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Z_OVPL) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_Z_OVPL, frm_Z_OVPL)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        addChooseFromList(oForm)
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        oForm.Items.Item("_8_").Enabled = True
        oApplication.Utilities.setEdittextvalue(oForm, "_8_", DocEntry)
        oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
        End If
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
        oForm.Freeze(False)
    End Sub

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OVPL")

            oForm.Items.Item("1").Enabled = True
            oForm.Items.Item("4").Enabled = True
            oForm.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oApplication.Utilities.setEdittextvalue(oForm, "9", System.DateTime.Now.ToString("yyyyMMdd"))
            oForm.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            If oForm.DataSources.UserDataSources.Item("series").ValueEx <> "" Then
                CType(oForm.Items.Item("8").Specific, SAPbouiCOM.ComboBox).Select(oForm.DataSources.UserDataSources.Item("series").ValueEx, SAPbouiCOM.BoSearchKey.psk_ByValue)
            End If
            oForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True) 'Invoice Button
            oForm.Items.Item("3").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub addChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim strUserCode As String = oApplication.Company.UserName
            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList

            oCFL = oCFLs.Item("CFL_1")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add
            '// ((CardType = 'C') Or
            oCon.BracketOpenNum = 2
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "S"
            oCon.BracketCloseNum = 1
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

            oCon = oCons.Add
            '// (CardType = 'S'))
            oCon.BracketOpenNum = 1
            oCon.Alias = "U_Z_USERCODE"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_CONTAIN
            oCon.CondVal = strUserCode
            oCon.BracketCloseNum = 2
            oCFL.SetConditions(oCons)



            'oCFL = oCFLs.Item("CFL_1")
            'oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.Alias = "CardType"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "S"
            'oCFL.SetConditions(oCons)

            'oCon = oCons.Add()
            'oCon.Alias = "U_Z_USERCODE"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_CONTAIN
            'oCon.CondVal = strUserCode
            'oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_2")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "PrchseItem"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub enableControls(ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            Select Case aform.Mode
                Case SAPbouiCOM.BoFormMode.fm_ADD_MODE
                    aform.Items.Item("1").Enabled = True
                    aform.Items.Item("4").Enabled = False
                    aform.Items.Item("6").Enabled = False
                Case SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    aform.Items.Item("1").Enabled = True
                    aform.Items.Item("4").Enabled = True
                    aform.Items.Item("6").Enabled = True
            End Select
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub

    Private Function validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim strSupplier, StrItemCode, strCurrency, strPrice As String
            Dim dblPrice As Double = 0

            strSupplier = oApplication.Utilities.getEdittextvalue(aForm, "3")
            StrItemCode = oApplication.Utilities.getEdittextvalue(aForm, "4")
            oCombobox = aForm.Items.Item("6").Specific
            strCurrency = oCombobox.Selected.Value
            strPrice = oApplication.Utilities.getEdittextvalue(aForm, "6_")
            Double.TryParse(strPrice, dblPrice)
            If strSupplier = "" Then
                oApplication.Utilities.Message("Select Supplier....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf StrItemCode = "" Then
                oApplication.Utilities.Message("Select Item Code....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf strCurrency = "" Then
                oApplication.Utilities.Message("Select Price Currency....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf dblPrice = 0 Then
                oApplication.Utilities.Message("Enter New Price....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
        Return True
    End Function

    Private Sub fillCombo(ByVal aForm As SAPbouiCOM.Form)
        Try

            Dim oTempRec As SAPbobsCOM.Recordset
            oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oCombobox = aForm.Items.Item("8").Specific
            For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
                oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
            oCombobox.ValidValues.Add("", "")
            oTempRec.DoQuery("Select ""Series"",""SeriesName"" From NNM1 Where ""ObjectCode"" = 'Z_OVPL'")
            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                oCombobox.ValidValues.Add(oTempRec.Fields.Item("Series").Value, oTempRec.Fields.Item("SeriesName").Value)
                oTempRec.MoveNext()
            Next

            Dim strDSeries As String = getUserSeries(oForm)
            oCombobox.Select(strDSeries, SAPbouiCOM.BoSearchKey.psk_ByValue)
            oForm.DataSources.UserDataSources.Item("series").ValueEx = strDSeries


            oCombobox = aForm.Items.Item("5").Specific
            For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
                oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
            oCombobox.ValidValues.Add("", "")
            oTempRec.DoQuery("Select ""CurrCode"",""CurrName"" From OCRN")
            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                oCombobox.ValidValues.Add(oTempRec.Fields.Item("CurrCode").Value, oTempRec.Fields.Item("CurrName").Value)
                oTempRec.MoveNext()
            Next

            oCombobox = aForm.Items.Item("6").Specific
            For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
                oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
            oCombobox.ValidValues.Add("", "")
            oTempRec.DoQuery("Select ""CurrCode"",""CurrName"" From OCRN")
            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                oCombobox.ValidValues.Add(oTempRec.Fields.Item("CurrCode").Value, oTempRec.Fields.Item("CurrName").Value)
                oTempRec.MoveNext()
            Next

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function getUserSeries(ByVal oForm As SAPbouiCOM.Form) As Integer
        Dim _retVal As Integer = 0
        Try
            Dim oCmpSrv As SAPbobsCOM.CompanyService
            Dim oSeriesService As SAPbobsCOM.SeriesService
            Dim oSeries As SAPbobsCOM.Series
            Dim oDocumentTypeParams As SAPbobsCOM.DocumentTypeParams
            oCmpSrv = oApplication.Company.GetCompanyService()
            oSeriesService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.SeriesService)
            oSeries = oSeriesService.GetDataInterface(SAPbobsCOM.SeriesServiceDataInterfaces.ssdiSeries)
            oDocumentTypeParams = oSeriesService.GetDataInterface(SeriesServiceDataInterfaces.ssdiDocumentTypeParams)
            oDocumentTypeParams.Document = "Z_OVPL"
            'oDocumentTypeParams.DocumentSubType = "C"
            oSeries = oSeriesService.GetDefaultSeries(oDocumentTypeParams)
            _retVal = oSeries.Series
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function



#End Region

End Class
