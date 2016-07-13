Public Class clsPromotionMapping
    Inherits clsBase

    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Dim oStatic As SAPbouiCOM.StaticText
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oComboColumn As SAPbouiCOM.ComboBoxColumn
    Private oCustomerGrid As SAPbouiCOM.Grid
    Private oItemGrid As SAPbouiCOM.Grid
    Private dtFirst As SAPbouiCOM.DataTable
    Private dtSecond As SAPbouiCOM.DataTable
    Private dtThird As SAPbouiCOM.DataTable
    Private ocombo As SAPbouiCOM.ComboBoxColumn
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Dim oRecordSet As SAPbobsCOM.Recordset

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Public Sub LoadForm()
        Try
            oForm = oApplication.Utilities.LoadForm(xml_Z_OCPR, frm_Z_OCPR)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            oForm.PaneLevel = 1
            initialize(oForm)
            addChooseFromListConditions(oForm)
            FillCombo(oForm)
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Z_OCPR Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" And (oForm.PaneLevel = 2) Then
                                    If Validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" And (oForm.PaneLevel = 1) Then
                                    oForm.PaneLevel = oForm.PaneLevel + 1
                                    changeLabel(oForm)
                                ElseIf pVal.ItemUID = "3" And (oForm.PaneLevel = 2) Then
                                    LoadPromotion(oForm)
                                    oCustomerGrid = oForm.Items.Item("11").Specific
                                    oItemGrid = oForm.Items.Item("14").Specific
                                    If oCustomerGrid.DataTable.Rows.Count >= 1 And oItemGrid.DataTable.Rows.Count >= 1 Then
                                        oForm.PaneLevel = oForm.PaneLevel + 1
                                        changeLabel(oForm)
                                    Else
                                        If oCustomerGrid.DataTable.Rows.Count = 0 Then
                                            oApplication.Utilities.Message("No Customer Found for the Selection...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        ElseIf oItemGrid.DataTable.Rows.Count = 0 Then
                                            oApplication.Utilities.Message("No Promotion Items Found for the Selection...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        End If
                                    End If
                                ElseIf pVal.ItemUID = "6" And (oForm.PaneLevel = 3 Or oForm.PaneLevel = 4 Or oForm.PaneLevel = 5) Then
                                    If oApplication.SBO_Application.MessageBox("Do you want to Map Promotion Documents...?", , "Continue", "Cancel") = 2 Then
                                    Else
                                        If InsertPromotionMapping(oForm) = True Then
                                            oApplication.Utilities.Message("Promotion Mapped Successfully...", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            oForm.Close()
                                        End If
                                    End If
                                ElseIf pVal.ItemUID = "4" Then
                                    oForm.Freeze(True)
                                    If oForm.PaneLevel <> 2 Then
                                        oForm.PaneLevel = 2
                                    Else
                                        oForm.PaneLevel = oForm.PaneLevel - 1
                                        changeLabel(oForm)
                                    End If
                                    oForm.Freeze(False)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                If oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized Or oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Minimized Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    reDrawForm(oForm)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Dim strValue, strName As String
                                Dim dtEffDate, dtEffTo As Date
                                Try
                                    oCFLEvento = pVal
                                    oDataTable = oCFLEvento.SelectedObjects
                                    If pVal.ItemUID = "8" Or pVal.ItemUID = "19" Then
                                        strValue = oDataTable.GetValue(CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).ChooseFromListAlias, 0)
                                        Try
                                            oForm.Items.Item(pVal.ItemUID).Specific.value = strValue
                                        Catch ex As Exception
                                            oForm.Items.Item(pVal.ItemUID).Specific.value = strValue
                                        End Try
                                    ElseIf pVal.ItemUID = "12" Then
                                        strValue = oDataTable.GetValue(CType(oForm.Items.Item(pVal.ItemUID).Specific, SAPbouiCOM.EditText).ChooseFromListAlias, 0)
                                        strName = oDataTable.GetValue("U_Z_PrName", 0)
                                        Try
                                            oForm.Items.Item(pVal.ItemUID).Specific.value = strValue
                                            oForm.Items.Item("18").Specific.value = strName
                                            dtEffDate = oDataTable.GetValue("U_Z_EffFrom", 0)
                                            dtEffTo = oDataTable.GetValue("U_Z_EffTo", 0)
                                            oForm.Items.Item("29").Specific.value = dtEffDate.Date.ToString("yyyyMMdd")
                                            oForm.Items.Item("30").Specific.value = dtEffTo.Date.ToString("yyyyMMdd")
                                        Catch ex As Exception
                                            oForm.Items.Item(pVal.ItemUID).Specific.value = strValue
                                            oForm.Items.Item("18").Specific.value = strName
                                            dtEffDate = oDataTable.GetValue("U_Z_EffFrom", 0)
                                            dtEffTo = oDataTable.GetValue("U_Z_EffTo", 0)
                                            oForm.Items.Item("29").Specific.value = dtEffDate.Date.ToString("yyyyMMdd")
                                            oForm.Items.Item("30").Specific.value = dtEffTo.Date.ToString("yyyyMMdd")
                                        End Try
                                    End If
                                Catch ex As Exception

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
                Case mnu_Z_OCPR
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Data Event"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Validations"
    Private Function Validation(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Try
            Dim strFromCustomer, strToCustomer, strPromotion, strCustGroup, strCustProp As String
            strFromCustomer = oApplication.Utilities.getEditTextvalue(oForm, "8")
            strToCustomer = oApplication.Utilities.getEditTextvalue(oForm, "19")
            strPromotion = oApplication.Utilities.getEditTextvalue(oForm, "12")
            strCustGroup = CType(oForm.Items.Item("10").Specific, SAPbouiCOM.ComboBox).Value
            strCustProp = CType(oForm.Items.Item("26").Specific, SAPbouiCOM.ComboBox).Value

            If strFromCustomer = "" And strCustGroup.Length = 0 And strCustProp.Length = 0 Then
                oApplication.Utilities.Message("Enter From Customer ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf strToCustomer = "" And strCustGroup.Length = 0 And strCustProp.Length = 0 Then
                oApplication.Utilities.Message("Enter To Customer ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf strPromotion = "" Then
                oApplication.Utilities.Message("Enter To Promotion ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

#End Region

#Region "Function"

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Items.Item("1").TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD
            oForm.Items.Item("17").TextStyle = SAPbouiCOM.BoTextStyle.ts_BOLD
            oForm.DataSources.DataTables.Add("dtCustomers")
            oForm.DataSources.DataTables.Add("dtPromotion")
            oForm.Items.Item("13").TextStyle = 5
            oForm.Items.Item("24").TextStyle = 5
            changeLabel(oForm)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub addChooseFromListConditions(ByVal oForm As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList

            oCFLs = oForm.ChooseFromLists

            oCFL = oCFLs.Item("CFL_2")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_3")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_5")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_Active"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub LoadPromotion(ByVal aform As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            Dim strqry As String
            Dim strFromCust, strToCust, strPromotion, strPromotionName, strCustGroup, strProperty As String

            strFromCust = oForm.Items.Item("8").Specific.value
            strToCust = oForm.Items.Item("19").Specific.value
            strCustGroup = CType(oForm.Items.Item("10").Specific, SAPbouiCOM.ComboBox).Value.Trim()
            strProperty = CType(oForm.Items.Item("26").Specific, SAPbouiCOM.ComboBox).Value.Trim()
            strPromotion = oForm.Items.Item("12").Specific.value
            strPromotionName = oForm.Items.Item("18").Specific.value

            oForm.Items.Item("23").Specific.value = strPromotion
            oForm.Items.Item("25").Specific.value = strPromotionName

            oCustomerGrid = oForm.Items.Item("11").Specific
            oCustomerGrid.DataTable = oForm.DataSources.DataTables.Item("dtCustomers")

            strqry = " Select 'Y' As ""Select"",""CardCode"",""CardName"" From OCRD Where ""CardType"" = 'C' "

            If strFromCust.Length > 0 And strToCust.Length > 0 Then
                strqry += "And ""CardCode"" Between '" + strFromCust + "' AND '" + strToCust + "'"
            End If

            If strCustGroup.Length > 0 Then
                strqry += " And ""GroupCode"" = '" + strCustGroup + "'"
            End If

            If strProperty.Length > 0 Then
                strqry += " And ""QryGroup" + strProperty + """  = 'Y'"
            End If

            oCustomerGrid.DataTable.ExecuteQuery(strqry)

            oCustomerGrid.Columns.Item("Select").TitleObject.Caption = "Select"
            oCustomerGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox


            oCustomerGrid.Columns.Item("CardCode").TitleObject.Caption = "Customer Code"
            oCustomerGrid.Columns.Item("CardCode").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oEditTextColumn = oCustomerGrid.Columns.Item("CardCode")
            oEditTextColumn.LinkedObjectType = "2"
            oCustomerGrid.Columns.Item("CardCode").Editable = False


            oCustomerGrid.Columns.Item("CardName").TitleObject.Caption = "Customer Name"
            oCustomerGrid.Columns.Item("CardName").Editable = False

            oItemGrid = oForm.Items.Item("14").Specific
            oItemGrid.DataTable = oForm.DataSources.DataTables.Item("dtPromotion")

            strqry = " Select T0.""U_Z_PrCode"",T0.""U_Z_PrName"",T0.""U_Z_EffFrom"",T0.""U_Z_EffTo"" "
            strqry = strqry & " ,T1.""U_Z_ItmCode"" ,T1.""U_Z_ItmName"",""U_Z_UOMGroup"",T1.""U_Z_Qty"",T1.""U_Z_OffCode"",T1.""U_Z_OffName"",T1.""U_Z_OQty"",T1.""U_Z_ODis"",""U_Z_OUOMGroup"" "
            strqry = strqry & " From ""@Z_OPRM"" T0 JOIN ""@Z_PRM1"" T1 On T0.""DocEntry"" = T1.""DocEntry""  "
            strqry = strqry & " Where T0.""U_Z_PrCode"" = '" & strPromotion & "'"
            oItemGrid.DataTable.ExecuteQuery(strqry)

            'oForm.Items.Item("29").Specific.value = oItemGrid.DataTable.GetValue("U_EffFrom", 0)
            'oForm.Items.Item("30").Specific.value = oItemGrid.DataTable.GetValue("U_EffTo", 0)

            oItemGrid.Columns.Item("U_Z_PrCode").TitleObject.Caption = "Promotion Code"
            oItemGrid.Columns.Item("U_Z_PrCode").Editable = False
            'oEditTextColumn.LinkedObjectType = 28

            oItemGrid.Columns.Item("U_Z_PrName").TitleObject.Caption = "Promotion Name"
            oItemGrid.Columns.Item("U_Z_PrName").Editable = False

            oItemGrid.Columns.Item("U_Z_EffFrom").TitleObject.Caption = "Effective From"
            oItemGrid.Columns.Item("U_Z_EffFrom").Editable = False

            oItemGrid.Columns.Item("U_Z_EffTo").TitleObject.Caption = "Effective To"
            oItemGrid.Columns.Item("U_Z_EffTo").Editable = False

            oItemGrid.Columns.Item("U_Z_ItmCode").TitleObject.Caption = "Item Code"
            oItemGrid.Columns.Item("U_Z_ItmCode").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oEditTextColumn = oItemGrid.Columns.Item("U_Z_ItmCode")
            oEditTextColumn.LinkedObjectType = "4"
            oItemGrid.Columns.Item("U_Z_ItmCode").Editable = False


            oItemGrid.Columns.Item("U_Z_ItmName").TitleObject.Caption = "Item Name"
            oItemGrid.Columns.Item("U_Z_ItmName").Editable = False

            oItemGrid.Columns.Item("U_Z_Qty").TitleObject.Caption = "Quantity"
            oItemGrid.Columns.Item("U_Z_Qty").Editable = False
            oItemGrid.Columns.Item("U_Z_Qty").RightJustified = True

            oItemGrid.Columns.Item("U_Z_OffCode").TitleObject.Caption = "Offer Item"
            oItemGrid.Columns.Item("U_Z_OffCode").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
            oEditTextColumn = oItemGrid.Columns.Item("U_Z_OffCode")
            oEditTextColumn.LinkedObjectType = "4"
            oItemGrid.Columns.Item("U_Z_OffCode").Editable = False

            oItemGrid.Columns.Item("U_Z_OffName").TitleObject.Caption = "Offer Name"
            oItemGrid.Columns.Item("U_Z_OffName").Editable = False

            oItemGrid.Columns.Item("U_Z_OQty").TitleObject.Caption = "Offer Qty"
            oItemGrid.Columns.Item("U_Z_OQty").Editable = False
            oItemGrid.Columns.Item("U_Z_OQty").RightJustified = True

            oItemGrid.Columns.Item("U_Z_ODis").TitleObject.Caption = "Offer Discount % "
            oItemGrid.Columns.Item("U_Z_ODis").Editable = False
            oItemGrid.Columns.Item("U_Z_ODis").RightJustified = True

            oItemGrid.Columns.Item("U_Z_UOMGroup").TitleObject.Caption = "Unit of Mesurement"
            oItemGrid.Columns.Item("U_Z_UOMGroup").Editable = False

            oItemGrid.Columns.Item("U_Z_OUOMGroup").TitleObject.Caption = "Unit of Mesurement"
            oItemGrid.Columns.Item("U_Z_OUOMGroup").Editable = False


            oItemGrid.CollapseLevel = 1

            oItemGrid.AutoResizeColumns()
            oItemGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Private Sub FillCombo(ByVal aForm As SAPbouiCOM.Form)
        Try
            Dim oTempRec As SAPbobsCOM.Recordset
            oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


            oCombobox = aForm.Items.Item("10").Specific
            For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
                oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
            oCombobox.ValidValues.Add("", "")
            oTempRec.DoQuery("Select ""GroupCode"",""GroupName"" From OCRG Where ""GroupType"" = 'C'")
            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                oCombobox.ValidValues.Add(oTempRec.Fields.Item("GroupCode").Value, oTempRec.Fields.Item("GroupName").Value)
                oTempRec.MoveNext()
            Next

            oCombobox = aForm.Items.Item("26").Specific
            For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
                oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
            oCombobox.ValidValues.Add("", "")
            oTempRec.DoQuery("Select ""GroupCode"",""GroupName"" From OCQG")
            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                oCombobox.ValidValues.Add(oTempRec.Fields.Item("GroupCode").Value, oTempRec.Fields.Item("GroupName").Value)
                oTempRec.MoveNext()
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function InsertPromotionMapping(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim _retVal As Boolean = True
        Try
            Dim oOCPR As SAPbobsCOM.UserTable
            oOCPR = oApplication.Company.UserTables.Item("Z_OCPR")
            oCustomerGrid = oForm.Items.Item("11").Specific
            oItemGrid = oForm.Items.Item("14").Specific

            oApplication.Company.StartTransaction()

            For index As Integer = 0 To oCustomerGrid.Rows.Count - 1

                For indexItem As Integer = 0 To oItemGrid.Rows.Count - 1
                    If Not oItemGrid.Rows.IsLeaf(indexItem) Then
                        Dim intStatus As Integer

                        Dim strCode As String = oApplication.Utilities.getPromotionCode(oCustomerGrid.DataTable.GetValue("CardCode", index), oItemGrid.DataTable.GetValue("U_Z_PrCode", indexItem))
                        If strCode <> "" Then
                            If oOCPR.GetByKey(strCode) Then
                                oOCPR.UserFields.Fields.Item("U_Z_Active").Value = oCustomerGrid.DataTable.GetValue("Select", index)
                                intStatus = oOCPR.Update
                                If intStatus <> 0 Then
                                    _retVal = False
                                    Throw New Exception(oApplication.Company.GetLastErrorDescription())
                                End If
                            End If
                        Else
                            Dim intCode As Integer = oApplication.Utilities.getMaxCode("@Z_OCPR", "Code")
                            oOCPR.Code = intCode.ToString()
                            oOCPR.Name = intCode.ToString()
                            oOCPR.UserFields.Fields.Item("U_Z_PrCode").Value = oItemGrid.DataTable.GetValue("U_Z_PrCode", indexItem)
                            oOCPR.UserFields.Fields.Item("U_Z_CustCode").Value = oCustomerGrid.DataTable.GetValue("CardCode", index)
                            oOCPR.UserFields.Fields.Item("U_Z_EffFrom").Value = oItemGrid.DataTable.GetValue("U_Z_EffFrom", indexItem)
                            oOCPR.UserFields.Fields.Item("U_Z_EffTo").Value = oItemGrid.DataTable.GetValue("U_Z_EffTo", indexItem)
                         
                            oOCPR.UserFields.Fields.Item("U_Z_Active").Value = oCustomerGrid.DataTable.GetValue("Select", index)
                            intStatus = oOCPR.Add()
                            If intStatus <> 0 Then
                                _retVal = False
                                Throw New Exception(oApplication.Company.GetLastErrorDescription())
                            End If
                        End If

                    End If
                Next
            Next

            If oApplication.Company.InTransaction Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If
            Return _retVal
        Catch ex As Exception
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Private Sub reDrawForm(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)

            oForm.Items.Item("11").Top = oForm.Items.Item("13").Top + oForm.Items.Item("13").Height + 1
            oForm.Items.Item("11").Height = (oForm.Height - 150) / 2

            oForm.Items.Item("24").Top = oForm.Items.Item("11").Top + oForm.Items.Item("11").Height + 2
            oForm.Items.Item("14").Top = oForm.Items.Item("24").Top + oForm.Items.Item("24").Height + 1
            oForm.Items.Item("14").Height = (oForm.Height - 160) / 2

            oForm.Items.Item("11").Width = oForm.Width - 20
            oForm.Items.Item("14").Width = oForm.Width - 20

            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

    Private Sub changeLabel(ByVal oForm As SAPbouiCOM.Form)
        Try
            oStatic = oForm.Items.Item("17").Specific
            oStatic.Caption = "Step " & oForm.PaneLevel & " of 3"
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

End Class


'oOCPR.UserFields.Fields.Item("U_Z_ItmCode").Value = oItemGrid.DataTable.GetValue("U_Z_ItmCode", indexItem)
'oOCPR.UserFields.Fields.Item("U_Z_ItmName").Value = oItemGrid.DataTable.GetValue("U_Z_ItmName", indexItem)
'oOCPR.UserFields.Fields.Item("U_Z_Qty").Value = oItemGrid.DataTable.GetValue("U_Z_Qty", indexItem)
'oOCPR.UserFields.Fields.Item("U_Z_OffCode").Value = oItemGrid.DataTable.GetValue("U_Z_OffCode", indexItem)
'oOCPR.UserFields.Fields.Item("U_Z_OffName").Value = oItemGrid.DataTable.GetValue("U_Z_OffName", indexItem)
'oOCPR.UserFields.Fields.Item("U_Z_OQty").Value = oItemGrid.DataTable.GetValue("U_Z_OQty", indexItem)
'oOCPR.UserFields.Fields.Item("U_Z_ODis").Value = oItemGrid.DataTable.GetValue("U_Z_ODis", indexItem)
'oOCPR.UserFields.Fields.Item("U_Z_UOMGroup").Value = oItemGrid.DataTable.GetValue("U_Z_UOMGroup", indexItem)
'oOCPR.UserFields.Fields.Item("U_Z_OUOMGroup").Value = oItemGrid.DataTable.GetValue("U_Z_OUOMGroup", indexItem)