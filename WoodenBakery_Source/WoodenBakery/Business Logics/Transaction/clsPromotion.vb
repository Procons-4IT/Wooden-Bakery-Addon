Public Class clsPromotion
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private InvForConsumedItems, count As Integer
    Dim oDBDataSource As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines As SAPbouiCOM.DBDataSource
    Public intSelectedMatrixrow As Integer = 0
    Private RowtoDelete As Integer
    Private MatrixId As String
    Private oRecordSet As SAPbobsCOM.Recordset
    Private dtValidFrom, dtValidTo As Date
    Private strQuery As String
    Private oCombo As SAPbouiCOM.ComboBox

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Public Sub LoadForm()
        Try
            oForm = oApplication.Utilities.LoadForm(xml_Z_OPRM, frm_Z_OPRM)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            initialize(oForm)
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            oForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

    Public Sub LoadForm(ByVal strDocEntry As String)
        Try
            oForm = oApplication.Utilities.LoadForm(xml_Z_OPRM, frm_Z_OPRM)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            initialize(oForm)
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            oForm.Freeze(False)
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            oForm.Items.Item("16").Specific.value = strDocEntry
            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try

            If pVal.FormTypeEx = frm_Z_OPRM Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        If oApplication.SBO_Application.MessageBox("Do you want to confirm the information?", , "Yes", "No") = 2 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        Else
                                            If validation(oForm) = False Then
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                If pVal.ItemUID = "3" Then
                                    oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPRM")
                                    intSelectedMatrixrow = pVal.Row
                                    If (oDBDataSource.GetValue("U_Z_PrCode", 0).ToString() = "") Then
                                        BubbleEvent = False
                                        oApplication.SBO_Application.SetStatusBarMessage("Select Promotion Code to Proceed...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        oForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    ElseIf (oDBDataSource.GetValue("U_Z_EffFrom", 0).ToString() = "" Or oDBDataSource.GetValue("U_Z_EffTo", 0).ToString() = "") Then
                                        BubbleEvent = False
                                        oApplication.SBO_Application.SetStatusBarMessage("Enter Effective From & To Date to Proceed...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        oForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    End If
                                End If
                                'Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                '    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                '    oMatrix = oForm.Items.Item("3").Specific
                                '    If pVal.ItemUID = "3" And pVal.ColUID = "V_2_" Then
                                '        Dim strItemCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", pVal.Row)
                                '        'filterUOMChooseFromList(oForm, "CFL_5", strItemCode)
                                '    End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "3" And pVal.ColUID = "V_8" Then
                                    oMatrix = oForm.Items.Item("3").Specific
                                    oCombo = oMatrix.Columns.Item("V_8").Cells().Item(pVal.Row).Specific
                                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PRM1")
                                    oMatrix.FlushToDataSource()
                                    If oCombo.Selected.Value = "I" Then
                                        oDBDataSourceLines.SetValue("U_Z_Dis", pVal.Row - 1, "")
                                    ElseIf oCombo.Selected.Value = "D" Then
                                        oDBDataSourceLines.SetValue("U_Z_OffCode", pVal.Row - 1, "")
                                        oDBDataSourceLines.SetValue("U_Z_OffName", pVal.Row - 1, "")
                                        oDBDataSourceLines.SetValue("U_Z_OUOMGroup", pVal.Row - 1, "")
                                        oDBDataSourceLines.SetValue("U_Z_OQty", pVal.Row - 1, "")
                                    End If
                                    oMatrix.LoadFromDataSource()
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "14"
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_ADD_ROW)
                                    Case "15"
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_DELETE_ROW)
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPRM")
                                oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PRM1")
                                oMatrix = oForm.Items.Item("3").Specific
                                oMatrix.FlushToDataSource()
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Dim strCode, strName, strCustomer, strCustName As String
                                Try
                                    oCFLEvento = pVal
                                    oDataTable = oCFLEvento.SelectedObjects

                                    If Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If pVal.ItemUID = "6" Then
                                            strCode = oDataTable.GetValue("PrCode", 0)
                                            strName = oDataTable.GetValue("PrName", 0)
                                            strCustomer = oDataTable.GetValue("U_CardCode", 0)
                                            strCustName = oDataTable.GetValue("U_CardName", 0)
                                            Try
                                                oDBDataSource.SetValue("U_PrCode", oDBDataSource.Offset, strCode)
                                                oDBDataSource.SetValue("U_PrName", oDBDataSource.Offset, strName)
                                            Catch ex As Exception

                                            End Try
                                        ElseIf (pVal.ItemUID = "3" And (pVal.ColUID = "V_0" Or pVal.ColUID = "V_1")) Then
                                            oMatrix = oForm.Items.Item("3").Specific
                                            oMatrix.LoadFromDataSource()
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If intAddRows > 1 Then
                                                intAddRows -= 1
                                                oMatrix.AddRow(intAddRows, pVal.Row - 1)
                                            End If
                                            oMatrix.FlushToDataSource()
                                            For index As Integer = 0 To oDataTable.Rows.Count - 1
                                                oDBDataSourceLines.SetValue("LineId", pVal.Row + index - 1, (pVal.Row + index).ToString())
                                                oDBDataSourceLines.SetValue("U_Z_ItmCode", pVal.Row + index - 1, oDataTable.GetValue("ItemCode", index))
                                                oDBDataSourceLines.SetValue("U_Z_ItmName", pVal.Row + index - 1, oDataTable.GetValue("ItemName", index))
                                                oDBDataSourceLines.SetValue("U_Z_Qty", pVal.Row + index - 1, "1")
                                                Dim strUOM As String
                                                If oDataTable.GetValue("UgpEntry", index) <> "-1" Then
                                                    strUOM = " Select ""UomCode"" From OUGP Where ""UgpEntry"" = '" & oDataTable.GetValue("UgpEntry", index) & "'"
                                                Else
                                                    strUOM = " Select ""UomCode"" From OUOM Where ""UomCode"" = '" & oDataTable.GetValue("SalUnitMsr", index) & "'"
                                                End If
                                                oDBDataSourceLines.SetValue("U_Z_UOMGROUP", pVal.Row + index - 1, oDataTable.GetValue("SalUnitMsr", index))
                                                ''  SUoMEntry()
                                                ''strUOM = " SELECT  T2.""UomEntry"" as ""UomEntry"", T3.""UomCode"" as ""UomCode"" FROM OITM T0 INNER JOIN OUGP T1 ON T0.""UgpEntry"" = T1.""UgpEntry"" "
                                                ''strUOM += " INNER JOIN UGP1 T2 ON T1.""UgpEntry"" = T2.""UgpEntry"" INNER JOIN OUOM T3 ON T3.""UomEntry"" = T2.""UomEntry"" "
                                                ''strUOM += " Where T0.""ItemCode"" = '" & oDataTable.GetValue("ItemCode", index) & "' "
                                                'strUOM = " SELECT  T2.""UomEntry"" as ""UomEntry"", T3.""UomCode"" as ""UomCode"" FROM OITM T0 INNER JOIN OUGP T1 ON T0.""SUoMEntry"" = T1.""UgpEntry"" "
                                                'strUOM += " INNER JOIN UGP1 T2 ON T1.""UgpEntry"" = T2.""UgpEntry"" INNER JOIN OUOM T3 ON T3.""UomEntry"" = T2.""UomEntry"" "
                                                'strUOM += " Where T0.""ItemCode"" = '" & oDataTable.GetValue("ItemCode", index) & "' "

                                                'Dim oUOMRS As SAPbobsCOM.Recordset
                                                'oUOMRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                'oUOMRS.DoQuery(strUOM)
                                                'If Not oUOMRS.EoF Then
                                                '    oDBDataSourceLines.SetValue("U_Z_UOMGroup", pVal.Row + index - 1, oUOMRS.Fields.Item(1).Value)
                                                'End If
                                            Next
                                            oMatrix.LoadFromDataSource()
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ElseIf (pVal.ItemUID = "3" And pVal.ColUID = "V_3") Then
                                            oMatrix = oForm.Items.Item("3").Specific
                                            oMatrix.LoadFromDataSource()
                                            Dim intAddRows As Integer = oDataTable.Rows.Count
                                            If intAddRows > 1 Then
                                                intAddRows -= 1
                                                oMatrix.AddRow(intAddRows, pVal.Row - 1)
                                            End If
                                            oMatrix.FlushToDataSource()
                                            For index As Integer = 0 To oDataTable.Rows.Count - 1
                                                oDBDataSourceLines.SetValue("LineId", pVal.Row + index - 1, (pVal.Row + index).ToString())
                                                oDBDataSourceLines.SetValue("U_Z_OffCode", pVal.Row + index - 1, oDataTable.GetValue("ItemCode", index))
                                                oDBDataSourceLines.SetValue("U_Z_OffName", pVal.Row + index - 1, oDataTable.GetValue("ItemName", index))
                                                oDBDataSourceLines.SetValue("U_Z_OQty", pVal.Row + index - 1, "1")
                                                Dim strUOM As String = " Select ""UomCode"" From OUOM Where ""UomEntry"" = '" & oDataTable.GetValue("UgpEntry", index) & "'"
                                                Dim oUOMRS As SAPbobsCOM.Recordset
                                                oUOMRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                oDBDataSourceLines.SetValue("U_Z_OUOMGroup", pVal.Row + index - 1, oDataTable.GetValue("SalUnitMsr", index))


                                                'strUOM = " SELECT  T2.""UomEntry"" as ""UomEntry"", T3.""UomCode"" as ""UomCode"" FROM OITM T0 INNER JOIN OUGP T1 ON T0.""UgpEntry"" = T1.""UgpEntry"" "
                                                'strUOM += " INNER JOIN UGP1 T2 ON T1.""UgpEntry"" = T2.""UgpEntry"" INNER JOIN OUOM T3 ON T3.""UomEntry"" = T2.""UomEntry"" "
                                                'strUOM += " Where T0.""ItemCode"" = '" & oDataTable.GetValue("ItemCode", index) & "' "


                                                'oUOMRS.DoQuery(strUOM)
                                                'If Not oUOMRS.EoF Then
                                                '    oDBDataSourceLines.SetValue("U_Z_OUOMGroup", pVal.Row + index - 1, oUOMRS.Fields.Item(1).Value)
                                                'End If
                                            Next
                                            oMatrix.LoadFromDataSource()
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ElseIf (pVal.ItemUID = "3" And pVal.ColUID = "V_2_") Then
                                            oMatrix = oForm.Items.Item("3").Specific
                                            oMatrix.LoadFromDataSource()
                                            For index As Integer = 0 To oDataTable.Rows.Count - 1
                                                oDBDataSourceLines.SetValue("U_Z_UOMGroup", pVal.Row + index - 1, oDataTable.GetValue("UomCode", index))
                                            Next
                                            oMatrix.LoadFromDataSource()
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ElseIf (pVal.ItemUID = "3" And pVal.ColUID = "V_5_") Then
                                            oMatrix = oForm.Items.Item("3").Specific
                                            oMatrix.LoadFromDataSource()
                                            For index As Integer = 0 To oDataTable.Rows.Count - 1
                                                oDBDataSourceLines.SetValue("U_Z_OUOMGroup", pVal.Row + index - 1, oDataTable.GetValue("UomCode", index))
                                            Next
                                            oMatrix.LoadFromDataSource()
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        End If
                                    End If
                                Catch ex As Exception

                                End Try
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("3").Specific
                                If pVal.ItemUID = "3" And pVal.ColUID = "V_2_" And pVal.CharPressed = 9 And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    Dim strItemCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", pVal.Row)
                                    'filterUOMChooseFromList(oForm, "CFL_5", strItemCode)
                                    Dim objChooseForm As SAPbouiCOM.Form
                                    Dim objChoose As New clsChooseFromList_UOM
                                    objChoose.ItemUID = pVal.ItemUID
                                    objChoose.SourceFormUID = FormUID
                                    objChoose.SourceLabel = 0 'pVal.Row
                                    objChoose.CFLChoice = "I"
                                    objChoose.choice = "PROMOTION"
                                    objChoose.ItemCode = strItemCode
                                    objChoose.Documentchoice = "" ' oApplication.Utilities.GetDocType(oForm)
                                    If pVal.ItemUID = "13" Then
                                        objChoose.sourceColumID = "28"
                                    Else
                                        objChoose.sourceColumID = "29"
                                    End If

                                    objChoose.sourcerowId = pVal.Row
                                    objChoose.BinDescrUID = ""
                                    oApplication.Utilities.LoadForm("CFL_UOM.xml", frm_ChoosefromList_UOM)
                                    objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                    objChoose.databound(objChooseForm)
                                ElseIf pVal.ItemUID = "3" And pVal.ColUID = "V_5_" And pVal.CharPressed = 9 And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    Dim strItemCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", pVal.Row)
                                    'filterUOMChooseFromList(oForm, "CFL_5", strItemCode)
                                    Dim objChooseForm As SAPbouiCOM.Form
                                    Dim objChoose As New clsChooseFromList_UOM
                                    objChoose.ItemUID = pVal.ItemUID
                                    objChoose.SourceFormUID = FormUID
                                    objChoose.SourceLabel = 0 'pVal.Row
                                    objChoose.CFLChoice = "O"
                                    objChoose.choice = "PROMOTION"
                                    objChoose.ItemCode = strItemCode
                                    objChoose.Documentchoice = "" ' oApplication.Utilities.GetDocType(oForm)
                                    If pVal.ItemUID = "13" Then
                                        objChoose.sourceColumID = "28"
                                    Else
                                        objChoose.sourceColumID = "29"
                                    End If

                                    objChoose.sourcerowId = pVal.Row
                                    objChoose.BinDescrUID = ""
                                    oApplication.Utilities.LoadForm("CFL_UOM.xml", frm_ChoosefromList_UOM)
                                    objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                    objChoose.databound(objChooseForm)
                                End If
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
            Select Case pVal.BeforeAction
                Case True
                    Select Case pVal.MenuUID
                        Case mnu_DELETE_ROW
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            oMatrix = oForm.Items.Item("3").Specific
                            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            If intSelectedMatrixrow > 0 Then
                                Dim strItemCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intSelectedMatrixrow)
                                strQuery = "Select ""DocEntry"" From RDR1 "
                                strQuery += " Where ""U_Z_PrCode"" = '" & oApplication.Utilities.getEdittextvalue(oForm, "6") & "'"
                                strQuery += " And ""ItemCode"" = '" & strItemCode & "'"
                                oRecordSet.DoQuery(strQuery)
                                If Not oRecordSet.EoF Then
                                    BubbleEvent = False
                                    oApplication.Utilities.Message("Item Already linked to Sale Promotion...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Exit Sub
                                End If
                            Else
                                BubbleEvent = False
                                oApplication.Utilities.Message("Select Row and delete...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Exit Sub
                            End If
                    End Select
                
                Case False
                    Select Case pVal.MenuUID
                        Case mnu_Z_OPRM
                            LoadForm()
                        Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            If pVal.BeforeAction = False Then

                            End If
                        Case mnu_ADD_ROW
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            If pVal.BeforeAction = False Then
                                AddRow(oForm)
                            End If
                        Case mnu_DELETE_ROW
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            If pVal.BeforeAction = False Then
                                RefereshDeleteRow(oForm)
                            End If
                        Case mnu_ADD
                            If pVal.BeforeAction = False Then
                                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                                initialize(oForm)
                                EnableControls(oForm, True)
                            End If
                        Case mnu_FIND
                            If pVal.BeforeAction = False Then

                            End If
                        Case mnu_CPRL_IP
                            oForm = oApplication.SBO_Application.Forms.ActiveForm
                            If oForm.TypeEx.ToString() = frm_Z_OPRM Then
                                If Not oForm.Items.Item("13").Specific.value = "" Then
                                    Dim objPromList As clsCustPromotionList
                                    objPromList = New clsCustPromotionList
                                    objPromList.LoadForm(oForm.Items.Item("13").Specific.value, "P")
                                Else
                                    oApplication.Utilities.Message("Select Item to Get Promotion List(Customer)...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If
                                Dim oMenuItem As SAPbouiCOM.MenuItem
                                oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                                If oMenuItem.SubMenus.Exists(pVal.MenuUID) Then
                                    oApplication.SBO_Application.Menus.RemoveEx(pVal.MenuUID)
                                End If
                            End If


                    End Select
            End Select
            
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Data Events"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
            If oForm.TypeEx = frm_Z_OPRM Then
                Select Case BusinessObjectInfo.BeforeAction
                    Case True

                    Case False
                        Select Case BusinessObjectInfo.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                EnableControls(oForm, False)
                        End Select
                End Select
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
            If oForm.TypeEx = frm_Z_OPRM Then
                intSelectedMatrixrow = eventInfo.Row
                If (eventInfo.BeforeAction = True) Then
                    Dim oMenuItem As SAPbouiCOM.MenuItem
                    Dim oMenus As SAPbouiCOM.Menus
                    Try
                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            'Promotion List
                            oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                            If Not oMenuItem.SubMenus.Exists(mnu_CPRL_IP) Then
                                Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                                oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                                oCreationPackage.UniqueID = mnu_CPRL_IP
                                oCreationPackage.String = "Promotion List(Customer)"
                                oCreationPackage.Enabled = True
                                oMenus = oMenuItem.SubMenus
                                oMenus.AddEx(oCreationPackage)
                            End If

                        End If
                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try
                Else
                    Dim oMenuItem As SAPbouiCOM.MenuItem
                    Dim oMenus As SAPbouiCOM.Menus
                    Try

                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            '  oApplication.SBO_Application.Menus.RemoveEx("TraDetails")
                        End If

                        oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                        If oMenuItem.SubMenus.Exists(mnu_CPRL_IP) Then
                            oApplication.SBO_Application.Menus.RemoveEx(mnu_CPRL_IP)
                        End If

                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try
                End If
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

#Region "Function"

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPRM")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PRM1")

            oMatrix = oForm.Items.Item("3").Specific
            oMatrix.LoadFromDataSource()
            oMatrix.AddRow(1, -1)
            oMatrix.FlushToDataSource()

            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If blnIsHanaDB = True Then
                oRecordSet.DoQuery("Select IfNull(MAX(""DocEntry""),1) From ""@Z_OPRM""")
            Else
                oRecordSet.DoQuery("Select IsNull(MAX(""DocEntry""),1) From ""@Z_OPRM""")
            End If

            If Not oRecordSet.EoF Then
                oDBDataSource.SetValue("DocNum", 0, oRecordSet.Fields.Item(0).Value.ToString())
            End If
            oDBDataSource.SetValue("U_Z_Active", 0, "Y")
            MatrixId = "3"
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#Region "Methods"
    Private Sub AssignLineNo(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item("3").Specific
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PRM1")
            oMatrix.FlushToDataSource()
            For count = 1 To oDBDataSourceLines.Size
                oDBDataSourceLines.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub
#End Region

    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            Select Case aForm.PaneLevel
                Case "0"
                    oMatrix = aForm.Items.Item("3").Specific
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PRM1")
                    If oMatrix.RowCount <= 0 Then
                        oMatrix.AddRow()
                    Else
                        oMatrix.AddRow(1, oMatrix.RowCount + 1)
                        oMatrix.ClearRowData(oMatrix.RowCount)
                        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                            aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        End If
                    End If
                    oMatrix.FlushToDataSource()
                    For count = 1 To oDBDataSourceLines.Size
                        oDBDataSourceLines.SetValue("LineId", count - 1, count)
                    Next
                    oMatrix.LoadFromDataSource()
                    oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    AssignLineNo(aForm)
            End Select
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub deleterow(ByVal aForm As SAPbouiCOM.Form)
        Try
            Select Case aForm.PaneLevel
                Case "0", "1"
                    oMatrix = aForm.Items.Item("3").Specific
                    oDBDataSourceLines = aForm.DataSources.DBDataSources.Item("@Z_PRM1")
            End Select
            oMatrix.FlushToDataSource()
            For introw As Integer = 1 To oMatrix.RowCount
                If oMatrix.IsRowSelected(introw) Then
                    oMatrix.DeleteRow(introw)
                    oDBDataSourceLines.RemoveRecord(introw - 1)
                    oMatrix.FlushToDataSource()
                    For count As Integer = 1 To oDBDataSourceLines.Size
                        oDBDataSourceLines.SetValue("LineId", count - 1, count)
                    Next
                    Select Case aForm.PaneLevel
                        Case "0", "1"
                            oMatrix = aForm.Items.Item("3").Specific
                            oDBDataSourceLines = aForm.DataSources.DBDataSources.Item("@Z_PRM1")
                            AssignLineNo(aForm)
                    End Select
                    oMatrix.LoadFromDataSource()
                    If aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If
                    Exit Sub
                End If
            Next
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

#Region "Delete Row"
    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            oMatrix = aForm.Items.Item("3").Specific
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPRM")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PRM1")

            'If Me.MatrixId = "3" Then
            '    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PRM1")
            'End If

            Me.RowtoDelete = intSelectedMatrixrow
            oDBDataSourceLines.RemoveRecord(Me.RowtoDelete - 1)
            oMatrix.LoadFromDataSource()
            oMatrix.FlushToDataSource()
            For count = 1 To oDBDataSourceLines.Size
                oDBDataSourceLines.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub
#End Region

#Region "Validations"
    Private Function validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            oMatrix = oForm.Items.Item("3").Specific
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OPRM")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_PRM1")

            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oApplication.Utilities.getEditTextvalue(aForm, "6") = "" Then
                oApplication.Utilities.Message("Enter Promotion Code...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf oApplication.Utilities.getEditTextvalue(aForm, "7") = "" Then
                oApplication.Utilities.Message("Enter Promotion Name...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf oApplication.Utilities.getEditTextvalue(aForm, "10") = "" Then
                oApplication.Utilities.Message("Enter Effective From Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf oApplication.Utilities.getEditTextvalue(aForm, "11") = "" Then
                oApplication.Utilities.Message("Enter Effective To Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            'Dim dtFromDt As Integer
            'Dim dtToDt As Integer
            'dtFromDt = oApplication.Utilities.getEditTextvalue(aForm, "10")
            'dtToDt = oApplication.Utilities.getEditTextvalue(aForm, "11")
            'If dtFromDt > dtToDt Then
            '    oApplication.Utilities.Message("Effective To Date Should be Greater than Effective From Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If

            oMatrix = oForm.Items.Item("3").Specific
            If oMatrix.RowCount <= 0 Then
                oApplication.Utilities.Message("Line Details Missing....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                If oApplication.Utilities.getMatrixValues(oMatrix, "V_0", 1) = "" Then
                    oApplication.Utilities.Message("Line Details Missing....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If


                End If
                For index As Integer = 1 To oMatrix.VisualRowCount
                    Dim strItemCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index)
                    Dim strDisType As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_8", index)
                    Dim strDiscount As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_9", index)
                    Dim dblDiscount As Double

                    If strDisType = "D" Then
                        If Double.TryParse(strDiscount, dblDiscount) Then
                            If dblDiscount = 0 Then
                                oApplication.Utilities.Message("Enter Discount for Row No " & index.ToString, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                            If strItemCode = "" Then
                                oApplication.Utilities.Message("Item Code is missing... Row No : " & index.ToString, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                        End If
                    ElseIf strDisType = "I" Then
                        Dim strOItemCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_3", index)
                        If strItemCode.Length > 0 And strOItemCode.Trim() = "" Then
                            oApplication.Utilities.Message("Enter Offer Item for Row No " & index.ToString, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                Next

                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                If aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    strQuery = "Select 1 As ""Return"",""DocEntry"" From ""@Z_OPRM"" "
                    strQuery += " Where "
                    strQuery += " ""U_Z_PrCode"" = '" + oDBDataSource.GetValue("U_Z_PrCode", 0).Trim() + "' And ""DocEntry"" <> '" + oDBDataSource.GetValue("DocEntry", 0).ToString() + "'"
                    oRecordSet.DoQuery(strQuery)
                    If Not oRecordSet.EoF Then
                        oApplication.Utilities.Message("Promotion Code Already Exist...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                ElseIf aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    strQuery = "Select 1 As ""Return"",""DocEntry"" From ""@Z_OPRM"" "
                    strQuery += " Where "
                    strQuery += " ""U_Z_PrCode"" = '" & oDBDataSource.GetValue("U_Z_PrCode", 0).Trim() & "'"
                    oRecordSet.DoQuery(strQuery)
                    If Not oRecordSet.EoF Then
                        oApplication.Utilities.Message("Promotion Code Already Exist...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If


                Return True
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Function
#End Region

#Region "Disable Controls"

    Private Sub EnableControls(ByVal oForm As SAPbouiCOM.Form, ByVal blnEnable As Boolean)
        Try
            oForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oForm.Items.Item("6").Enabled = blnEnable
            oForm.Items.Item("7").Enabled = blnEnable
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

    'Private Sub filterUOMChooseFromList(ByVal oForm As SAPbouiCOM.Form, ByVal strCFLID As String, ByVal strItemCode As String)
    '    Try
    '        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
    '        Dim oCons As SAPbouiCOM.Conditions
    '        Dim oCon As SAPbouiCOM.Condition
    '        Dim oCFL As SAPbouiCOM.ChooseFromList

    '        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        'strQuery = "Select ""UomEntry"" From ITM12 "
    '        'strQuery += " Where ""ItemCode"" = '" & strItemCode & "' "
    '        '  WHERE T0.""ItemCode""=''

    '        strQuery = " SELECT  T2.""UomEntry"" as ""UomEntry"", T3.""UomCode"" as ""UomCode"" FROM OITM T0 INNER JOIN OUGP T1 ON T0.""UgpEntry"" = T1.""UgpEntry"" "
    '        strQuery += " INNER JOIN UGP1 T2 ON T1.""UgpEntry"" = T2.""UgpEntry"" INNER JOIN OUOM T3 ON T3.""UoMEntry"" = T2.""UomEntry"" "
    '        strQuery += " Where T0.""ItemCode"" = '" & strItemCode & "' "
    '        oRecordSet.DoQuery(strQuery)
    '        If Not oRecordSet.EoF Then

    '            oCFLs = oForm.ChooseFromLists
    '            oCFL = oCFLs.Item(strCFLID)
    '            oCons = oCFL.GetConditions()

    '            Dim blnConExist As Boolean = False
    '            If oCons.Count > 0 Then
    '                oCon = oCons.Add()
    '                oCon.BracketOpenNum = 2
    '                blnConExist = True
    '            End If
    '            For intRow As Integer = 0 To oCons.Count - 1
    '                oCon = oCons.Item(intRow)
    '                oCon.Alias = "UomEntry"
    '                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '                oCon.CondVal = "-1"
    '            Next
    '            If blnConExist Then
    '                oCon.BracketCloseNum = 2
    '            End If
    '            If blnConExist Then
    '                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
    '            End If

    '            oCon = oCons.Add()
    '            oCon.BracketOpenNum = 2
    '            Dim intConCount As Integer = 0

    '            While Not oRecordSet.EoF
    '                Dim strIG As String = oRecordSet.Fields.Item(0).Value
    '                If intConCount > 0 Then
    '                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
    '                    oCon = oCons.Add()
    '                    oCon.BracketOpenNum = 1
    '                End If
    '                oCon.[Alias] = "UomEntry"
    '                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
    '                oCon.CondVal = strIG

    '                oRecordSet.MoveNext()
    '                If Not oRecordSet.EoF Then
    '                    oCon.BracketCloseNum = 1
    '                End If

    '                intConCount += 1
    '            End While

    '            oCon.BracketCloseNum = 2
    '            oCFL.SetConditions(oCons)

    '        End If


    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

#End Region

End Class
