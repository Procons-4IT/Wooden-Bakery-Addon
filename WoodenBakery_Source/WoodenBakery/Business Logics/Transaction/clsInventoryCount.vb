Public Class clsInventoryCount
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
    Private oLoadForm As SAPbouiCOM.Form

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Public Sub LoadForm()
        Try
            oForm = oApplication.Utilities.LoadForm(xml_Z_OICT, frm_Z_OICT)
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
            oForm = oApplication.Utilities.LoadForm(xml_Z_OICT, frm_Z_OICT)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            initialize(oForm)
            oForm.EnableMenu(mnu_ADD_ROW, True)
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            oForm.Freeze(False)
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            oForm.Items.Item("6_").Specific.value = strDocEntry
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

            If pVal.FormTypeEx = frm_Z_OICT Then
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
                                ElseIf pVal.ItemUID = "13" Then 'Browse
                                    oApplication.Utilities.OpenFileDialogBox(oForm, "15")
                                ElseIf (pVal.ItemUID = "14") Then 'Import
                                    If CType(oForm.Items.Item("15").Specific, SAPbouiCOM.StaticText).Caption <> "" Then
                                        If oApplication.Utilities.ValidateFile(oForm, "15") Then
                                            oMatrix = oForm.Items.Item("3").Specific
                                            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_ICT1")
                                            'oLoadForm = Nothing
                                            'oLoadForm = oApplication.Utilities.LoadMessageForm(xml_Load, frm_Load)
                                            'oLoadForm = oApplication.SBO_Application.Forms.ActiveForm()
                                            'oLoadForm.Items.Item("3").TextStyle = 4
                                            'oLoadForm.Items.Item("4").TextStyle = 5
                                            'CType(oLoadForm.Items.Item("3").Specific, SAPbouiCOM.StaticText).Caption = "PLEASE WAIT..."
                                            'CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Importing..."
                                            If oApplication.Utilities.GetData(oForm, oLoadForm, "15", oMatrix, oDBDataSourceLines) Then
                                                oApplication.Utilities.Message("Inventory Data Imported Successfully....", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            Else
                                                BubbleEvent = False
                                            End If
                                        End If
                                    Else
                                        oApplication.Utilities.Message("Select File to Import....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OICT")
                                If pVal.ItemUID = "3" Then
                                    intSelectedMatrixrow = pVal.Row
                                    If (oDBDataSource.GetValue("U_Z_CntDate", 0).ToString() = "") Then
                                        BubbleEvent = False
                                        oApplication.SBO_Application.SetStatusBarMessage("Select Count Date to Proceed...", SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                                        oForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    End If
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
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
                                oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OICT")
                                oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_ICT1")
                                oMatrix = oForm.Items.Item("3").Specific
                                oMatrix.FlushToDataSource()
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oDataTable As SAPbouiCOM.DataTable
                                Dim strCode, strName, strCustomer, strCustName As String
                                Try
                                    oCFLEvento = pVal
                                    oDataTable = oCFLEvento.SelectedObjects

                                    If Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                        If (pVal.ItemUID = "3" And (pVal.ColUID = "V_0" Or pVal.ColUID = "V_1")) Then
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
                                                Dim strUOM As String = " Select ""UomCode"" From OUOM Where ""UomEntry"" = '" & oDataTable.GetValue("UgpEntry", index) & "'"
                                                Dim oUOMRS As SAPbobsCOM.Recordset
                                                oUOMRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                oUOMRS.DoQuery(strUOM)
                                                If Not oUOMRS.EoF Then
                                                    oDBDataSourceLines.SetValue("U_Z_UOM", pVal.Row + index - 1, oUOMRS.Fields.Item(0).Value)
                                                End If
                                            Next
                                            oMatrix.LoadFromDataSource()
                                            oMatrix.Columns.Item("V_2").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        ElseIf (pVal.ItemUID = "3" And (pVal.ColUID = "V_2")) Then
                                            oMatrix = oForm.Items.Item("3").Specific
                                            For index As Integer = 0 To oDataTable.Rows.Count - 1
                                                oDBDataSourceLines.SetValue("U_Z_WareHouse", pVal.Row + index - 1, oDataTable.GetValue("WhsCode", index))
                                            Next
                                            oMatrix.LoadFromDataSource()
                                            oMatrix.Columns.Item("V_3").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                        End If
                                    End If
                                Catch ex As Exception

                                End Try
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("3").Specific
                                If pVal.ItemUID = "3" And pVal.ColUID = "V_4" And pVal.CharPressed = 9 And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    Dim strItemCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", pVal.Row)
                                    'filterUOMChooseFromList(oForm, "CFL_5", strItemCode)
                                    Dim objChooseForm As SAPbouiCOM.Form
                                    Dim objChoose As New clsChooseFromList_UOM
                                    objChoose.ItemUID = pVal.ItemUID
                                    objChoose.SourceFormUID = FormUID
                                    objChoose.SourceLabel = 0 'pVal.Row
                                    objChoose.CFLChoice = "I"
                                    objChoose.choice = "INVENTORY"
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
                            Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                If oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized Or oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Minimized Or oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Restore Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    Try
                                        reDrawForm(oForm)
                                    Catch ex As Exception

                                    End Try
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
                        Case mnu_Cancel, mnu_CLOSE
                            BubbleEvent = False
                            oApplication.Utilities.Message("Not Possible to Cancel or Close the Document...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Sub
                        Case mnu_ICCancel
                            If oApplication.SBO_Application.MessageBox("Do You Want to Cancel Inventory Count Document?", 2, "Yes", "No") = 2 Then
                                BubbleEvent = False
                                Exit Sub
                            Else
                                Dim DocEntry As String = oDBDataSource.GetValue("DocEntry", oDBDataSource.Offset)
                                oApplication.Utilities.changeStatus(oForm, "L")
                                oApplication.SBO_Application.Menus.Item(mnu_ADD).Activate()
                            End If
                        Case mnu_DELETE_ROW
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            oMatrix = oForm.Items.Item("3").Specific
                            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            If intSelectedMatrixrow > 0 Then
                                Dim strItemCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", intSelectedMatrixrow)
                                strQuery = "Select ""DocEntry"" From OINC "
                                strQuery += " Where ""U_Z_ICTREF"" = '" & oApplication.Utilities.getEdittextvalue(oForm, "6_") & "'"
                                oRecordSet.DoQuery(strQuery)
                                If Not oRecordSet.EoF Then
                                    BubbleEvent = False
                                    oApplication.Utilities.Message("Row Already linked to Inventory Count Document...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
                        Case mnu_Z_OICT
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
            If oForm.TypeEx = frm_Z_OICT Then
                Select Case BusinessObjectInfo.BeforeAction
                    Case True

                    Case False
                        Select Case BusinessObjectInfo.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                                Dim stXML As String = BusinessObjectInfo.ObjectKey
                                stXML = stXML.Replace("<?xml version=""1.0"" encoding=""UTF-16"" ?><Inventory_CountParams><DocEntry>", "")
                                stXML = stXML.Replace("</DocEntry></Inventory_CountParams>", "")
                                stXML = stXML.Replace("<?xml version=""1.0"" encoding=""UTF-16"" ?><Inventory_CountParams><DocEntry>", "")
                                stXML = stXML.Replace("</DocEntry></Inventory_CountParams>", "")
                                oApplication.Utilities.CreateInventoryCountDocument(oForm, stXML.Trim())
                                oApplication.Utilities.checkAllDocumentStatus(oForm, stXML)
                            Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                                EnableControls(oForm, False)
                                oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OICT")
                                oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_ICT1")

                                If (oDBDataSource.GetValue("U_Z_Status", oDBDataSource.Offset) = "L" _
                                    Or oDBDataSource.GetValue("U_Z_Status", oDBDataSource.Offset) = "C") Then
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                                End If
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
            If oForm.TypeEx = frm_Z_OICT Then
                intSelectedMatrixrow = eventInfo.Row
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub

#Region "Function"

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OICT")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_ICT1")

            oMatrix = oForm.Items.Item("3").Specific
            oMatrix.LoadFromDataSource()
            oMatrix.AddRow(1, -1)
            oMatrix.FlushToDataSource()

            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If blnIsHanaDB = True Then
                oRecordSet.DoQuery("Select IfNull(MAX(""DocEntry""),1)+1 From ""@Z_OICT""")
            Else
                oRecordSet.DoQuery("Select IsNull(MAX(""DocEntry""),1)+1 From ""@Z_OICT""")
            End If

            If Not oRecordSet.EoF Then
                oDBDataSource.SetValue("DocNum", 0, oRecordSet.Fields.Item(0).Value.ToString())
            End If
            oApplication.Utilities.setEdittextvalue(oForm, "7", System.DateTime.Now.ToString("yyyyMMdd"))
            oForm.Items.Item("5").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
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
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_ICT1")
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
                Case "0", "1"
                    oMatrix = aForm.Items.Item("3").Specific
                    oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_ICT1")
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
                    oDBDataSourceLines = aForm.DataSources.DBDataSources.Item("@Z_ICT1")
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
                            oDBDataSourceLines = aForm.DataSources.DBDataSources.Item("@Z_ICT1")
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
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OICT")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_ICT1")

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
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_OICT")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_ICT1")

            Dim oTest As SAPbobsCOM.Recordset
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oApplication.Utilities.getEdittextvalue(aForm, "5") = "" Then
                oApplication.Utilities.Message("Enter Count Date...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            oMatrix = oForm.Items.Item("3").Specific
            For index As Integer = 1 To oMatrix.VisualRowCount

                Dim strItemCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_0", index)
                Dim strWhsCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_2", index)
                Dim strUOM As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_4", index)
                Dim strQty As String = oApplication.Utilities.getMatrixValues(oMatrix, "V_3", index)

                Dim dblQty As Double

                If strItemCode.Length > 0 Then
                    If strWhsCode.Trim().Length = 0 Then
                        oApplication.Utilities.Message("Enter Ware House for Row No " & index.ToString, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf strUOM.Trim().Length = 0 Then
                        oApplication.Utilities.Message("Enter UOM for Row No " & index.ToString, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    ElseIf Double.TryParse(strQty, dblQty) Then
                        If dblQty = 0 Then
                            oApplication.Utilities.Message("Enter Quantity for Row No " & index.ToString, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                End If

            Next


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
            'oForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            'oForm.Items.Item("6").Enabled = blnEnable
            'oForm.Items.Item("7").Enabled = blnEnable
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

    Private Sub reDrawForm(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)

            oForm.Items.Item("4").Height = (oForm.Items.Item("3").Height) + 10
            oForm.Items.Item("4").Width = (oForm.Items.Item("3").Width) + 10

            oForm.Freeze(False)
        Catch ex As Exception
            'oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub


#End Region

End Class
