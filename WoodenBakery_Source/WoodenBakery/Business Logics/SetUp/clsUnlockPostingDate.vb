Public Class clsUnlockPostingDate
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
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
    Dim oDataSrc_Line_User, oDataSrc_Line As SAPbouiCOM.DBDataSource
    Private RowtoDelete As Integer
    Private oMenuobject As Object
    Private count As Integer
    Dim MatrixId As Integer
    Private oColumn As SAPbouiCOM.Column

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Public Sub LoadForm()
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_UnLock) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_UnLock, frm_UnLock)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.DataBrowser.BrowseBy = "4"
        oForm.EnableMenu(mnu_ADD_ROW, True)
        oForm.EnableMenu(mnu_DELETE_ROW, True)
        AddChooseFromList(oForm)
        'oForm.EnableMenu("1283", True)
        oEditText = oForm.Items.Item("4").Specific
        oEditText.ChooseFromListUID = "CFL1"
        oEditText.ChooseFromListAlias = "USER_CODE"

        oMatrix = oForm.Items.Item("7").Specific
        oColumn = oMatrix.Columns.Item("V_0")
        oColumn.ChooseFromListUID = "CFL2"
        oColumn.ChooseFromListAlias = "U_Z_Code"
        oDataSrc_Line_User = oForm.DataSources.DBDataSources.Item("@Z_LUSR1")
        For count = 1 To oDataSrc_Line_User.Size - 1
            oDataSrc_Line_User.SetValue("LineId", count - 1, count)
        Next
        oMatrix = oForm.Items.Item("7").Specific
        oMatrix.AutoResizeColumns()
        oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

        oMatrix = oForm.Items.Item("17").Specific
        oColumn = oMatrix.Columns.Item("V_0")
        oColumn.ChooseFromListUID = "CFL21"
        oColumn.ChooseFromListAlias = "U_Z_Code"
        oDataSrc_Line_User = oForm.DataSources.DBDataSources.Item("@Z_LUSR2")
        For count = 1 To oDataSrc_Line_User.Size - 1
            oDataSrc_Line_User.SetValue("LineId", count - 1, count)
        Next
        oMatrix = oForm.Items.Item("17").Specific
        oMatrix.AutoResizeColumns()
        oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

        oMatrix = oForm.Items.Item("18").Specific
        oColumn = oMatrix.Columns.Item("V_0")
        oColumn.ChooseFromListUID = "CFL22"
        oColumn.ChooseFromListAlias = "U_Z_Code"
        oDataSrc_Line_User = oForm.DataSources.DBDataSources.Item("@Z_LUSR3")
        For count = 1 To oDataSrc_Line_User.Size - 1
            oDataSrc_Line_User.SetValue("LineId", count - 1, count)
        Next
        oMatrix = oForm.Items.Item("18").Specific
        oMatrix.AutoResizeColumns()
        oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

        oMatrix = oForm.Items.Item("19").Specific
        oColumn = oMatrix.Columns.Item("V_0")
        oColumn.ChooseFromListUID = "CFL23"
        oColumn.ChooseFromListAlias = "U_Z_Code"
        oDataSrc_Line_User = oForm.DataSources.DBDataSources.Item("@Z_LUSR4")
        For count = 1 To oDataSrc_Line_User.Size - 1
            oDataSrc_Line_User.SetValue("LineId", count - 1, count)
        Next
        oMatrix = oForm.Items.Item("19").Specific
        oMatrix.AutoResizeColumns()
        oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

      
        'AddChooseFromList(oForm)
        oForm.Items.Item("9").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        oForm.PaneLevel = 1
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        '  oForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        oForm.Freeze(False)
    End Sub

    Private Sub AddChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition


            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFLCreationParams.MultiSelection = False
            oCFLCreationParams.ObjectType = "12"
            oCFLCreationParams.UniqueID = "CFL1"
            oCFLCreationParams.MultiSelection = True
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Locked"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "N"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFLCreationParams.ObjectType = "Z_ODOC"
            oCFLCreationParams.UniqueID = "CFL2"
            'oCFLCreationParams.MultiSelection = True
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_Active"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()


          


            oCFLCreationParams.ObjectType = "Z_OITC"
            oCFLCreationParams.UniqueID = "CFL21"
            oCFLCreationParams.MultiSelection = True
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_Active"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFLCreationParams.ObjectType = "Z_OBPC"
            oCFLCreationParams.UniqueID = "CFL22"
            oCFLCreationParams.MultiSelection = True
            oCFL = oCFLs.Add(oCFLCreationParams)

            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_Active"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()


            oCFLCreationParams.ObjectType = "Z_OWHC"
            oCFLCreationParams.UniqueID = "CFL23"
            oCFLCreationParams.MultiSelection = True
            oCFL = oCFLs.Add(oCFLCreationParams)
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "U_Z_Active"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

           
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub EnableControls(ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            Select Case aform.Mode
                Case SAPbouiCOM.BoFormMode.fm_ADD_MODE
                    aform.Items.Item("4").Enabled = True
                    aform.Items.Item("6").Enabled = True
                Case SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    aform.Items.Item("4").Enabled = True
                    aform.Items.Item("6").Enabled = True
            End Select
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub

#Region "AddRow /Delete Row"
    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)

        Select Case aForm.PaneLevel
            Case "1"
                oMatrix = aForm.Items.Item("7").Specific
                oDataSrc_Line_User = oForm.DataSources.DBDataSources.Item("@Z_LUSR1")
            Case "2"
                oMatrix = aForm.Items.Item("17").Specific
                oDataSrc_Line_User = oForm.DataSources.DBDataSources.Item("@Z_LUSR2")
            Case "3"
                oMatrix = aForm.Items.Item("18").Specific
                oDataSrc_Line_User = oForm.DataSources.DBDataSources.Item("@Z_LUSR3")
            Case "4"
                oMatrix = aForm.Items.Item("19").Specific
                oDataSrc_Line_User = oForm.DataSources.DBDataSources.Item("@Z_LUSR4")
        End Select
        Try
            aForm.Freeze(True)
            ' oMatrix = aForm.Items.Item("7").Specific
            If oMatrix.RowCount <= 0 Then
                oMatrix.AddRow()
            End If
            oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
            Try
                If oEditText.String <> "" Then
                    oMatrix.AddRow()
                    oMatrix.ClearRowData(oMatrix.RowCount)
                End If

            Catch ex As Exception
                aForm.Freeze(False)
                oMatrix.AddRow()
            End Try

            oMatrix.FlushToDataSource()
            ' oDataSrc_Line_User = oForm.DataSources.DBDataSources.Item("@Z_LUSR1")
            For count = 1 To oDataSrc_Line_User.Size
                oDataSrc_Line_User.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
            oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            '  oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, "")
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try

    End Sub


#End Region

    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)

            oDataSrc_Line_User = oForm.DataSources.DBDataSources.Item("@Z_LUSR1")
            If aForm.PaneLevel = 1 Then
                oDataSrc_Line_User = oForm.DataSources.DBDataSources.Item("@Z_LUSR1")
                frmSourceMatrix = aForm.Items.Item("13").Specific
            ElseIf aForm.PaneLevel = 2 Then
                frmSourceMatrix = aForm.Items.Item("17").Specific
                oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_LUSR2")
            ElseIf aForm.PaneLevel = 3 Then
                frmSourceMatrix = aForm.Items.Item("18").Specific
                oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_LUSR2")
            Else
                oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_LUSR4")
                frmSourceMatrix = aForm.Items.Item("18").Specific
            End If

            If intSelectedMatrixrow <= 0 Then
                Exit Sub
            End If
            Me.RowtoDelete = intSelectedMatrixrow
            oDataSrc_Line_User.RemoveRecord(Me.RowtoDelete - 1)
            oMatrix = frmSourceMatrix
            oMatrix.FlushToDataSource()
            For count = 1 To oDataSrc_Line_User.Size - 1
                oDataSrc_Line_User.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()
            If oMatrix.RowCount > 0 Then
                oMatrix.DeleteRow(oMatrix.RowCount)
                If aForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End If
            End If
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
        Finally
            aForm.Freeze(False)
        End Try
        

    End Sub

    Private Sub DeleteRow(ByVal aform As SAPbouiCOM.Form)
        aform.Freeze(True)

        Select Case aform.PaneLevel
            Case "1"
                oMatrix = aform.Items.Item("7").Specific
                oDataSrc_Line_User = oForm.DataSources.DBDataSources.Item("@Z_LUSR1")
            Case "2"
                oMatrix = aform.Items.Item("17").Specific
                oDataSrc_Line_User = oForm.DataSources.DBDataSources.Item("@Z_LUSR2")
            Case "3"
                oMatrix = aform.Items.Item("18").Specific
                oDataSrc_Line_User = oForm.DataSources.DBDataSources.Item("@Z_LUSR3")
            Case "4"
                oMatrix = aform.Items.Item("19").Specific
                oDataSrc_Line_User = oForm.DataSources.DBDataSources.Item("@Z_LUSR4")
        End Select
        For introw As Integer = 1 To oMatrix.RowCount
            If oMatrix.IsRowSelected(introw) Then
                oMatrix.DeleteRow(introw)
            End If
        Next
        aform.Freeze(False)
    End Sub

    Private Function Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim strsubfee, strMAfee, strWeekEndCode As String
        aForm.Freeze(True)
        If oApplication.Utilities.getEdittextvalue(oForm, "4") = "" Then
            oApplication.Utilities.Message("User Code is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
            Return False
        End If


        If oApplication.Utilities.getEdittextvalue(oForm, "6") = "" Then
            oApplication.Utilities.Message("User Name is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
            Return False
        End If


        Dim oRec As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strWeekEndCode = oApplication.Utilities.getEdittextvalue(aForm, "4")
        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            oRec.DoQuery("Select * from ""@Z_OLUSR"" where ""U_Z_UserCode""='" & strWeekEndCode & "'")
            If oRec.RecordCount > 0 Then
                oApplication.Utilities.Message("User Mapping already exists", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            End If
        End If
       
        oMatrix = aForm.Items.Item("7").Specific
        oMatrix.FlushToDataSource()
        oDataSrc_Line_User = oForm.DataSources.DBDataSources.Item("@Z_LUSR1")
        For count = 1 To oDataSrc_Line_User.Size
            oDataSrc_Line_User.SetValue("LineId", count - 1, count)
        Next
        oMatrix.LoadFromDataSource()
        oMatrix.LoadFromDataSource()

        oMatrix = aForm.Items.Item("17").Specific
        oMatrix.FlushToDataSource()
        oDataSrc_Line_User = oForm.DataSources.DBDataSources.Item("@Z_LUSR2")
        For count = 1 To oDataSrc_Line_User.Size ' - 1
            oDataSrc_Line_User.SetValue("LineId", count - 1, count)
        Next
        oMatrix.LoadFromDataSource()
        oMatrix.LoadFromDataSource()

        oMatrix = aForm.Items.Item("18").Specific
        oMatrix.FlushToDataSource()
        oDataSrc_Line_User = oForm.DataSources.DBDataSources.Item("@Z_LUSR3")
        For count = 1 To oDataSrc_Line_User.Size '- 1
            oDataSrc_Line_User.SetValue("LineId", count - 1, count)
        Next
        oMatrix.LoadFromDataSource()
        oMatrix.LoadFromDataSource()


        oMatrix = aForm.Items.Item("19").Specific
        oMatrix.FlushToDataSource()
        oDataSrc_Line_User = oForm.DataSources.DBDataSources.Item("@Z_LUSR4")
        For count = 1 To oDataSrc_Line_User.Size ' - 1
            oDataSrc_Line_User.SetValue("LineId", count - 1, count)
        Next
        oMatrix.LoadFromDataSource()
        oMatrix.LoadFromDataSource()

        aForm.Freeze(False)
        Return True
    End Function

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_UnLock Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If Validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                If (pVal.ItemUID = "7") And pVal.Row > 0 Then
                                    Me.RowtoDelete = pVal.Row
                                    intSelectedMatrixrow = pVal.Row
                                    Me.MatrixId = "7"
                                    frmSourceMatrix = oMatrix
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "10"
                                        oForm.PaneLevel = 2
                                    Case "11"
                                        oForm.PaneLevel = 3
                                    Case "12"
                                        oForm.PaneLevel = 4
                                    Case "9"
                                        oForm.PaneLevel = 1
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim val1 As String
                                Dim sCHFL_ID, val, val2, val3, val4 As String
                                Dim intChoice As Integer
                                Dim codebar As String
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        intChoice = 0
                                        oForm.Freeze(True)
                                        If oCFL.ObjectType = "Z_ODOC" Or oCFL.ObjectType = "Z_OITC" Or oCFL.ObjectType = "Z_OBPC" Or oCFL.ObjectType = "Z_OWHC" Then

                                            oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                            If Not IsNothing(oDataTable) Then
                                                Dim intAddRows As Integer = oDataTable.Rows.Count
                                                If intAddRows > 1 Then
                                                    intAddRows -= 1
                                                    oMatrix.AddRow(intAddRows, pVal.Row - 1)
                                                End If
                                                For index As Integer = 0 To oDataTable.Rows.Count - 1
                                                    Try
                                                        oApplication.Utilities.SetMatrixValues(oMatrix, pVal.ColUID, pVal.Row + index, oDataTable.GetValue("U_Z_Code", index))
                                                    Catch ex As Exception
                                                    End Try
                                                    Try
                                                        If oCFL.ObjectType <> "Z_ODOC" Then
                                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", pVal.Row + index, oDataTable.GetValue("U_Z_Name", index))
                                                        End If
                                                    Catch ex As Exception
                                                    End Try
                                                Next
                                            End If
                                            'val = oDataTable.GetValue("U_Z_Code", 0)
                                            'oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                            'Try
                                            '    oApplication.Utilities.SetMatrixValues(oMatrix, pVal.ColUID, pVal.Row, val)
                                            'Catch ex As Exception
                                            'End Try
                                            'Try
                                            '    If oCFL.ObjectType <> "Z_ODOC" Then
                                            '        oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", pVal.Row, oDataTable.GetValue("U_Z_Name", 0))
                                            '    End If
                                            'Catch ex As Exception
                                            'End Try
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            End If
                                            oForm.Freeze(False)
                                        End If

                                        If oCFL.ObjectType = "12" Then
                                            val = oDataTable.GetValue("USER_CODE", 0)
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                            Catch ex As Exception

                                            End Try
                                            oApplication.Utilities.setEdittextvalue(oForm, "6", oDataTable.GetValue("U_NAME", 0))
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            End If
                                            oForm.Freeze(False)
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
                Case mnu_UnLock
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If
                Case mnu_ADD_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    AddRow(oForm)
                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        RefereshDeleteRow(oForm)
                    Else

                    End If
                Case mnu_ADD
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        oForm.Items.Item("4").Enabled = True
                        oForm.Items.Item("6").Enabled = True
                    End If
                Case mnu_FIND
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        oForm.Items.Item("4").Enabled = True
                        oForm.Items.Item("6").Enabled = True
                    End If
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                If BusinessObjectInfo.BeforeAction = False Then
                    oForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    oForm.Items.Item("4").Enabled = False
                End If
            ElseIf BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                If BusinessObjectInfo.FormTypeEx = frm_UnLock Then
                    Dim s As String = oApplication.Company.GetNewObjectType()
                    Dim stXML As String = BusinessObjectInfo.ObjectKey
                    stXML = stXML.Replace("<?xml version=""1.0"" encoding=""UTF-16"" ?><Unlock_PostingDateParams><DocEntry>", "")
                    stXML = stXML.Replace("</DocEntry></Unlock_PostingDateParams>", "")
                    stXML = stXML.Replace("<?xml version=""1.0"" encoding=""UTF-16"" ?><User_MappingParams><DocEntry>", "")
                    stXML = stXML.Replace("</DocEntry></User_MappingParams>", "")
                    Dim otest, otest1 As SAPbobsCOM.Recordset
                    otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    If stXML <> "" Then
                        otest.DoQuery("select * from ""@Z_OLUSR""  where ""DocEntry""=" & stXML)
                        If otest.RecordCount > 0 Then
                            Dim strUserCode As String = otest.Fields.Item("U_Z_UserCode").Value
                            oApplication.Utilities.UpdateCategores()
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
        'If eventInfo.FormUID = "RightClk" Then
        If 1 = 1 Then 'oForm.TypeEx = frm_HR_Trainner Then
            If (eventInfo.BeforeAction = True) Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        'Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                        'oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                        'oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                        'oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                        'oCreationPackage.UniqueID = "TraDetails"
                        'oCreationPackage.String = "Trainning Details"
                        'oCreationPackage.Enabled = True
                        'oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                        'oMenus = oMenuItem.SubMenus
                        'oMenus.AddEx(oCreationPackage)
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
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            End If
        End If
    End Sub
End Class
