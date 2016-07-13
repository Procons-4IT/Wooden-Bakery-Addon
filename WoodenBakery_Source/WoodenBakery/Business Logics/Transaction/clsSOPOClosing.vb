Imports System.IO
Public Class clsSOPOClosing
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
    Private blnErrorflag As Boolean = False

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Private Sub LoadForm()
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_SOClosing) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_SOClosing, frm_SOClosing)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        AddChooseFromList(oForm)
        oForm.DataSources.UserDataSources.Add("DocType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oCombobox = oForm.Items.Item("7").Specific
        oCombobox.DataBind.SetBound(True, "", "DocType")
        oCombobox.ValidValues.Add("SO", "Sales Order")
        oCombobox.ValidValues.Add("PO", "Purchase Order")
        oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
        FillCombo(oForm)
        oForm.PaneLevel = 1
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
            'oCFLCreationParams.ObjectType = "Z_RECO"
            'oCFLCreationParams.UniqueID = "CFL12"
            'oCFL = oCFLs.Add(oCFLCreationParams)


            'oCFLCreationParams.ObjectType = "Z_HR_OPEOB"
            'oCFLCreationParams.UniqueID = "CFL2"
            'oCFL = oCFLs.Add(oCFLCreationParams)


            'oCFLCreationParams.ObjectType = "Z_HR_OCOMP"
            'oCFLCreationParams.UniqueID = "CFL3"
            'oCFL = oCFLs.Add(oCFLCreationParams)

            'oCFLCreationParams.ObjectType = "Z_HR_OPOSIN"
            'oCFLCreationParams.UniqueID = "CFL4"
            'oCFL = oCFLs.Add(oCFLCreationParams)

            'oCFLCreationParams.ObjectType = "Z_HR_OCOCA"
            'oCFLCreationParams.UniqueID = "CFL5"
            oCFL = oCFLs.Item("CFL_6")

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

#Region "Data Bind"
    Private Sub DataBind(aForm As SAPbouiCOM.Form)
        Dim strItemGroup, strCust_Route, strDocNo, strDocNo1, strDocType, strReasonCode, strDocDate, strDocDate1, strCardcode, strCardCode1, strItemCode, strItemCode1, strCondition, strSQL As String
        Dim dtDate, dtDate1 As Date
        Try
            aForm.Freeze(True)
            oCombobox = aForm.Items.Item("7").Specific
            strDocType = oCombobox.Selected.Value
            strReasonCode = oApplication.Utilities.getEdittextvalue(aForm, "9")
            strDocDate = oApplication.Utilities.getEdittextvalue(aForm, "11")
            strDocDate1 = oApplication.Utilities.getEdittextvalue(aForm, "13")
            strCardcode = oApplication.Utilities.getEdittextvalue(aForm, "15")
            strCardCode1 = oApplication.Utilities.getEdittextvalue(aForm, "17")
            strItemCode = oApplication.Utilities.getEdittextvalue(aForm, "19")
            strItemCode1 = oApplication.Utilities.getEdittextvalue(aForm, "21")
            strItemGroup = CType(oForm.Items.Item("29").Specific, SAPbouiCOM.ComboBox).Selected.Value
            strDocNo = oApplication.Utilities.getEdittextvalue(aForm, "30")
            strDocNo1 = oApplication.Utilities.getEdittextvalue(aForm, "30_")
            strCust_Route = oApplication.Utilities.getEdittextvalue(aForm, "31")

            If strDocType = "SO" Then
                strSQL = "SELECT 'Y' ""Select"", T1.""DocEntry"", T0.""DocNum"", T0.""DocDate"", T0.""CardCode"", T0.""CardName"", T1.""VisOrder"" ""LineNum"", T1.""ItemCode"", T1.""Dscription"", T1.""Quantity"",T1.""OpenCreQty""  ""OpenQty"", T1.""U_Z_RECODE"",T1.""OpenQty"" As ""OpenQty1"" FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" INNER JOIN OITM T2 ON T1.""ItemCode"" = T2.""ItemCode"" INNER JOIN OITB T3 ON T3.""ItmsGrpCod"" = T2.""ItmsGrpCod"" INNER JOIN OCRD T4 ON T0.""CardCode"" = T4.""CardCode"" Left Outer  JOIN ""@CUSTOMERROUTE"" T5 ON T5.""Code"" = T4.""U_Cust_Route"" "
            Else
                strSQL = "SELECT 'Y' ""Select"", T1.""DocEntry"", T0.""DocNum"", T0.""DocDate"", T0.""CardCode"", T0.""CardName"", T1.""VisOrder"" ""LineNum"", T1.""ItemCode"", T1.""Dscription"", T1.""Quantity"",T1.""OpenCreQty""  ""OpenQty"", T1.""U_Z_RECODE"",T1.""OpenQty"" As ""OpenQty1"" FROM OPOR T0  INNER JOIN POR1 T1 ON T0.""DocEntry"" = T1.""DocEntry"" INNER JOIN OITM T2 ON T1.""ItemCode"" = T2.""ItemCode"" INNER JOIN OITB T3 ON T3.""ItmsGrpCod"" = T2.""ItmsGrpCod"" INNER JOIN OCRD T4 ON T0.""CardCode"" = T4.""CardCode"" Left Outer  JOIN ""@CUSTOMERROUTE"" T5 ON T5.""Code"" = T4.""U_Cust_Route"" "
            End If

            'If strReasonCode <> "" Then
            '    strCondition = "T1.""U_Z_RECCODE""='" & strReasonCode & "'"
            'Else
            '    strCondition = "1=1"
            'End If

            strCondition = "1=1"

            If strCardcode <> "" Then
                strCondition = strCondition & " and ( T0.""CardCode"">='" & strCardcode & "'"
            Else
                strCondition = strCondition & " and ( 1=1 "
            End If
            If strCardCode1 <> "" Then
                strCondition = strCondition & " and T0.""CardCode""<='" & strCardCode1 & "')"
            Else
                strCondition = strCondition & " and 1=1) "
            End If


            If strItemCode <> "" Then
                strCondition = strCondition & " and ( T1.""ItemCode"">='" & strItemCode & "'"
            Else
                strCondition = strCondition & " and ( 1=1 "
            End If
            If strItemCode1 <> "" Then
                strCondition = strCondition & " and T1.""ItemCode""<='" & strItemCode1 & "')"
            Else
                strCondition = strCondition & " and 1=1) "
            End If

            If strDocDate <> "" Then
                dtDate = oApplication.Utilities.GetDateTimeValue(strDocDate)
                strCondition = strCondition & " and ( T0.""DocDate"">='" & dtDate.ToString("yyyy-MM-dd") & "'"
            Else
                strCondition = strCondition & " and ( 1=1 "
            End If
            If strDocDate1 <> "" Then
                dtDate1 = oApplication.Utilities.GetDateTimeValue(strDocDate1)
                strCondition = strCondition & " and T0.""DocDate""<='" & dtDate1.ToString("yyyy-MM-dd") & "')"
            Else
                strCondition = strCondition & " and 1=1) "
            End If

            If strDocNo <> "" Then
                strCondition = strCondition & " and ( T0.""DocNum"">='" & strDocNo & "'"
            Else
                strCondition = strCondition & " and ( 1=1 "
            End If
            If strDocNo1 <> "" Then
                strCondition = strCondition & " and T0.""DocNum""<='" & strDocNo1 & "')"
            Else
                strCondition = strCondition & " and 1=1) "
            End If


            If strItemGroup <> "" Then
                strCondition = strCondition & " and T2.""ItmsGrpCod""= '" & strItemGroup & "'"
            Else
                strCondition = strCondition & " and 1=1  "
            End If


            If strCust_Route <> "" Then
                strCondition = strCondition & " and T5.""Code""= '" & strCust_Route & "'"
            Else
                strCondition = strCondition & " and 1=1  "
            End If

            strSQL = strSQL & " Where  T1.""LineStatus""='O' and " & strCondition
            oGrid = aForm.Items.Item("22").Specific
            oGrid.DataTable.ExecuteQuery(strSQL)
            oGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
            For intR As Integer = 0 To oGrid.Columns.Count - 1

                If intR > 0 Then
                    oGrid.Columns.Item(intR).Editable = False
                End If

                If oGrid.Columns.Item(intR).UniqueID = "OpenQty" Then
                    oGrid.Columns.Item(intR).Editable = True
                End If

            Next
            oGrid.Columns.Item("U_Z_RECODE").TitleObject.Caption = "Closing Reason Code"
            oGrid.Columns.Item("U_Z_RECODE").Visible = False
            oGrid.Columns.Item("OpenQty1").Visible = False

            oEditTextColumn = oGrid.Columns.Item("DocEntry")
            If strDocType = "SO" Then
                oEditTextColumn.LinkedObjectType = "17"
            Else
                oEditTextColumn.LinkedObjectType = "22"
            End If
            oEditTextColumn = oGrid.Columns.Item("CardCode")
            oEditTextColumn.LinkedObjectType = "2"
            oEditTextColumn = oGrid.Columns.Item("ItemCode")
            oEditTextColumn.LinkedObjectType = "4"
            oGrid.AutoResizeColumns()
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub
#End Region

    Private Sub Select1(aform As SAPbouiCOM.Form, aflag As Boolean)
        aform.Freeze(True)
        oGrid = aform.Items.Item("22").Specific
        Dim ocheck As SAPbouiCOM.CheckBoxColumn
        ocheck = oGrid.Columns.Item("Select")
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            ocheck.Check(intRow, aflag)
        Next
        aform.Freeze(False)
    End Sub

    Private Function CloseDOcumentLines(aform As SAPbouiCOM.Form, aFileName As String) As Boolean
        oCombobox = aform.Items.Item("7").Specific
        Dim oDoc As SAPbobsCOM.Documents
        Dim oCheckbox As SAPbouiCOM.CheckBoxColumn
        If oCombobox.Selected.Value = "SO" Then
            oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
        Else
            oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
        End If
        Dim strReasoncode As String = oApplication.Utilities.getEdittextvalue(aform, "9")
        Dim oRec As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oGrid = aform.Items.Item("22").Specific
        oCheckbox = oGrid.Columns.Item("Select")
        Dim strDocType As String = oCombobox.Selected.Value
        If strDocType = "SO" Then
            strDocType = "Sales Order"
        Else
            strDocType = "Purchase Order"
        End If

        ' oApplication.Utilities.Trace_Process("Started Creating Reserve Invoice : " + System.DateTime.Now, strFile)
        Dim s As System.Text.StringBuilder = New System.Text.StringBuilder
        Dim strRecquery, strdocnum As String
        '  strRecquery = GetSalesOrders()
        Dim otemprec As SAPbobsCOM.Recordset
        otemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemprec.DoQuery(strSQL)
        s.Remove(0, s.Length)
        s.Append("DocType" + vbTab)
        s.Append("DocumentNumber" + vbTab)
        s.Append("CustomerCode" + vbTab)
        s.Append("CustomerName" + vbTab)
        s.Append("LineNum" + vbTab)
        s.Append("ItemCode" + vbTab)
        s.Append("ItemDesc" + vbTab)
        s.Append("Quantity" + vbTab)
        s.Append("Remakrs" + vbCrLf)
        blnErrorflag = False
        oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            If oCheckbox.IsChecked(intRow) Then
                If oDoc.GetByKey(oGrid.DataTable.GetValue("DocEntry", intRow)) Then
                    oDoc.Lines.SetCurrentLine(oGrid.DataTable.GetValue("LineNum", intRow))
                    If oDoc.Lines.LineStatus <> SAPbobsCOM.BoStatus.bost_Close Then
                        Dim dblOpenQty As Double = CDbl(oGrid.DataTable.GetValue("OpenQty", intRow))
                        Dim dblOpenQty1 As Double = CDbl(oGrid.DataTable.GetValue("OpenQty1", intRow))
                        If dblOpenQty = dblOpenQty1 Then
                            oDoc.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Close
                        Else
                            oDoc.Lines.Quantity = dblOpenQty
                        End If
                        '        oDoc.Lines.UserFields.Fields.Item("U_Z_RECODE").Value = strReasoncode
                    End If
                    If oDoc.Update <> 0 Then
                        '    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    Return False
                        s.Append(strDocType + vbTab)
                        s.Append(oGrid.DataTable.GetValue("DocEntry", intRow).ToString + vbTab)
                        s.Append(oGrid.DataTable.GetValue("DocNum", intRow).ToString + vbTab)
                        s.Append(oGrid.DataTable.GetValue("CardCode", intRow).ToString + vbTab)
                        s.Append(oGrid.DataTable.GetValue("CardName", intRow).ToString + vbTab)
                        s.Append(oGrid.DataTable.GetValue("LineNum", intRow).ToString + vbTab)
                        s.Append(oGrid.DataTable.GetValue("ItemCode", intRow).ToString + vbTab)
                        s.Append(oGrid.DataTable.GetValue("Dscription", intRow).ToString + vbTab)
                        s.Append(oApplication.Company.GetLastErrorDescription.ToCharArray + vbTab)
                        s.Append(vbCrLf)
                        blnErrorflag = True
                    Else
                        If oCombobox.Selected.Value = "SO" Then
                            oRec.DoQuery("Update  RDR1 set ""U_Z_RECODE""='" & strReasoncode & "' where ""DocEntry""=" & oGrid.DataTable.GetValue("DocEntry", intRow) & " and ""VisOrder""=" & oGrid.DataTable.GetValue("LineNum", intRow))
                        Else
                            oRec.DoQuery("Update  POR1 set ""U_Z_RECODE""='" & strReasoncode & "' where ""DocEntry""=" & oGrid.DataTable.GetValue("DocEntry", intRow) & " and ""VisOrder""=" & oGrid.DataTable.GetValue("LineNum", intRow))
                        End If
                        s.Append(strDocType + vbTab)
                        s.Append(oGrid.DataTable.GetValue("DocEntry", intRow).ToString + vbTab)
                        s.Append(oGrid.DataTable.GetValue("DocNum", intRow).ToString + vbTab)
                        s.Append(oGrid.DataTable.GetValue("CardCode", intRow).ToString + vbTab)
                        s.Append(oGrid.DataTable.GetValue("CardName", intRow).ToString + vbTab)
                        s.Append(oGrid.DataTable.GetValue("LineNum", intRow).ToString + vbTab)
                        s.Append(oGrid.DataTable.GetValue("ItemCode", intRow).ToString + vbTab)
                        s.Append(oGrid.DataTable.GetValue("Dscription", intRow).ToString + vbTab)
                        s.Append("Closed Successfully" + vbTab)
                        s.Append(vbCrLf)
                    End If
                End If
            End If
        Next
        Dim strPath As String = System.Windows.Forms.Application.StartupPath.ToString() + aFileName
        aFileName = strPath 'strPath & "\" & strFile
        My.Computer.FileSystem.WriteAllText(aFileName, s.ToString, False)
       
        Return True
    End Function

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_SOClosing Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim sCHFL_ID As String
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    filterChooseFromList(oForm, sCHFL_ID)
                                Catch ex As Exception

                                End Try
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "7" Then
                                    oCombobox = oForm.Items.Item("7").Specific
                                    If oCombobox.Selected.Value = "SO" Then
                                        CType(oForm.Items.Item("30").Specific, SAPbouiCOM.EditText).ChooseFromListUID = "CFL_7"
                                        oEditText = oForm.Items.Item("30").Specific
                                        oEditText.ChooseFromListAlias = "DocNum"
                                        CType(oForm.Items.Item("30_").Specific, SAPbouiCOM.EditText).ChooseFromListUID = "CFL_7_0"
                                        oEditText = oForm.Items.Item("30_").Specific
                                        oEditText.ChooseFromListAlias = "DocNum"
                                    Else
                                        CType(oForm.Items.Item("30").Specific, SAPbouiCOM.EditText).ChooseFromListUID = "CFL_8"

                                        oEditText = oForm.Items.Item("30").Specific
                                        oEditText.ChooseFromListAlias = "DocNum"
                                        CType(oForm.Items.Item("30_").Specific, SAPbouiCOM.EditText).ChooseFromListUID = "CFL_8_0"
                                        oEditText = oForm.Items.Item("30_").Specific
                                        oEditText.ChooseFromListAlias = "DocNum"
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "3"
                                        oForm.PaneLevel = oForm.PaneLevel - 1
                                    Case "4"
                                        If oForm.PaneLevel = 2 Then
                                            If oApplication.Utilities.getEdittextvalue(oForm, "9") = "" Then
                                                oApplication.Utilities.Message("Reason Code is missing..", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                Exit Sub
                                            End If
                                        End If
                                        oForm.PaneLevel = oForm.PaneLevel + 1
                                        If oForm.PaneLevel = 3 Then
                                            DataBind(oForm)
                                        End If
                                    Case "23"
                                        Select1(oForm, True)
                                    Case "24"
                                        Select1(oForm, False)
                                    Case "5"

                                        If oApplication.SBO_Application.MessageBox("Do you want to close the selected document lines ?", , "Continue", "Cancel") = 2 Then
                                            Exit Sub
                                        End If
                                        'If oApplication.Company.InTransaction() Then
                                        '    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                        'End If
                                        '       oApplication.Company.StartTransaction()
                                        Dim strFile As String = "\Log\Closing_Documents" + System.DateTime.Now.ToString("yyyyMMddmmss") + ".txt"
                                        'If File.Exists(strFile) Then
                                        'Else
                                        '    File.Create(strFile)
                                        'End If
                                        CloseDOcumentLines(oForm, strFile)
                                        If blnErrorflag = False Then
                                            'If oApplication.Company.InTransaction() Then
                                            '    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                            'End If
                                            oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                        Else
                                            If oApplication.SBO_Application.MessageBox("Some of the lines are not closed . Check the Log file for the reason", , "OK") = 1 Then
                                                'If oApplication.Company.InTransaction() Then
                                                '    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                                'End If
                                            Else
                                                'If oApplication.Company.InTransaction() Then
                                                '    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                                                'End If
                                            End If
                                        End If
                                        oForm.Close()
                                        Dim strPath As String = System.Windows.Forms.Application.StartupPath.ToString() + strFile
                                        If (File.Exists(strPath)) Then
                                            System.Diagnostics.Process.Start(strPath)
                                        End If
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

                                        If oCFL.ObjectType = "4" Then
                                            val = oDataTable.GetValue("ItemCode", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                        End If
                                        If oCFL.ObjectType = "2" And pVal.ItemUID <> "31" Then
                                            val = oDataTable.GetValue("CardCode", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                        End If
                                        If oCFL.ObjectType = "Z_RECO" Then
                                            val = oDataTable.GetValue("U_Z_Code", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                        End If
                                        If oCFL.ObjectType = "17" Then
                                            val = oDataTable.GetValue("DocNum", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                        End If
                                        If oCFL.ObjectType = "22" Then
                                            val = oDataTable.GetValue("DocNum", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                        End If
                                        If pVal.ItemUID = "31" Then
                                            val = oDataTable.GetValue("Code", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
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
                Case mnu_SOClosing
                    If pVal.MenuUID = mnu_SOClosing Then
                        LoadForm()
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
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub filterChooseFromList(ByVal oForm As SAPbouiCOM.Form, ByVal strCFLID As String)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim strType As String = CType(oForm.Items.Item("7").Specific, SAPbouiCOM.ComboBox).Selected.Value
            If strCFLID <> "" Then
                oCFLs = oForm.ChooseFromLists
                oCFL = oCFLs.Item(strCFLID)
                If strCFLID = "CFL_6" Then
                    oCons = oCFL.GetConditions()
                    If oCons.Count = 0 Then
                        oCon = oCons.Add()
                    Else
                        oCon = oCons.Item(0)
                    End If
                    oCon.Alias = "U_Z_Type"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = strType
                    oCFL.SetConditions(oCons)
                End If

                If (strCFLID = "CFL_2" Or strCFLID = "CFL_3") And strType = "SO" Then
                    oCons = oCFL.GetConditions()
                    If oCons.Count = 0 Then
                        oCon = oCons.Add()
                    Else
                        oCon = oCons.Item(0)
                    End If
                    oCon.Alias = "CardType"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "C"
                    oCFL.SetConditions(oCons)
                ElseIf (strCFLID = "CFL_2" Or strCFLID = "CFL_3") And strType = "PO" Then
                    oCons = oCFL.GetConditions()
                    If oCons.Count = 0 Then
                        oCon = oCons.Add()
                    Else
                        oCon = oCons.Item(0)
                    End If
                    oCon.Alias = "CardType"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "S"
                    oCFL.SetConditions(oCons)
                End If

                If (strCFLID = "CFL_4" Or strCFLID = "CFL_5") And strType = "SO" Then
                    oCons = oCFL.GetConditions()
                    If oCons.Count = 0 Then
                        oCon = oCons.Add()
                    Else
                        oCon = oCons.Item(0)
                    End If
                    oCon.Alias = "SellItem"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                    oCFL.SetConditions(oCons)
                ElseIf (strCFLID = "CFL_4" Or strCFLID = "CFL_5") And strType = "PO" Then
                    oCons = oCFL.GetConditions()
                    If oCons.Count = 0 Then
                        oCon = oCons.Add()
                    Else
                        oCon = oCons.Item(0)
                    End If
                    oCon.Alias = "PrchseItem"
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCon.CondVal = "Y"
                    oCFL.SetConditions(oCons)
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub FillCombo(ByVal aForm As SAPbouiCOM.Form)
        Try
            Dim oTempRec As SAPbobsCOM.Recordset
            oTempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


            oCombobox = aForm.Items.Item("29").Specific
            For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
                oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
            oCombobox.ValidValues.Add("", "")
            oTempRec.DoQuery("Select ""ItmsGrpCod"",""ItmsGrpNam"" From OITB")
            For intRow As Integer = 0 To oTempRec.RecordCount - 1
                oCombobox.ValidValues.Add(oTempRec.Fields.Item("ItmsGrpCod").Value, oTempRec.Fields.Item("ItmsGrpNam").Value)
                oTempRec.MoveNext()
            Next

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

End Class
