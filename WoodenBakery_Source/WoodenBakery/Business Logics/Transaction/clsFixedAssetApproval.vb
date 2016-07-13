Public Class clsFixedAssetApproval
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oFolder, oFolder1 As SAPbouiCOM.Folder
    Private ocombo As SAPbouiCOM.ComboBoxColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo, sQuery As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Public Sub LoadForm(ByVal oForm As SAPbouiCOM.Form)
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_FATransactionApp) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_FATransactionApp, frm_FATransactionApp)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.PaneLevel = 1
        oForm.DataSources.DataTables.Add("dtDocumentList")
        oForm.DataSources.DataTables.Add("dtHistoryList")
        InitializationApproval(oForm, HeaderDoctype.Fix)
        ApprovalSummary(oForm, HeaderDoctype.Fix)
        oGrid = oForm.Items.Item("1").Specific
        oGrid.Columns.Item("RowsHeader").Click(0, False, False)
        oGrid = oForm.Items.Item("19").Specific
        oGrid.Columns.Item("RowsHeader").Click(0, False, False)
        oForm.Items.Item("4").TextStyle = 7
        oForm.Items.Item("5").TextStyle = 7
        oForm.Freeze(False)
    End Sub

#Region "Approval Functions"
    Public Sub InitializationApproval(ByVal aForm As SAPbouiCOM.Form, ByVal enDocType As modVariables.HeaderDoctype)
        aForm.Freeze(True)
        Dim oTempDt As SAPbouiCOM.DataTable
        Dim oGrid As SAPbouiCOM.Grid
        oGrid = aForm.Items.Item("1").Specific
        Select Case enDocType
            Case HeaderDoctype.Fix
                ''sQuery = " Select T0.""DocEntry"",T0.""U_Z_Code"",T0.""U_Z_Name"",CASE ""U_Z_TransType"" when 'L' then ""Location"" when 'E' then ""Employee Transfer"" else ""CostCenter"" end AS ""U_Z_TransType"","
                'sQuery = "SELECT T0.""DocEntry"", T0.""U_Z_Code"", T0.""U_Z_Name"", CASE ""U_Z_TransType""  WHEN 'L' THEN 'Location'  WHEN 'E' THEN 'Employee Transfer'  ELSE 'CostCenter' END AS ""U_Z_TransType"","
                'sQuery += " ""U_Z_FName"", ""U_Z_TName"",CASE ""U_Z_DocStatus"" when 'D' then ""Draft"" when 'N' then ""Confirm"" when 'P' then ""Pending for Approval"" when 'A' then ""Approved"" when 'R' then ""Rejected"" when 'C' then ""Cancel"" else ""Close"" end AS ""U_Z_DocStatus"","
                'sQuery += " ""U_Z_Attachment"", ""U_Z_Remarks"", ""U_Z_AppStatus"", ""U_Z_CurrApprover"" ""Current Approver"",""U_Z_NextApprover"" ""Next Approver"",T0.""Creator"", "
                'sQuery += " Case ""U_Z_IsApp"" when ""Y"" then ""Yes"" else ""No"" End as  ""Approval Required"",""U_Z_AppReqDate"" ""Requested Date"",T0.""U_Z_ApproveId"""
                'sQuery += " From ""@Z_OFATA"" T0 JOIN ""@Z_OAPPT"" T3 ON T3.""DocEntry"" = T0.""U_Z_ApproveId"" and (T0.""U_Z_AppStatus""='P' or T0.""U_Z_AppStatus""='-') "
                'sQuery += " JOIN ""@Z_APPT2"" T2 ON T3.""DocEntry"" = T2.""DocEntry"" "

                If blnIsHanaDB Then
                    sQuery = "SELECT T0.""DocEntry"", T0.""U_Z_Code"", T0.""U_Z_Name"", CASE ""U_Z_TransType""  WHEN 'L' THEN 'Location'  WHEN 'E' THEN 'Employee Transfer'  ELSE 'CostCenter' END AS ""U_Z_TransType"", ""U_Z_FName"", ""U_Z_TName"", CASE ""U_Z_DocStatus"" WHEN 'D' THEN 'Draft' WHEN 'N' THEN 'Confirm' WHEN 'P' THEN 'Pending for Approval'  WHEN 'A' THEN 'Approved' WHEN 'R' THEN 'Rejected' WHEN 'C' THEN 'Cancel' ELSE 'Close' END AS ""U_Z_DocStatus"", ""U_Z_Attachment"", ""U_Z_Remarks"", ""U_Z_AppStatus"", ""U_Z_CurrApprover"" AS ""Current Approver"", ""U_Z_NextApprover"" AS ""Next Approver"", T0.""Creator"", CASE ""U_Z_IsApp"" WHEN 'Y' THEN 'Yes'  ELSE 'No' END AS ""Approval Required"", ""U_Z_AppReqDate"" AS ""Requested Date"", T0.""U_Z_ApproveId"" FROM ""@Z_OFATA"" T0    INNER JOIN ""@Z_OAPPT"" T3 ON T3.""DocEntry"" = T0.""U_Z_ApproveId"" AND (T0.""U_Z_AppStatus"" = 'P' OR T0.""U_Z_AppStatus"" = '-') INNER JOIN ""@Z_APPT2"" T2 ON T3.""DocEntry"" = T2.""DocEntry"""
                    sQuery += " And (T0.""U_Z_CurrApprover"" = '" + oApplication.Company.UserName + "' OR T0.""U_Z_NextApprover"" = '" + oApplication.Company.UserName + "')"
                    sQuery += " And IFNULL(T2.""U_Z_AMan"",'N')='Y' AND IFNULL(T3.""U_Z_Active"",'N')='Y' and  IFNULL(T0.""U_Z_IsApp"",'N')='Y' and  T2.""U_Z_AUser"" = '" + oApplication.Company.UserName + "' And T3.""U_Z_DocType"" = '" + enDocType.ToString() + "' Order by T0.""DocEntry"" desc "
                Else
                    sQuery = "SELECT T0.""DocEntry"", T0.""U_Z_Code"", T0.""U_Z_Name"", CASE ""U_Z_TransType""  WHEN 'L' THEN 'Location'  WHEN 'E' THEN 'Employee Transfer'  ELSE 'CostCenter' END AS ""U_Z_TransType"", ""U_Z_FName"", ""U_Z_TName"", CASE ""U_Z_DocStatus"" WHEN 'D' THEN 'Draft' WHEN 'N' THEN 'Confirm' WHEN 'P' THEN 'Pending for Approval'  WHEN 'A' THEN 'Approved' WHEN 'R' THEN 'Rejected' WHEN 'C' THEN 'Cancel' ELSE 'Close' END AS ""U_Z_DocStatus"", ""U_Z_Attachment"", ""U_Z_Remarks"", ""U_Z_AppStatus"", ""U_Z_CurrApprover"" AS ""Current Approver"", ""U_Z_NextApprover"" AS ""Next Approver"", T0.""Creator"", CASE ""U_Z_IsApp"" WHEN 'Y' THEN 'Yes'  ELSE 'No' END AS ""Approval Required"", ""U_Z_AppReqDate"" AS ""Requested Date"", T0.""U_Z_ApproveId"" FROM ""@Z_OFATA"" T0    INNER JOIN ""@Z_OAPPT"" T3 ON T3.""DocEntry"" = T0.""U_Z_ApproveId"" AND (T0.""U_Z_AppStatus"" = 'P' OR T0.""U_Z_AppStatus"" = '-') INNER JOIN ""@Z_APPT2"" T2 ON T3.""DocEntry"" = T2.""DocEntry"""
                    sQuery += " And (T0.""U_Z_CurrApprover"" = '" + oApplication.Company.UserName + "' OR T0.""U_Z_NextApprover"" = '" + oApplication.Company.UserName + "')"
                    sQuery += " And ISNULL(T2.""U_Z_AMan"",'N')='Y' AND ISNULL(T3.""U_Z_Active"",'N')='Y' and  ISNULL(T0.""U_Z_IsApp"",'N')='Y' and  T2.""U_Z_AUser"" = '" + oApplication.Company.UserName + "' And T3.""U_Z_DocType"" = '" + enDocType.ToString() + "' Order by T0.""DocEntry"" desc "
                End If
        End Select
        oTempDt = aForm.DataSources.DataTables.Item("dtDocumentList")
        oTempDt.ExecuteQuery(sQuery)
        oGrid.DataTable.ExecuteQuery(sQuery)
        formatDocument(aForm, enDocType)
        oApplication.Utilities.assignMatrixLineno(oGrid, aForm)
        oGrid.Columns.Item("RowsHeader").Click(0, False, False)
        aForm.Freeze(False)
    End Sub
    Public Sub ApprovalSummary(ByVal aForm As SAPbouiCOM.Form, ByVal enDocType As modVariables.HeaderDoctype)
        aForm.Freeze(True)
        Dim oTempDt As SAPbouiCOM.DataTable
        Dim oGrid As SAPbouiCOM.Grid
        oGrid = aForm.Items.Item("19").Specific
        Select Case enDocType
            Case HeaderDoctype.Fix

                sQuery = " Select T0.""DocEntry"",T0.""U_Z_Code"",T0.""U_Z_Name"",CASE ""U_Z_TransType"" when 'L' then ""Location"" when 'E' then ""Employee Transfer"" else ""CostCenter"" end AS ""U_Z_TransType"","
                sQuery += " ""U_Z_FName"", ""U_Z_TName"",CASE ""U_Z_DocStatus"" when 'D' then ""Draft"" when 'N' then ""Confirm"" when 'P' then ""Pending for Approval"" when 'A' then ""Approved"" when 'R' then ""Rejected"" when 'C' then ""Cancel"" else ""Close"" end AS ""U_Z_DocStatus"","
                sQuery += " ""U_Z_Attachment"", ""U_Z_Remarks"", ""U_Z_AppStatus"", ""U_Z_CurrApprover"" ""Current Approver"",""U_Z_NextApprover"" ""Next Approver"",T0.""Creator"", "
                sQuery += " Case ""U_Z_IsApp"" when ""Y"" then ""Yes"" else ""No"" End as  ""Approval Required"",""U_Z_AppReqDate"" ""Requested Date"",T0.""U_Z_ApproveId"""
                sQuery += " From ""@Z_OFATA"" T0 JOIN ""@Z_OAPPT"" T3 ON T3.""DocEntry"" = T0.""U_Z_ApproveId""" ' and (T0.""U_Z_AppStatus""='P' or T0.""U_Z_AppStatus""='-') "
                sQuery += " JOIN ""@Z_APPT2"" T2 ON T3.""DocEntry"" = T2.""DocEntry"" "

                If blnIsHanaDB Then
                    sQuery = "SELECT T0.""DocEntry"", T0.""U_Z_Code"", T0.""U_Z_Name"", CASE ""U_Z_TransType""  WHEN 'L' THEN 'Location'  WHEN 'E' THEN 'Employee Transfer'  ELSE 'CostCenter' END AS ""U_Z_TransType"", ""U_Z_FName"", ""U_Z_TName"", CASE ""U_Z_DocStatus"" WHEN 'D' THEN 'Draft' WHEN 'N' THEN 'Confirm' WHEN 'P' THEN 'Pending for Approval'  WHEN 'A' THEN 'Approved' WHEN 'R' THEN 'Rejected' WHEN 'C' THEN 'Cancel' ELSE 'Close' END AS ""U_Z_DocStatus"", ""U_Z_Attachment"", ""U_Z_Remarks"", ""U_Z_AppStatus"", ""U_Z_CurrApprover"" AS ""Current Approver"", ""U_Z_NextApprover"" AS ""Next Approver"", T0.""Creator"", CASE ""U_Z_IsApp"" WHEN 'Y' THEN 'Yes'  ELSE 'No' END AS ""Approval Required"", ""U_Z_AppReqDate"" AS ""Requested Date"", T0.""U_Z_ApproveId"" FROM ""@Z_OFATA"" T0    INNER JOIN ""@Z_OAPPT"" T3 ON T3.""DocEntry"" = T0.""U_Z_ApproveId"" INNER JOIN ""@Z_APPT2"" T2 ON T3.""DocEntry"" = T2.""DocEntry"""
                    sQuery += " And (T0.""U_Z_CurrApprover"" = '" + oApplication.Company.UserName + "' OR T0.""U_Z_NextApprover"" = '" + oApplication.Company.UserName + "')"
                    sQuery += " And IFNULL(T2.""U_Z_AMan"",'N')='Y' AND IFNULL(T3.""U_Z_Active"",'N')='Y' and  IFNULL(T0.""U_Z_IsApp"",'N')='Y' and  T2.""U_Z_AUser"" = '" + oApplication.Company.UserName + "' And T3.""U_Z_DocType"" = '" + enDocType.ToString() + "' Order by T0.""DocEntry"" desc "
                Else
                    sQuery = "SELECT T0.""DocEntry"", T0.""U_Z_Code"", T0.""U_Z_Name"", CASE ""U_Z_TransType""  WHEN 'L' THEN 'Location'  WHEN 'E' THEN 'Employee Transfer'  ELSE 'CostCenter' END AS ""U_Z_TransType"", ""U_Z_FName"", ""U_Z_TName"", CASE ""U_Z_DocStatus"" WHEN 'D' THEN 'Draft' WHEN 'N' THEN 'Confirm' WHEN 'P' THEN 'Pending for Approval'  WHEN 'A' THEN 'Approved' WHEN 'R' THEN 'Rejected' WHEN 'C' THEN 'Cancel' ELSE 'Close' END AS ""U_Z_DocStatus"", ""U_Z_Attachment"", ""U_Z_Remarks"", ""U_Z_AppStatus"", ""U_Z_CurrApprover"" AS ""Current Approver"", ""U_Z_NextApprover"" AS ""Next Approver"", T0.""Creator"", CASE ""U_Z_IsApp"" WHEN 'Y' THEN 'Yes'  ELSE 'No' END AS ""Approval Required"", ""U_Z_AppReqDate"" AS ""Requested Date"", T0.""U_Z_ApproveId"" FROM ""@Z_OFATA"" T0    INNER JOIN ""@Z_OAPPT"" T3 ON T3.""DocEntry"" = T0.""U_Z_ApproveId"" INNER JOIN ""@Z_APPT2"" T2 ON T3.""DocEntry"" = T2.""DocEntry"""
                    sQuery += " And (T0.""U_Z_CurrApprover"" = '" + oApplication.Company.UserName + "' OR T0.""U_Z_NextApprover"" = '" + oApplication.Company.UserName + "')"
                    sQuery += " And ISNULL(T2.""U_Z_AMan"",'N')='Y' AND ISNULL(T3.""U_Z_Active"",'N')='Y' and  ISNULL(T0.""U_Z_IsApp"",'N')='Y' and  T2.""U_Z_AUser"" = '" + oApplication.Company.UserName + "' And T3.""U_Z_DocType"" = '" + enDocType.ToString() + "' Order by T0.""DocEntry"" desc "
                End If
                

                'sQuery = " Select T0.DocEntry,T0.U_Z_Code,T0.U_Z_Name,CASE U_Z_TransType when 'L' then 'Location' when 'E' then 'Employee Transfer' else 'CostCenter' end AS U_Z_TransType,"
                'sQuery += " U_Z_FName,U_Z_TName,U_Z_ApproveId,CASE U_Z_DocStatus when 'D' then 'Draft' when 'N' then 'Confirm' when 'P' then 'Pending for Approval' when 'A' then 'Approved' when 'R' then 'Rejected' when 'C' then 'Cancel' else 'Close' end AS U_Z_DocStatus,"
                'sQuery += " U_Z_Attachment,U_Z_Remarks,T0.U_Z_AppStatus, U_Z_CurrApprover 'Current Approver',U_Z_NextApprover 'Next Approver', "
                'sQuery += " Case U_Z_IsApp when 'Y' then 'Yes' else 'No' End as  'Approval Required',U_Z_AppReqDate 'Requested Date',T0.Creator"
                'sQuery += " From [@Z_OFATA] T0 JOIN [@Z_OAPPT] T3 ON T3.DocEntry = T0.U_Z_ApproveId "
                'sQuery += " JOIN [@Z_APPT2] T2 ON T3.DocEntry = T2.DocEntry "
                'sQuery += " And IFNULL(T2.U_Z_AMan,'N')='Y' AND IFNULL(T3.""U_Z_Active"",'N')='Y' and  IFNULL(T0.""U_Z_IsApp"",'N')='Y' and  T2.""U_Z_AUser"" = '" + oApplication.Company.UserName + "' And T3.U_Z_DocType = '" + enDocType.ToString() + "' Order by T0.DocEntry desc "
        End Select
        oTempDt = aForm.DataSources.DataTables.Item("dtDocumentList")
        oTempDt.ExecuteQuery(sQuery)
        oGrid.DataTable.ExecuteQuery(sQuery)
        SummaryDocument(aForm, enDocType)
        oApplication.Utilities.assignMatrixLineno(oGrid, aForm)
        oGrid.Columns.Item("RowsHeader").Click(0, False, False)
        aForm.Freeze(False)
    End Sub
    Private Sub formatDocument(ByVal aForm As SAPbouiCOM.Form, ByVal enDocType As modVariables.HeaderDoctype)
        Try
            aForm.Freeze(True)
            Dim oGrid As SAPbouiCOM.Grid
            Dim oGridCombo As SAPbouiCOM.ComboBoxColumn
            oGrid = aForm.Items.Item("1").Specific
            Select Case enDocType
                Case HeaderDoctype.Fix
                    oGrid.Columns.Item("DocEntry").TitleObject.Caption = "Document Number"
                    oEditTextColumn = oGrid.Columns.Item("DocEntry")
                    oEditTextColumn.LinkedObjectType = "Z_OAPPT"
                    oGrid.Columns.Item("U_Z_Code").TitleObject.Caption = "Item Code"
                    oEditTextColumn = oGrid.Columns.Item("U_Z_Code")
                    oEditTextColumn.LinkedObjectType = "4"
                    oGrid.Columns.Item("U_Z_Name").TitleObject.Caption = "Item Description"
                    oGrid.Columns.Item("U_Z_TransType").TitleObject.Caption = "Transaction Type"
                    oGrid.Columns.Item("U_Z_FName").TitleObject.Caption = "Transfer From"
                    oGrid.Columns.Item("U_Z_TName").TitleObject.Caption = "Transfer To"
                    oGrid.Columns.Item("U_Z_DocStatus").TitleObject.Caption = "Document Status"
                    oGrid.Columns.Item("U_Z_Attachment").TitleObject.Caption = "Attachments"
                    oEditTextColumn = oGrid.Columns.Item("U_Z_Attachment")
                    oEditTextColumn.LinkedObjectType = "Z_OFATA"
                    oGrid.Columns.Item("U_Z_Remarks").TitleObject.Caption = "Remarks"
                    oGrid.Columns.Item("U_Z_AppStatus").TitleObject.Caption = "Approval Status"
                    oGrid.Columns.Item("U_Z_AppStatus").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                    oGridCombo = oGrid.Columns.Item("U_Z_AppStatus")
                    oGridCombo.ValidValues.Add("P", "Pending")
                    oGridCombo.ValidValues.Add("A", "Approved")
                    oGridCombo.ValidValues.Add("C", "Cancelled")
                    oGridCombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
                    oGrid.Columns.Item("U_Z_ApproveId").Visible = False
                    oGrid.Columns.Item("Creator").Visible = False
                    oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                    oGrid.AutoResizeColumns()
            End Select
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub SummaryDocument(ByVal aForm As SAPbouiCOM.Form, ByVal enDocType As modVariables.HeaderDoctype)
        Try
            aForm.Freeze(True)
            Dim oGrid As SAPbouiCOM.Grid
            Dim oGridCombo As SAPbouiCOM.ComboBoxColumn
            oGrid = aForm.Items.Item("19").Specific
            Select Case enDocType
                Case HeaderDoctype.Fix
                    oGrid.Columns.Item("DocEntry").TitleObject.Caption = "Document Number"
                    oEditTextColumn = oGrid.Columns.Item("DocEntry")
                    oEditTextColumn.LinkedObjectType = "Z_OAPPT"
                    oGrid.Columns.Item("U_Z_Code").TitleObject.Caption = "Item Code"
                    oEditTextColumn = oGrid.Columns.Item("U_Z_Code")
                    oEditTextColumn.LinkedObjectType = "4"
                    oGrid.Columns.Item("U_Z_Name").TitleObject.Caption = "Item Description"
                    oGrid.Columns.Item("U_Z_TransType").TitleObject.Caption = "Transaction Type"
                    oGrid.Columns.Item("U_Z_FName").TitleObject.Caption = "Transfer From"
                    oGrid.Columns.Item("U_Z_TName").TitleObject.Caption = "Transfer To"
                    oGrid.Columns.Item("U_Z_DocStatus").TitleObject.Caption = "Document Status"
                    oGrid.Columns.Item("U_Z_Attachment").TitleObject.Caption = "Attachments"
                    oEditTextColumn = oGrid.Columns.Item("U_Z_Attachment")
                    oEditTextColumn.LinkedObjectType = "Z_OFATA"
                    oGrid.Columns.Item("U_Z_Remarks").TitleObject.Caption = "Remarks"
                    oGrid.Columns.Item("U_Z_AppStatus").TitleObject.Caption = "Approval Status"
                    oGrid.Columns.Item("U_Z_AppStatus").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                    oGridCombo = oGrid.Columns.Item("U_Z_AppStatus")
                    oGridCombo.ValidValues.Add("P", "Pending")
                    oGridCombo.ValidValues.Add("A", "Approved")
                    oGridCombo.ValidValues.Add("C", "Cancelled")
                    oGridCombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
                    oGrid.Columns.Item("U_Z_ApproveId").Visible = False
                    oGrid.Columns.Item("Creator").Visible = False
                    oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                    oGrid.AutoResizeColumns()
            End Select
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_FATransactionApp Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "_1" Then
                                    If oApplication.ApplProcedure.ApprovalValidation(oForm, HeaderDoctype.Fix) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If (pVal.ItemUID = "1" Or pVal.ItemUID = "19") And pVal.ColUID = "DocEntry" Then
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    For intRow As Integer = pVal.Row To pVal.Row
                                        If 1 = 1 Then
                                            Dim strcode As String
                                            strcode = oGrid.DataTable.GetValue("DocEntry", intRow)
                                            Dim objct As New clsFixedAssetTransaction
                                            objct.LoadForm1(strcode)
                                        End If
                                    Next
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If (pVal.ItemUID = "1" Or pVal.ItemUID = "19") And pVal.ColUID = "U_Z_Attachment" Then
                                    oGrid = oForm.Items.Item(pVal.ItemUID).Specific
                                    oGrid.Columns.Item("RowsHeader").Click(pVal.Row)
                                    oApplication.Utilities.LoadFiles(oGrid.DataTable.GetValue("U_Z_Attachment", pVal.Row))
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "12" Then
                                    oForm.PaneLevel = 1
                                    oGrid = oForm.Items.Item("1").Specific
                                    oGrid.Columns.Item("RowsHeader").Click(0)
                                ElseIf pVal.ItemUID = "13" Then
                                    oForm.PaneLevel = 2
                                    oGrid = oForm.Items.Item("19").Specific
                                    oGrid.Columns.Item("RowsHeader").Click(0)
                                End If
                                If pVal.ItemUID = "1" And pVal.ColUID = "RowsHeader" And pVal.Row > -1 Then
                                    oGrid = oForm.Items.Item("1").Specific
                                    Dim strDocEntry As String = oGrid.DataTable.GetValue("DocEntry", pVal.Row)
                                    oApplication.ApplProcedure.LoadViewHistory(oForm, HeaderDoctype.Fix, strDocEntry, "3")
                                    oCombobox = oForm.Items.Item("8").Specific
                                    oCombobox.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                    oApplication.Utilities.setEdittextvalue(oForm, "10", "")
                                ElseIf (pVal.ItemUID = "3" And pVal.ColUID = "RowsHeader" And pVal.Row > -1) Then
                                    'oApplication.Utilities.LoadStatusRemarks(oForm, pVal.Row)
                                ElseIf pVal.ItemUID = "_1" Then
                                    Dim intRet As Integer = oApplication.SBO_Application.MessageBox("Are you sure want to submit the document?", 2, "Yes", "No", "")
                                    If intRet = 1 Then
                                        oApplication.ApplProcedure.addUpdateDocument(oForm, HeaderDoctype.Fix)
                                        InitializationApproval(oForm, HeaderDoctype.Fix)
                                        ApprovalSummary(oForm, HeaderDoctype.Fix)
                                    End If
                                End If
                                If pVal.ItemUID = "19" And pVal.ColUID = "RowsHeader" And pVal.Row > -1 Then
                                    oGrid = oForm.Items.Item("19").Specific
                                    Dim strDocEntry As String = oGrid.DataTable.GetValue("DocEntry", pVal.Row)
                                    oApplication.ApplProcedure.LoadViewHistory(oForm, HeaderDoctype.Fix, strDocEntry, "20")
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                            Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                If oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Restore Or oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized Then
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
            Select Case pVal.MenuUID
                Case mnu_FATransactionApp
                    LoadForm(oForm)
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

    Private Sub reDrawForm(ByVal oForm As SAPbouiCOM.Form)
        Try
            oForm.Freeze(True)
            'Rectangle
            oForm.Items.Item("16").Width = oForm.Width - 25
            oForm.Items.Item("16").Height = oForm.Height - 100



            oForm.Items.Item("1").Height = (oForm.Items.Item("16").Height - 40) / 2
            oForm.Items.Item("1").Width = oForm.Items.Item("16").Width - 10

            oForm.Items.Item("4").Top = oForm.Items.Item("1").Top + oForm.Items.Item("1").Height + 20


            oForm.Items.Item("3").Top = oForm.Items.Item("4").Top + oForm.Items.Item("4").Height + 10
            oForm.Items.Item("3").Height = oForm.Items.Item("1").Height - 20
            oForm.Items.Item("3").Width = (oForm.Items.Item("1").Width + 50) / 2

            oForm.Items.Item("19").Height = (oForm.Items.Item("16").Height - 40) / 2
            oForm.Items.Item("19").Width = oForm.Items.Item("16").Width - 10


            oForm.Items.Item("20").Top = oForm.Items.Item("19").Top + oForm.Items.Item("19").Height + 20
            oForm.Items.Item("20").Width = oForm.Items.Item("19").Width
            oForm.Items.Item("20").Height = oForm.Items.Item("19").Height



            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
        End Try
    End Sub

End Class
