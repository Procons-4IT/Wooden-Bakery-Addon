Public Class clsAppHistory
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oCombo As SAPbouiCOM.ComboBoxColumn
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtDocumentList As SAPbouiCOM.DataTable
    Private dtHistoryList As SAPbouiCOM.DataTable
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Public Sub LoadForm(ByVal oForm As SAPbouiCOM.Form, ByVal DocNo As String, ByVal EnDocType As modVariables.HeaderDoctype)
        Try
            oForm = oApplication.Utilities.LoadForm(xm_AppHistory, frm_AppHistory)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            LoadViewHistory(oForm, DocNo, EnDocType)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Public Sub LoadViewHistory(ByVal aForm As SAPbouiCOM.Form, ByVal strDocEntry As String, ByVal enDocType As modVariables.HeaderDoctype)
        Try
            aForm.Freeze(True)
            Dim sQuery As String
            oGrid = aForm.Items.Item("1").Specific
            Select Case enDocType
                Case HeaderDoctype.Fix
                    sQuery = " Select ""DocEntry"",""U_Z_DocEntry"",""U_Z_DocType"",""U_Z_EmpId"",""U_Z_EmpName"",""U_Z_ApproveBy"",""CreateDate"" ,""CreateTime"",""UpdateDate"",""UpdateTime"",""U_Z_AppStatus"",""U_Z_Remarks"" From ""@Z_APHIS"" "
                    sQuery += " Where ""U_Z_DocType"" = '" + enDocType.ToString() + "'"
                    sQuery += " And ""U_Z_DocEntry"" = '" + strDocEntry + "'"
                    oGrid.DataTable.ExecuteQuery(sQuery)
                    formatHistory(aForm, enDocType)
                    oApplication.Utilities.assignMatrixLineno(oGrid, aForm)
                Case HeaderDoctype.Spl
                    sQuery = " Select ""DocEntry"",""U_Z_DocEntry"",""U_Z_DocType"",""U_Z_EmpId"",""U_Z_EmpName"",""U_Z_ApproveBy"",""CreateDate"" ,""CreateTime"",""UpdateDate"",""UpdateTime"",""U_Z_AppStatus"",""U_Z_Remarks"" From ""@Z_APHIS"" "
                    sQuery += " Where ""U_Z_DocType"" = '" + enDocType.ToString() + "'"
                    sQuery += " And ""U_Z_DocEntry"" = '" + strDocEntry + "'"
                    oGrid.DataTable.ExecuteQuery(sQuery)
                    formatHistory(aForm, enDocType)
                    oApplication.Utilities.assignMatrixLineno(oGrid, aForm)
            End Select
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub formatHistory(ByVal aForm As SAPbouiCOM.Form, ByVal enDocType As modVariables.HeaderDoctype)
        Try
            aForm.Freeze(True)
            Dim oGrid As SAPbouiCOM.Grid
            Dim oComboBox As SAPbouiCOM.ComboBox
            Dim oGridCombo As SAPbouiCOM.ComboBoxColumn
            Dim oEditTextColumn As SAPbouiCOM.EditTextColumn
            Select Case enDocType
                Case HeaderDoctype.Fix
                    oGrid = aForm.Items.Item("1").Specific
                    oGrid.Columns.Item("DocEntry").Visible = False
                    oGrid.Columns.Item("U_Z_DocEntry").TitleObject.Caption = "Reference No."
                    oGrid.Columns.Item("U_Z_DocEntry").Visible = False
                    oGrid.Columns.Item("U_Z_DocType").Visible = False
                    oGrid.Columns.Item("U_Z_EmpId").TitleObject.Caption = "Employee ID"
                    oEditTextColumn = oGrid.Columns.Item("U_Z_EmpId")
                    oEditTextColumn.LinkedObjectType = "171"
                    oGrid.Columns.Item("U_Z_EmpName").TitleObject.Caption = "Employee Name"
                    oGrid.Columns.Item("U_Z_ApproveBy").TitleObject.Caption = "Approved By"
                    oGrid.Columns.Item("U_Z_AppStatus").TitleObject.Caption = "Approved Status"
                    oGrid.Columns.Item("U_Z_AppStatus").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                    oGridCombo = oGrid.Columns.Item("U_Z_AppStatus")
                    oGridCombo.ValidValues.Add("A", "Approved")
                    oGridCombo.ValidValues.Add("R", "Rejected")
                    oGridCombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
                    oGrid.Columns.Item("U_Z_Remarks").TitleObject.Caption = "Remarks"
                    oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                    oGrid.AutoResizeColumns()
                Case HeaderDoctype.Spl
                    oGrid = aForm.Items.Item("1").Specific
                    oGrid.Columns.Item("DocEntry").Visible = False
                    oGrid.Columns.Item("U_Z_DocEntry").TitleObject.Caption = "Reference No."
                    oGrid.Columns.Item("U_Z_DocEntry").Visible = False
                    oGrid.Columns.Item("U_Z_DocType").Visible = False
                    oGrid.Columns.Item("U_Z_EmpId").TitleObject.Caption = "Employee ID"
                    oEditTextColumn = oGrid.Columns.Item("U_Z_EmpId")
                    oEditTextColumn.LinkedObjectType = "171"
                    oGrid.Columns.Item("U_Z_EmpName").TitleObject.Caption = "Employee Name"
                    oGrid.Columns.Item("U_Z_ApproveBy").TitleObject.Caption = "Approved By"
                    oGrid.Columns.Item("U_Z_AppStatus").TitleObject.Caption = "Approved Status"
                    oGrid.Columns.Item("U_Z_AppStatus").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                    oGridCombo = oGrid.Columns.Item("U_Z_AppStatus")
                    oGridCombo.ValidValues.Add("A", "Approved")
                    oGridCombo.ValidValues.Add("R", "Rejected")
                    oGridCombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
                    oGrid.Columns.Item("U_Z_Remarks").TitleObject.Caption = "Remarks"
                    oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                    oGrid.AutoResizeColumns()
            End Select
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

End Class
