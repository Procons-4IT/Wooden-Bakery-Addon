Public Class clsDeliveryDoc
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oMode As SAPbouiCOM.BoFormMode
    Dim oDBDataSource As SAPbouiCOM.DBDataSource
    Dim oDBDataSourceLines As SAPbouiCOM.DBDataSource
    Private RowtoDelete As Integer
    Private count As Integer
    Dim MatrixId As Integer
    Private oColumn As SAPbouiCOM.Column
    Private oDts As SAPbouiCOM.DataTable

    Private oRecordSet As SAPbobsCOM.Recordset
    Private strQuery As String

    Public Sub New()
        MyBase.New()

    End Sub

    Public Sub LoadForm()
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Z_ODEL) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_Z_ODEL, frm_Z_ODEL)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.DataBrowser.BrowseBy = "4"
        oForm.EnableMenu(mnu_DELETE_ROW, True)
        oForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
        initialize(oForm)
        oForm.Freeze(False)
    End Sub

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.BeforeAction
                Case True
                    Select Case pVal.MenuUID
                        Case mnu_DELETE_ROW
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            If pVal.BeforeAction = True And oForm.TypeEx = frm_Z_ODEL Then
                                If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If

                            End If
                    End Select
                Case False
                    Select Case pVal.MenuUID
                        Case mnu_Z_ODEL
                            If pVal.BeforeAction = False Then
                                LoadForm()
                            End If
                        Case mnu_ADD_ROW
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        Case mnu_DELETE_ROW
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            If pVal.BeforeAction = False Then
                                RefereshDeleteRow(oForm, "11")
                                If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                            End If
                        Case mnu_ADD
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            oForm.Items.Item("4").Enabled = True
                            initialize(oForm)
                            oForm.Items.Item("1").Enabled = True
                            oForm.Items.Item("11").Enabled = True
                            oForm.Items.Item("14").Enabled = True
                            oForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Case mnu_FIND
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            If pVal.BeforeAction = False Then
                                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                                oForm.Items.Item("4").Enabled = True
                                oForm.Items.Item("11").Enabled = True
                                oForm.Items.Item("1").Enabled = True
                            End If
                        Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                    End Select
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Z_ODEL Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If Validation(oForm) Then
                                        If oApplication.SBO_Application.MessageBox("Do you want to confirm the information?", , "Yes", "No") = 2 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    Else
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "10" Then
                                    If pVal.CharPressed = "9" Or pVal.CharPressed = "13" Then
                                        oMatrix = oForm.Items.Item("11").Specific
                                        Dim otest As SAPbobsCOM.Recordset
                                        otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        Dim strQuery, strScanedBarcode, strDocType, strDocNum, strTable As String
                                        strScanedBarcode = oApplication.Utilities.getEdittextvalue(oForm, "10")
                                        If strScanedBarcode.Length > 2 Then
                                            strDocNum = strScanedBarcode.Substring(2, strScanedBarcode.Length - 2)
                                            Select Case strScanedBarcode.Substring(0, 2)
                                                Case "13"
                                                    strQuery = "Select ""DocEntry"", ""DocNum"", ""CardCode"", ""CardName"",""DocDate"" From OINV where ""DocNum""='" & strDocNum & "'"
                                                    strDocType = "Invoice"
                                                    strTable = "OINV"
                                                Case "14"
                                                    strQuery = "Select ""DocEntry"", ""DocNum"", ""CardCode"", ""CardName"",""DocDate"" From ORIN where ""DocNum""='" & strDocNum & "'"
                                                    strDocType = "Credit Note"
                                                    strTable = "ORIN"
                                                Case "24"
                                                    strQuery = "Select ""DocEntry"", ""DocNum"", ""CardCode"", ""CardName"",""DocDate"" From ORCT where ""DocNum""='" & strDocNum & "'"
                                                    strDocType = "Incoming Payment"
                                                    strTable = "ORCT"
                                                Case "46"
                                                    strQuery = "Select ""DocEntry"", ""DocNum"", ""CardCode"", ""CardName"",""DocDate"" From OVPM where ""DocNum""='" & strDocNum & "'"
                                                    strDocType = "Vendor Payment"
                                                    strTable = "OVPM"
                                            End Select
                                        End If
                                        If strQuery <> "" Then
                                            'strQuery = "Select ""DocEntry"", ""DocNum"", ""CardCode"", ""CardName"",""DocDate"" From OINV where ""DocNum""='" & oApplication.Utilities.getEdittextvalue(oForm, "10") & "'"
                                            otest.DoQuery(strQuery)
                                            If otest.RecordCount > 0 Then
                                                If oApplication.Utilities.getMatrixValues(oMatrix, "V_0", oMatrix.RowCount) <> "" Then
                                                    oMatrix.AddRow(1, oMatrix.RowCount)
                                                End If
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, otest.Fields.Item("DocEntry").Value)
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", oMatrix.RowCount, otest.Fields.Item("DocNum").Value)
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_2", oMatrix.RowCount, otest.Fields.Item("CardCode").Value)
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", oMatrix.RowCount, otest.Fields.Item("CardName").Value)
                                                Dim dtDate As Date = otest.Fields.Item("DocDate").Value
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_6", oMatrix.RowCount, dtDate.ToString("yyyyMMdd"))
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_5", oMatrix.RowCount, strDocType)
                                                oApplication.Utilities.SetMatrixValues(oMatrix, "V_7", oMatrix.RowCount, strTable)
                                                oApplication.Utilities.setEdittextvalue(oForm, "10", "")
                                            Else
                                                oApplication.SBO_Application.MessageBox("Scanned document does not exists")
                                                oApplication.Utilities.setEdittextvalue(oForm, "10", "")
                                            End If
                                        Else
                                            oApplication.SBO_Application.MessageBox("Scanned document does not support in this module")
                                            oApplication.Utilities.setEdittextvalue(oForm, "10", "")
                                        End If
                                        oForm.Items.Item("10").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    End If

                                End If


                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)


                            Case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If (pVal.ItemUID = "11" And pVal.ColUID = "V_0") Then

                                End If

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "1"
                                        If pVal.Action_Success Then
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                                initialize(oForm)
                                                oForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            End If
                                        End If
                                    Case "14"
                                        RefereshDeleteRow(oForm, 11)

                                End Select

                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                alldataSource(oForm)
                                If (pVal.ItemUID = "11") Then
                                    intSelectedMatrixrow = pVal.Row
                                End If


                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                        End Select
                End Select

            End If


        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form)
        Try
            alldataSource(oForm)
            oMatrix = oForm.Items.Item("11").Specific
            oMatrix.LoadFromDataSource()
            oMatrix.AddRow(1, -1)
            oMatrix.FlushToDataSource()
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            Dim strCode As String = oApplication.Utilities.getMaxCode("@Z_ODEL", "DocEntry")
            oApplication.Utilities.setEdittextvalue(oForm, "4", strCode)
            oApplication.Utilities.setEdittextvalue(oForm, "6", System.DateTime.Now.ToString("yyyyMMdd"))
            oMatrix.AutoResizeColumns()
            oForm.Update()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub AssignLineNo(ByVal aForm As SAPbouiCOM.Form, ByVal strItem As String)
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item(strItem).Specific
            alldataSource(oForm)
            oMatrix.FlushToDataSource()

            Select Case strItem
                Case "11"
                    For count = 1 To oDBDataSourceLines.Size
                        oDBDataSourceLines.SetValue("LineId", count - 1, count)
                    Next
            End Select

            oMatrix.LoadFromDataSource()
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

#Region "AddRow /Delete Row"

    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form, ByVal strItem As String)
        Try
            oMatrix = aForm.Items.Item(strItem).Specific
            alldataSource(oForm)

            Me.RowtoDelete = intSelectedMatrixrow
            oDBDataSourceLines.RemoveRecord(Me.RowtoDelete - 1)
            oMatrix.LoadFromDataSource()
            oMatrix.FlushToDataSource()
            For count = 0 To oDBDataSourceLines.Size - 1
                oDBDataSourceLines.SetValue("LineId", count, count + 1)
            Next
            oMatrix.LoadFromDataSource()
            oMatrix.FlushToDataSource()
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

#End Region

    Private Function Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            If oApplication.Utilities.getEdittextvalue(aForm, "8") = "" Then
                oApplication.Utilities.Message("Delivery Date is missing....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        End If
        Return True

    End Function

    Private Sub alldataSource(ByVal aForm As SAPbouiCOM.Form)
        Try
            oDBDataSource = oForm.DataSources.DBDataSources.Item("@Z_ODEL")
            oDBDataSourceLines = oForm.DataSources.DBDataSources.Item("@Z_DEL1")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Dim oRecordSet As SAPbobsCOM.Recordset
        Try
            Select Case BusinessObjectInfo.BeforeAction
                Case True
                Case False
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                            If BusinessObjectInfo.ActionSuccess = True Then
                                oForm.Items.Item("1").Enabled = False
                                oForm.Items.Item("11").Enabled = False
                                oForm.Items.Item("14").Enabled = False

                            End If
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                            If BusinessObjectInfo.ActionSuccess Then
                                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                                Dim oXmlDoc As System.Xml.XmlDocument = New Xml.XmlDocument()
                                Dim strTable As String
                                oXmlDoc.LoadXml(BusinessObjectInfo.ObjectKey)
                                Dim DocEntry As String = oXmlDoc.SelectSingleNode("/Delivery_DocumentParams/DocEntry").InnerText
                                Try
                                    Dim oURecordSet As SAPbobsCOM.Recordset
                                    oURecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

                                    Dim strQuery As String = "Select T0.""DocEntry"",T1.""U_Z_DelDate"",T0.""U_Z_DocEntry"",T0.""U_Z_TarTable"" From ""@Z_DEL1"" T0 JOIN ""@Z_ODEL"" T1 On T1.""DocEntry"" = T0.""DocEntry"" Where T0.""DocEntry"" = '" & DocEntry & "'"
                                    oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRecordSet.DoQuery(strQuery)
                                    If Not oRecordSet.EoF Then
                                        While Not oRecordSet.EoF
                                            'MessageBox.Show(oRecordSet.Fields.Item("U_Z_DocEntry").Value)
                                            strTable = oRecordSet.Fields.Item("U_Z_TarTable").Value
                                            Dim dtDelDate As Date = oRecordSet.Fields.Item("U_Z_DelDate").Value
                                            ' strQuery = "Update OINV Set ""U_Z_DelDate""='" & dtDelDate.ToString("yyyy-MM-dd") & "' , ""U_Z_IsDel""='Y' , ""U_Z_DelRef""='" & oRecordSet.Fields.Item("DocEntry").Value & "' Where ""DocEntry"" = '" & oRecordSet.Fields.Item("U_Z_DocEntry").Value & "'"
                                            strQuery = "Update " & strTable & " Set ""U_Z_DelDate""='" & dtDelDate.ToString("yyyy-MM-dd") & "' , ""U_Z_IsDel""='Y' , ""U_Z_DelRef""='" & oRecordSet.Fields.Item("DocEntry").Value & "' Where ""DocEntry"" = '" & oRecordSet.Fields.Item("U_Z_DocEntry").Value & "'"
                                            oURecordSet.DoQuery(strQuery)

                                            oRecordSet.MoveNext()
                                        End While
                                    End If

                                Catch ex As Exception
                                    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                                    oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End Try
                            End If


                    End Select
            End Select

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

End Class