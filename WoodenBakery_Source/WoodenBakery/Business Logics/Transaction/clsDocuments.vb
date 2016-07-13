Public Class clsDocuments
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
    Private strQuery As String
    Private oRecordSet As SAPbobsCOM.Recordset
    Private intSelectedRow As Integer
    
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try

            If pVal.FormTypeEx = frm_ARInvoicePayment Then
                If pVal.Before_Action = True Then
                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                    If pVal.ItemUID = "1" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        frm_InvoiceForm = oForm
                    End If
                End If
            End If

            If pVal.FormTypeEx = frm_PaymentMeans Then
                If pVal.Before_Action = False Then
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_LOAD And pVal.Before_Action = False Then
                        oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                        oApplication.Utilities.PopulateDocTotaltoPaymentMeans(frm_InvoiceForm, oForm)
                        ' oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN And pVal.CharPressed = 13 Then
                        'MsgBox(pVal.CharPressed)
                        'oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                        'If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        '    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        'End If
                        'If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        '    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        'End If

                    End If

                    'If pVal.EventType = SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD Then
                    '    '    blnInvoiceForm = False
                    'End If
                Else
                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.Before_Action = True Then
                        If blnInvoiceForm = True Then
                            If pVal.ItemUID = "1" Or pVal.ItemUID = "2" Or pVal.ItemUID = "6" Or pVal.ItemUID = "38" Then
                                Dim oItem As SAPbouiCOM.Item
                                oItem = oForm.Items.Item(pVal.ItemUID)
                                If oItem.Type <> SAPbouiCOM.BoFormItemTypes.it_BUTTON And oItem.Type <> SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON And pVal.ItemUID <> "6" Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                                If pVal.ItemUID = 1 Or pVal.ItemUID = "2" Then
                                    blnInvoiceForm = False
                                End If
                            Else
                                Dim oItem As SAPbouiCOM.Item
                                oItem = oForm.Items.Item(pVal.ItemUID)
                                If oItem.Type <> SAPbouiCOM.BoFormItemTypes.it_BUTTON And oItem.Type <> SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON Then
                                    BubbleEvent = False
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED And pVal.Before_Action = True Then
                        If blnInvoiceForm = True Then
                            Dim oItem As SAPbouiCOM.Item
                            oItem = oForm.Items.Item(pVal.ItemUID)
                            If oItem.Type <> SAPbouiCOM.BoFormItemTypes.it_BUTTON Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                            If pVal.ItemUID = "1" Or pVal.ItemUID = "2" Or pVal.ItemUID = "6" Then ' Or pVal.ItemUID = "38" Then
                                If pVal.ItemUID = 1 Or pVal.ItemUID = "2" Then
                                    blnInvoiceForm = False
                                End If
                            Else

                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    End If

                    If pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN And pVal.Before_Action = True And (pVal.CharPressed <> 9 And pVal.CharPressed <> 13) Then
                        If blnInvoiceForm = True Then
                            Dim oItem As SAPbouiCOM.Item
                            oItem = oForm.Items.Item(pVal.ItemUID)
                            If oItem.Type <> SAPbouiCOM.BoFormItemTypes.it_BUTTON Then
                                BubbleEvent = False
                                Exit Sub
                            End If
                            If pVal.ItemUID = "1" Or pVal.ItemUID = "2" Or pVal.ItemUID = "6" Then 'Or pVal.ItemUID = "38" Then

                            Else

                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If


            If pVal.FormTypeEx = frm_SalesOrder Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "38" Then
                                    oMatrix = oForm.Items.Item("38").Specific
                                    Dim strRef As String = oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_PrRef", pVal.Row)
                                    If strRef <> "" Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                                If pVal.ItemUID = "38" And (pVal.ColUID = "1" Or pVal.ColUID = "3") And pVal.CharPressed = 9 Then
                                    Dim dtDeliverydate As Date
                                    If oApplication.Utilities.getEdittextvalue(oForm, "10") = "" Then
                                        oApplication.Utilities.Message("Posting Date is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK, SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("38").Specific
                                If pVal.ItemUID = "38" Then
                                    Dim strRef As String = oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_PrRef", pVal.Row)
                                    If strRef <> "" Then
                                        oApplication.Utilities.Message("Promotion details already applied for this Row. You Cannot edit Row...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If

                                If pVal.ItemUID = "38" And (pVal.ColUID = "U_Z_PrmApp" Or pVal.ColUID = "U_Z_PrCode" Or pVal.ColUID = "U_SPDocEty" Or pVal.ColUID = "U_Z_PrRef" Or pVal.ColUID = "U_PrLine") Then
                                    BubbleEvent = False
                                    Exit Sub
                                ElseIf pVal.ItemUID = "38" And pVal.Row > 0 Then
                                    If CType(oMatrix.Columns.Item("U_Z_PrmApp").Cells().Item(pVal.Row).Specific, SAPbouiCOM.ComboBox).Value = "Y" Then
                                        If pVal.ColUID = "1" Or pVal.ColUID = "2" Or pVal.ColUID = "3" Or pVal.ColUID = "11" Then
                                            oApplication.Utilities.Message("Promotion details already applied for this Row. You should delete the  .", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                        End If
                                    ElseIf (CType(oMatrix.Columns.Item("U_Z_PrCode").Cells().Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.Trim().Length > 0) Then
                                        If pVal.ColUID = "1" Or pVal.ColUID = "2" Or pVal.ColUID = "3" Or pVal.ColUID = "11" Or pVal.ColUID = "14" Or pVal.ColUID = "15" Or pVal.ColUID = "21" Then
                                            '  oApplication.Utilities.Message("Promotion details already applied for this Row. You should delete the  .", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            ' BubbleEvent = False
                                        End If
                                    ElseIf pVal.ItemUID = "38" And (pVal.ColUID = "14" Or pVal.ColUID = "15" Or pVal.ColUID = "17" Or pVal.ColUID = "21") Then
                                        If oApplication.Utilities.getMatrixValues(oMatrix, "31", pVal.Row) <> "" And oApplication.Utilities.getMatrixValues(oMatrix, "U_SPDocEty", pVal.Row) <> "" Then
                                            oApplication.Utilities.Message("Special Price is Linked to Selected Row...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If ValidatePromotion(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                                If pVal.ItemUID = "_2" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    oMatrix = oForm.Items.Item("38").Specific
                                    If oForm.Items.Item("12").Specific.value.ToString.Length = 0 Then
                                        oApplication.Utilities.Message("Please Enter Delivery Date to Apply Promotion...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    ElseIf oMatrix.RowCount = 1 Then
                                        oApplication.Utilities.Message("Add Items To Apply Promotion...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                    If oApplication.SBO_Application.MessageBox("No Possible to Change Items When Promotion Applied Want to Continue?", , "Yes", "No") = 2 Then
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        oForm.Freeze(True)
                                        applyPromotion(oForm)
                                        oForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        oForm.Freeze(False)
                                    End If
                                ElseIf pVal.ItemUID = "2_" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    oMatrix = oForm.Items.Item("38").Specific
                                    If oApplication.SBO_Application.MessageBox("Sure you Want to clear promotions Continue?", , "Yes", "No") = 2 Then
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        oForm.Freeze(True)
                                        clearPromotion(oForm)
                                        oForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        oForm.Freeze(False)
                                    End If
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                initializeControls(oForm)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "38" And (pVal.ColUID = "1" Or pVal.ColUID = "3") And pVal.CharPressed = 9 Then
                                    Dim dtDeliverydate As Date
                                    Dim stDate1, strItemCode As String
                                    oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                    stDate1 = oApplication.Utilities.getEdittextvalue(oForm, "10")
                                    strItemCode = oApplication.Utilities.getMatrixValues(oMatrix, "1", pVal.Row)
                                    Dim oRec As SAPbobsCOM.Recordset
                                    oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRec.DoQuery("Select * from OITM where ""ItemCode""='" & strItemCode & "'")
                                    If (oRec.RecordCount > 0) Then
                                        strItemCode = oRec.Fields.Item("ItemCode").Value
                                    Else
                                        strItemCode = ""
                                    End If

                                    If stDate1 <> "" And strItemCode <> "" Then
                                        dtDeliverydate = oApplication.Utilities.GetDateTimeValue(stDate1)
                                        dtDeliverydate = oApplication.Utilities.GetDeliveryDate(oApplication.Utilities.getEdittextvalue(oForm, "4"), oApplication.Utilities.getMatrixValues(oMatrix, "1", pVal.Row), dtDeliverydate)
                                        Dim stdate As String = dtDeliverydate.ToString("yyyyMMdd")
                                        Try
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "25", pVal.Row, stdate)
                                        Catch ex As Exception
                                        End Try
                                        oMatrix.Columns.Item("11").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                    End If
                                End If


                                If pVal.ItemUID = "10" And pVal.CharPressed = 9 And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    oCombobox = oForm.Items.Item("3").Specific
                                    If oCombobox.Selected.Value = "I" Then
                                        oMatrix = oForm.Items.Item("38").Specific
                                        oForm.Freeze(True)
                                        Dim dtDeliverydate, dtSODeliveryDate As Date
                                        Dim stDate1, strItemCode As String
                                        stDate1 = oApplication.Utilities.getEdittextvalue(oForm, "10")
                                        strItemCode = oApplication.Utilities.getMatrixValues(oMatrix, "1", pVal.Row)
                                        Dim oRec As SAPbobsCOM.Recordset
                                        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRec.DoQuery("Select * from OITM where ""ItemCode""='" & strItemCode & "'")
                                        If (oRec.RecordCount > 0) Then
                                            strItemCode = oRec.Fields.Item("ItemCode").Value
                                        Else
                                            strItemCode = ""
                                        End If

                                        If stDate1 <> "" And strItemCode <> "" Then
                                            dtDeliverydate = oApplication.Utilities.GetDateTimeValue(stDate1)
                                            dtSODeliveryDate = oApplication.Utilities.getsalesOrderDeliveryDate(oApplication.Utilities.getEdittextvalue(oForm, "4"), dtDeliverydate)
                                            oApplication.Utilities.setEdittextvalue(oForm, "12", dtSODeliveryDate.ToString("yyyyMMdd"))
                                            oApplication.SBO_Application.SendKeys("{TAB}")
                                            For intRow As Integer = 1 To oMatrix.RowCount
                                                dtDeliverydate = oApplication.Utilities.GetDeliveryDate(oApplication.Utilities.getEdittextvalue(oForm, "4"), oApplication.Utilities.getMatrixValues(oMatrix, "1", intRow), dtDeliverydate)
                                                Dim stdate As String = dtDeliverydate.ToString("yyyyMMdd")
                                                Try
                                                    oApplication.Utilities.SetMatrixValues(oMatrix, "25", intRow, stdate)
                                                Catch ex As Exception
                                                End Try
                                            Next
                                        End If


                                        oForm.Freeze(False)
                                    End If

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                        End Select
                End Select
            End If




            If pVal.FormTypeEx = "TESS" Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "38" Then
                                    oMatrix = oForm.Items.Item("38").Specific
                                    Dim strRef As String = oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_PrRef", pVal.Row)
                                    If strRef <> "" Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                                If pVal.ItemUID = "38" And (pVal.ColUID = "1" Or pVal.ColUID = "3") And pVal.CharPressed = 9 Then
                                    Dim dtDeliverydate As Date
                                    If oApplication.Utilities.getEdittextvalue(oForm, "10") = "" Then
                                        oApplication.Utilities.Message("Posting Date is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK, SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("38").Specific
                                If pVal.ItemUID = "38" Then
                                    Dim strRef As String = oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_PrRef", pVal.Row)
                                    If strRef <> "" Then
                                        oApplication.Utilities.Message("Promotion details already applied for this Row. You Cannot edit Row...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If

                                If pVal.ItemUID = "38" And (pVal.ColUID = "U_Z_PrmApp" Or pVal.ColUID = "U_Z_PrCode" Or pVal.ColUID = "U_SPDocEty" Or pVal.ColUID = "U_Z_PrRef" Or pVal.ColUID = "U_PrLine") Then
                                    BubbleEvent = False
                                    Exit Sub
                                ElseIf pVal.ItemUID = "38" And pVal.Row > 0 Then
                                    If CType(oMatrix.Columns.Item("U_Z_PrmApp").Cells().Item(pVal.Row).Specific, SAPbouiCOM.ComboBox).Value = "Y" Then
                                        If pVal.ColUID = "1" Or pVal.ColUID = "2" Or pVal.ColUID = "3" Or pVal.ColUID = "11" Then
                                            oApplication.Utilities.Message("Promotion details already applied for this Row. You should delete the  .", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                        End If
                                    ElseIf (CType(oMatrix.Columns.Item("U_Z_PrCode").Cells().Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.Trim().Length > 0) Then
                                        If pVal.ColUID = "1" Or pVal.ColUID = "2" Or pVal.ColUID = "3" Or pVal.ColUID = "11" Or pVal.ColUID = "14" Or pVal.ColUID = "15" Or pVal.ColUID = "21" Then
                                            oApplication.Utilities.Message("Promotion details already applied for this Row. You should delete the  .", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                        End If
                                    ElseIf pVal.ItemUID = "38" And (pVal.ColUID = "14" Or pVal.ColUID = "15" Or pVal.ColUID = "17" Or pVal.ColUID = "21") Then
                                        If oApplication.Utilities.getMatrixValues(oMatrix, "31", pVal.Row) <> "" And oApplication.Utilities.getMatrixValues(oMatrix, "U_SPDocEty", pVal.Row) <> "" Then
                                            oApplication.Utilities.Message("Special Price is Linked to Selected Row...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "_2" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    oMatrix = oForm.Items.Item("38").Specific
                                    If oForm.Items.Item("12").Specific.value.ToString.Length = 0 Then
                                        oApplication.Utilities.Message("Please Enter Delivery Date to Apply Promotion...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    ElseIf oMatrix.RowCount = 1 Then
                                        oApplication.Utilities.Message("Add Items To Apply Promotion...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                    If oApplication.SBO_Application.MessageBox("No Possible to Change Items When Promotion Applied Want to Continue?", , "Yes", "No") = 2 Then
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        oForm.Freeze(True)
                                        applyPromotion(oForm)
                                        oForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        oForm.Freeze(False)
                                    End If
                                ElseIf pVal.ItemUID = "2_" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    oMatrix = oForm.Items.Item("38").Specific
                                    If oApplication.SBO_Application.MessageBox("Sure you Want to clear promotions Continue?", , "Yes", "No") = 2 Then
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        oForm.Freeze(True)
                                        clearPromotion(oForm)
                                        oForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        oForm.Freeze(False)
                                    End If
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                initializeControls(oForm)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "38" And (pVal.ColUID = "1" Or pVal.ColUID = "3") And pVal.CharPressed = 9 Then
                                    Dim dtDeliverydate As Date
                                    Dim stDate1, strItemCode As String
                                    oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                    stDate1 = oApplication.Utilities.getEdittextvalue(oForm, "10")
                                    strItemCode = oApplication.Utilities.getMatrixValues(oMatrix, "1", pVal.Row)
                                    Dim oRec As SAPbobsCOM.Recordset
                                    oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    oRec.DoQuery("Select * from OITM where ""ItemCode""='" & strItemCode & "'")
                                    If (oRec.RecordCount > 0) Then
                                        strItemCode = oRec.Fields.Item("ItemCode").Value
                                    Else
                                        strItemCode = ""
                                    End If

                                    If stDate1 <> "" And strItemCode <> "" Then
                                        dtDeliverydate = oApplication.Utilities.GetDateTimeValue(stDate1)
                                        dtDeliverydate = oApplication.Utilities.GetDeliveryDate(oApplication.Utilities.getEdittextvalue(oForm, "4"), oApplication.Utilities.getMatrixValues(oMatrix, "1", pVal.Row), dtDeliverydate)
                                        Dim stdate As String = dtDeliverydate.ToString("yyyyMMdd")
                                        Try
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "25", pVal.Row, stdate)
                                        Catch ex As Exception
                                        End Try
                                    End If
                                End If


                                If pVal.ItemUID = "10" And pVal.CharPressed = 9 And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    oCombobox = oForm.Items.Item("3").Specific
                                    If oCombobox.Selected.Value = "I" Then
                                        oMatrix = oForm.Items.Item("38").Specific
                                        oForm.Freeze(True)
                                        Dim dtDeliverydate, dtSODeliveryDate As Date
                                        Dim stDate1, strItemCode As String
                                        stDate1 = oApplication.Utilities.getEdittextvalue(oForm, "10")
                                        strItemCode = oApplication.Utilities.getMatrixValues(oMatrix, "1", pVal.Row)
                                        Dim oRec As SAPbobsCOM.Recordset
                                        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        oRec.DoQuery("Select * from OITM where ""ItemCode""='" & strItemCode & "'")
                                        If (oRec.RecordCount > 0) Then
                                            strItemCode = oRec.Fields.Item("ItemCode").Value
                                        Else
                                            strItemCode = ""
                                        End If

                                        If stDate1 <> "" And strItemCode <> "" Then
                                            dtDeliverydate = oApplication.Utilities.GetDateTimeValue(stDate1)
                                            dtSODeliveryDate = oApplication.Utilities.getsalesOrderDeliveryDate(oApplication.Utilities.getEdittextvalue(oForm, "4"), dtDeliverydate)
                                            oApplication.Utilities.setEdittextvalue(oForm, "12", dtSODeliveryDate.ToString("yyyyMMdd"))
                                            oApplication.SBO_Application.SendKeys("{TAB}")
                                            For intRow As Integer = 1 To oMatrix.RowCount
                                                dtDeliverydate = oApplication.Utilities.GetDeliveryDate(oApplication.Utilities.getEdittextvalue(oForm, "4"), oApplication.Utilities.getMatrixValues(oMatrix, "1", intRow), dtDeliverydate)
                                                Dim stdate As String = dtDeliverydate.ToString("yyyyMMdd")
                                                Try
                                                    oApplication.Utilities.SetMatrixValues(oMatrix, "25", intRow, stdate)
                                                Catch ex As Exception
                                                End Try
                                            Next
                                        End If


                                        oForm.Freeze(False)
                                    End If

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                        End Select
                End Select
            End If


        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            If oForm.TypeEx = frm_SalesOrder Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.MenuUID
                            Case mnu_DELETE_ROW
                                Dim intRowCount As Integer = intSelectedMatrixrow 'oMatrix.GetCellFocus().rowIndex
                                If oMatrix.Columns.Item("U_Z_PrRef").Cells().Item(intRowCount).Specific.value <> "" Then
                                    BubbleEvent = False
                                    oApplication.Utilities.Message("Promotion already applied to Remove Clear Promotion...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    Exit Sub
                                End If
                        End Select
                    Case False
                        Select Case pVal.MenuUID
                            Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                            Case mnu_ADD
                                oForm.Items.Item("38").Enabled = True
                                oForm.Items.Item("_2").Enabled = True
                            Case mnu_CPRL_O
                                If Not oForm.Items.Item("4").Specific.value = "" Then
                                    Dim objPromList As clsCustPromotionList
                                    objPromList = New clsCustPromotionList
                                    objPromList.LoadForm(oForm.Items.Item("4").Specific.value, "C")
                                Else
                                    oApplication.Utilities.Message("Select Customer to Get Promotion List...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                End If
                                Dim oMenuItem As SAPbouiCOM.MenuItem
                                oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                                If oMenuItem.SubMenus.Exists(pVal.MenuUID) Then
                                    oApplication.SBO_Application.Menus.RemoveEx(pVal.MenuUID)
                                End If
                        End Select
                End Select
            End If
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

#Region "Right Click"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
            If oForm.TypeEx = frm_SalesOrder Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                If (eventInfo.BeforeAction = True) Then
                    intSelectedMatrixrow = eventInfo.Row
                    Try

                        'Promotion List
                        If Not oMenuItem.SubMenus.Exists(mnu_CPRL_O) Then
                            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                            oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                            oCreationPackage.UniqueID = mnu_CPRL_O
                            oCreationPackage.String = "Promotion List"
                            oCreationPackage.Enabled = True
                            oMenus = oMenuItem.SubMenus
                            oMenus.AddEx(oCreationPackage)
                        End If

                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try
                Else
                    intSelectedMatrixrow = eventInfo.Row
                    If oMenuItem.SubMenus.Exists(mnu_CPRL_O) Then
                        oApplication.SBO_Application.Menus.RemoveEx(mnu_CPRL_O)
                    End If
                End If
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

    Private Sub initializeControls(ByVal oForm As SAPbouiCOM.Form)
        Try
            oApplication.Utilities.AddControls(oForm, "_2", "2", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "RIGHT", 0, 0, "2", "Apply Promotion", 0, 0, 0, False)
            oApplication.Utilities.AddControls(oForm, "2_", "_2", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "RIGHT", 0, 0, "_2", "Clear Promotion", 0, 0, 0, False)
            oForm.Items.Item("_2").Width = "140"
            oForm.Items.Item("_2").Height = oForm.Items.Item("1").Height
            oForm.Items.Item("_2").Left = oForm.Items.Item("2").Left + oForm.Items.Item("2").Width + 5
            oForm.Items.Item("2_").Left = oForm.Items.Item("_2").Left + oForm.Items.Item("_2").Width + 5
            oForm.Items.Item("2_").Width = "140"
            oForm.Items.Item("2_").Height = oForm.Items.Item("1").Height
        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    Private Function ValidatePromotion(ByVal oForm As SAPbouiCOM.Form) As Boolean
        Try
            oMatrix = oForm.Items.Item("38").Specific
            Dim strItemCode As String
            Dim dblQty As Double
            Dim strCustomer As String
            Dim strDocDate As String
            Dim strDocEntry As String = String.Empty
            Dim strStatus As String = String.Empty
            Dim strUOM As String = String.Empty

            strCustomer = oForm.Items.Item("4").Specific.Value
            strDocDate = oForm.Items.Item("12").Specific.Value

            For intRow As Integer = 1 To oMatrix.RowCount
                strItemCode = oMatrix.Columns.Item("1").Cells().Item(intRow).Specific.value
                dblQty = oMatrix.Columns.Item("11").Cells().Item(intRow).Specific.value
                strUOM = oMatrix.Columns.Item("1470002145").Cells().Item(intRow).Specific.value
                Dim strRef As String = oMatrix.Columns.Item("U_Z_PrRef").Cells().Item(intRow).Specific.value

                strItemCode = oApplication.Utilities.getMatrixValues(oMatrix, "1", intRow)
                Dim oRec As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRec.DoQuery("Select * from OITM where ""ItemCode""='" & strItemCode & "'")
                If (oRec.RecordCount > 0) Then
                    strItemCode = oRec.Fields.Item("ItemCode").Value
                Else
                    strItemCode = ""
                End If
                If strRef <> "" And strItemCode <> "" Then
                    If CheckPromotionExists(oForm, strCustomer, strDocDate, strItemCode, dblQty, intRow, strStatus, strUOM) = False Then
                        Return False
                    End If
                End If
            Next
            Return True
        Catch ex As Exception
            Throw ex
            Return False
        End Try
    End Function

    Private Sub applyPromotion(ByVal oForm As SAPbouiCOM.Form)
        Try
            oMatrix = oForm.Items.Item("38").Specific
            Dim strItemCode As String
            Dim dblQty As Double
            Dim strCustomer As String
            Dim strDocDate As String
            Dim strDocEntry As String = String.Empty
            Dim strStatus As String = String.Empty
            Dim strUOM As String = String.Empty

            strCustomer = oForm.Items.Item("4").Specific.Value
            strDocDate = oForm.Items.Item("12").Specific.Value

            'Delete Promotion Items if Line Status is Open
            Dim intRowCount As Integer = oMatrix.RowCount
            While intRowCount >= 1
                strStatus = oMatrix.Columns.Item("40").Cells().Item(intRowCount).Specific.value
                If strStatus = "O" Then
                    If CType(oMatrix.Columns.Item("U_Z_PrCode").Cells().Item(intRowCount).Specific, SAPbouiCOM.EditText).Value.Trim().Length > 0 Then
                        oMatrix.DeleteRow(intRowCount)
                    End If
                End If
                intRowCount -= 1
            End While

            oForm.Refresh()
            For intRow As Integer = 1 To oMatrix.RowCount - 1
                strItemCode = oMatrix.Columns.Item("1").Cells().Item(intRow).Specific.value
                dblQty = oMatrix.Columns.Item("11").Cells().Item(intRow).Specific.value
                strUOM = oMatrix.Columns.Item("1470002145").Cells().Item(intRow).Specific.value
                Dim strRef As String = oMatrix.Columns.Item("U_Z_PrRef").Cells().Item(intRow).Specific.value

                strItemCode = oApplication.Utilities.getMatrixValues(oMatrix, "1", intRow)
                Dim oRec As SAPbobsCOM.Recordset
                oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRec.DoQuery("Select * from OITM where ""ItemCode""='" & strItemCode & "'")
                If (oRec.RecordCount > 0) Then
                    strItemCode = oRec.Fields.Item("ItemCode").Value
                Else
                    strItemCode = ""
                End If
                If strRef = "" And strItemCode <> "" Then
                    getFreeOfGoods(oForm, strCustomer, strDocDate, strItemCode, dblQty, intRow, strStatus, strUOM)
                End If
            Next

        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    Private Sub clearPromotion(ByVal oForm As SAPbouiCOM.Form)
        Try
            oMatrix = oForm.Items.Item("38").Specific
            Dim strStatus As String = String.Empty
            'Delete Promotion Items if Line Status is Open
            Dim intRowCount As Integer = oMatrix.RowCount
            While intRowCount >= 1
                strStatus = oMatrix.Columns.Item("40").Cells().Item(intRowCount).Specific.value
                If strStatus = "O" Then
                    If CType(oMatrix.Columns.Item("U_Z_PrRef").Cells().Item(intRowCount).Specific, SAPbouiCOM.EditText).Value.Trim().Length > 0 Then
                        If CType(oMatrix.Columns.Item("U_Z_IType").Cells().Item(intRowCount).Specific, SAPbouiCOM.ComboBox).Selected.Value = "R" Then
                            oMatrix.Columns.Item("U_Z_PrCode").Cells().Item(intRowCount).Specific.value = ""
                            oMatrix.Columns.Item("U_Z_PrRef").Cells().Item(intRowCount).Specific.value = ""
                            CType(oMatrix.Columns.Item("U_Z_IType").Cells().Item(intRowCount).Specific, SAPbouiCOM.ComboBox).Select("R", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            CType(oMatrix.Columns.Item("U_Z_PrmApp").Cells().Item(intRowCount).Specific, SAPbouiCOM.ComboBox).Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        ElseIf CType(oMatrix.Columns.Item("U_Z_IType").Cells().Item(intRowCount).Specific, SAPbouiCOM.ComboBox).Selected.Value = "F" Then
                            oMatrix.Columns.Item("1").Cells.Item(intRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            ' oApplication.SBO_Application.ActivateMenuItem(mnu_DELETE_ROW)
                            oMatrix.DeleteRow(intRowCount)
                        End If
                    End If
                End If
                intRowCount -= 1
            End While
            oForm.Refresh()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub



    Private Function CheckPromotionExists(ByVal oForm As SAPbouiCOM.Form, ByVal strCustomer As String, ByVal strDocDate As String, ByVal strItemCode As String, _
  ByRef dblQuantity As Double, ByVal intRow As Integer, ByVal strStatus As String, ByVal strUOM As String) As Boolean
        Try
            oMatrix = oForm.Items.Item("38").Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim dtDelDate As Date = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(oForm, "12"))
            Dim strPromoCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_PrCode", intRow)
            strQuery = " Select T1.""U_Z_OffCode"",T1.""U_Z_OQty"",T1.""U_Z_OUOMGroup"",T1.""U_Z_Dis"",T0.""U_Z_PrCode"",T1.""U_Z_DisType"",T1.""U_Z_ODis""  "
            strQuery += " From "
            strQuery += " ""@Z_OPRM"" T0 "
            strQuery += " JOIN ""@Z_PRM1"" T1 On T0.""DocEntry"" = T1.""DocEntry"" "
            strQuery += " JOIN ""@Z_OCPR"" T2 On T0.""U_Z_PrCode"" = T2.""U_Z_PrCode"" "
            strQuery += " Where  T2.""U_Z_CustCode"" = '" & strCustomer & "' "
            strQuery += " And '" & dtDelDate.ToString("yyyy-MM-dd") & "' Between  T0.""U_Z_EffFrom"" AND T0.""U_Z_EffTo"" "
            strQuery += " And T1.""U_Z_ItmCode"" = '" & strItemCode & "' "
            If strUOM <> "" Then
                strQuery += " And T1.""U_Z_UOMGroup"" = '" & strUOM & "' "
            End If
            strQuery += " And T2.""U_Z_Active"" = 'Y' "
            strQuery += " And T0.""U_Z_Active"" = 'Y' "
            'If strPromoCode <> "" Then
            '    strQuery = strQuery & " And T0.""U_Z_PrCode""='" & strPromoCode & "'"
            'End If
            oRecordSet.DoQuery(strQuery)
            If oRecordSet.RecordCount <= 0 Then
                CType(oMatrix.Columns.Item("U_Z_PrmApp").Cells().Item(intRow).Specific, SAPbouiCOM.ComboBox).Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)
                oCombobox = oMatrix.Columns.Item("U_Z_IType").Cells.Item(intRow).Specific
                If oCombobox.Selected.Value = "R" Then
                    oCombobox = oMatrix.Columns.Item("U_Z_PrmApp").Cells.Item(intRow).Specific
                    If oCombobox.Selected.Value = "Y" Then
                        oApplication.Utilities.Message("Promotion for the item : " & strItemCode & " is not applicable for this customer. Clear the promotion and try again", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    Else
                        Return True
                    End If
                End If
            Else
                Dim strP As String = oRecordSet.Fields.Item("U_Z_PrCode").Value
                If strPromoCode.ToUpper <> strP.ToUpper And strPromoCode <> "" Then
                    oCombobox = oMatrix.Columns.Item("U_Z_IType").Cells.Item(intRow).Specific
                    If oCombobox.Selected.Value = "R" Then
                        oCombobox = oMatrix.Columns.Item("U_Z_PrmApp").Cells.Item(intRow).Specific
                        If oCombobox.Selected.Value = "Y" Then
                            oApplication.Utilities.Message("Promotion for the item : " & strItemCode & " is not applicable for this customer. Clear the promotion and try again", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        Else
                            Return True
                        End If
                    Else
                        oCombobox = oMatrix.Columns.Item("U_Z_PrmApp").Cells.Item(intRow).Specific
                        If oCombobox.Selected.Value = "N" Then
                            oApplication.Utilities.Message("Promotion for the item : " & strItemCode & " is not applicable for this customer. Clear the promotion and try again", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        Else
                            Return True
                        End If
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub getFreeOfGoods(ByVal oForm As SAPbouiCOM.Form, ByVal strCustomer As String, ByVal strDocDate As String, ByVal strItemCode As String, _
  ByRef dblQuantity As Double, ByVal intRow As Integer, ByVal strStatus As String, ByVal strUOM As String)
        Try
            oMatrix = oForm.Items.Item("38").Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim dtDelDate As Date = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(oForm, "12"))
            strQuery = " Select T1.""U_Z_OffCode"",T1.""U_Z_OQty"",T1.""U_Z_OUOMGroup"",T1.""U_Z_Dis"",T0.""U_Z_PrCode"",T1.""U_Z_DisType"",T1.""U_Z_ODis"",T1.""U_Z_Qty""  "
            strQuery += " From "
            strQuery += " ""@Z_OPRM"" T0 "
            strQuery += " JOIN ""@Z_PRM1"" T1 On T0.""DocEntry"" = T1.""DocEntry"" "
            strQuery += " JOIN ""@Z_OCPR"" T2 On T0.""U_Z_PrCode"" = T2.""U_Z_PrCode"" "
            strQuery += " Where T2.""U_Z_CustCode"" = '" & strCustomer & "' "
            strQuery += " And '" & dtDelDate.ToString("yyyy-MM-dd") & "' Between  T0.""U_Z_EffFrom"" AND T0.""U_Z_EffTo"" "
            strQuery += " And T1.""U_Z_ItmCode"" = '" & strItemCode & "' "
            If strUOM <> "" Then
                strQuery += " And T1.""U_Z_UOMGroup"" = '" & strUOM & "' "
            End If
            strQuery += " And T2.""U_Z_Active"" = 'Y' "
            strQuery += " And T0.""U_Z_Active"" = 'Y' "

            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then

                Try
                    Dim strRef As String = String.Empty

                    Dim dblEligibleQty As Double = oRecordSet.Fields.Item("U_Z_Qty").Value
                    Dim dblOfferQty As Double = oRecordSet.Fields.Item("U_Z_OQty").Value
                    If dblQuantity >= dblEligibleQty Then
                        oApplication.Utilities.addPromotionReference(strRef)
                        Dim intOfferQty As Integer = Math.Floor(dblQuantity / dblEligibleQty)
                        dblOfferQty = dblOfferQty * intOfferQty
                        Dim dblOfferDiscount As Double = oRecordSet.Fields.Item("U_Z_ODis").Value
                        'Regular Item
                        CType(oMatrix.Columns.Item("U_Z_PrmApp").Cells().Item(intRow).Specific, SAPbouiCOM.ComboBox).Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        CType(oMatrix.Columns.Item("U_Z_IType").Cells().Item(intRow).Specific, SAPbouiCOM.ComboBox).Select("R", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        oMatrix.Columns.Item("U_Z_PrRef").Cells().Item(intRow).Specific.value = strRef

                        If oRecordSet.Fields.Item("U_Z_DisType").Value = "D" Then
                            oMatrix.Columns.Item("15").Cells().Item(intRow).Specific.value = oRecordSet.Fields.Item("U_Z_Dis").Value
                        Else
                            oMatrix.AddRow(1, oMatrix.RowCount)
                            'Free Item
                            oMatrix.Columns.Item("1").Cells().Item(oMatrix.RowCount - 1).Specific.value = oRecordSet.Fields.Item("U_Z_OffCode").Value
                            oMatrix.Columns.Item("11").Cells().Item(oMatrix.RowCount - 1).Specific.value = dblOfferQty ' oRecordSet.Fields.Item("U_Z_OQty").Value
                            Try
                                oMatrix.Columns.Item("1470002145").Cells().Item(oMatrix.RowCount - 1).Specific.value = oRecordSet.Fields.Item("U_Z_OUOMGroup").Value
                            Catch ex As Exception
                            End Try
                            'oApplication.Utilities.SetMatrixValues(oMatrix, "15", oMatrix.RowCount - 1, dblOfferDiscount)
                            oMatrix.Columns.Item("15").Cells().Item(oMatrix.RowCount - 1).Specific.string = dblOfferDiscount 'oRecordSet.Fields.Item("U_Z_ODis").Value
                            oMatrix.Columns.Item("U_Z_PrCode").Cells().Item(oMatrix.RowCount - 1).Specific.value = oRecordSet.Fields.Item("U_Z_PrCode").Value
                            oMatrix.Columns.Item("U_Z_PrRef").Cells().Item(oMatrix.RowCount - 1).Specific.value = strRef
                            CType(oMatrix.Columns.Item("U_Z_IType").Cells().Item(oMatrix.RowCount - 1).Specific, SAPbouiCOM.ComboBox).Select("F", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        End If
                    End If

                Catch ex As Exception

                End Try
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    

End Class
