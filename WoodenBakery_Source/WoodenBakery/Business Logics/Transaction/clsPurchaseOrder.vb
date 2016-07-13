Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsPurchaseOrder
    Inherits clsBase
    Private oEditText As SAPbouiCOM.EditText
    Private oRecordSet As SAPbobsCOM.Recordset
    Private strQuery As String
    Private oMatrix As SAPbouiCOM.Matrix
    Private oHTList As Hashtable

    Public Sub New()
        MyBase.New()
    End Sub


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

    Private Sub getFreeOfGoods(ByVal oForm As SAPbouiCOM.Form, ByVal strCustomer As String, ByVal strDocDate As String, ByVal strItemCode As String, _
  ByRef dblQuantity As Double, ByVal intRow As Integer, ByVal strStatus As String, ByVal strUOM As String)
        Try
            oMatrix = oForm.Items.Item("38").Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim dtDelDate As Date = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(oForm, "12"))
            strQuery = " Select T1.""U_Z_OffCode"",T1.""U_Z_OQty"",T1.""U_Z_OUOMGroup"",T1.""U_Z_Dis"",T0.""U_Z_PrCode"",T1.""U_Z_DisType"",T1.""U_Z_ODis"",T0.""U_Z_Qty""  "
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
                        Dim intOfferQty As Integer = Math.Floor(dblQuantity / dblEligibleQty)
                        dblOfferQty = dblOfferQty * intOfferQty
                        oApplication.Utilities.addPromotionReference(strRef)

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
                            oMatrix.Columns.Item("15").Cells().Item(oMatrix.RowCount - 1).Specific.value = oRecordSet.Fields.Item("U_Z_ODis").Value
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

    Private Function CheckPromotionExists(ByVal oForm As SAPbouiCOM.Form, ByVal strCustomer As String, ByVal strDocDate As String, ByVal strItemCode As String, _
  ByRef dblQuantity As Double, ByVal intRow As Integer, ByVal strStatus As String, ByVal strUOM As String) As Boolean
        Try
            Dim oCombobox As SAPbouiCOM.ComboBox
            oMatrix = oForm.Items.Item("38").Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strPromoCode As String = oApplication.Utilities.getMatrixValues(oMatrix, "U_Z_PrCode", intRow)
            Dim dtDelDate As Date = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(oForm, "12"))
            strQuery = " Select T1.""U_Z_OffCode"",T1.""U_Z_OQty"",T1.""U_Z_OUOMGroup"",T1.""U_Z_Dis"",T0.""U_Z_PrCode"",T1.""U_Z_DisType"",T1.""U_Z_ODis""  "
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
            If oRecordSet.RecordCount <= 0 Then
                CType(oMatrix.Columns.Item("U_Z_PrmApp").Cells().Item(intRow).Specific, SAPbouiCOM.ComboBox).Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)

                oCombobox = oMatrix.Columns.Item("U_Z_IType").Cells.Item(intRow).Specific
                If oCombobox.Selected.Value = "R" Then


                    oCombobox = oMatrix.Columns.Item("U_Z_PrmApp").Cells.Item(intRow).Specific
                    If oCombobox.Selected.Value = "Y" Then
                        oApplication.Utilities.Message("Promotion for the item : " & strItemCode & " is not applicable for this Supplier. Clear the promotion and try again", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
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
                    End If
                End If
            End If

            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#Region "Menu Event"

    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            Select Case pVal.MenuUID
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                Case mnu_ADD
                    oForm.Items.Item("38").Enabled = True
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

#Region "Item Event"

    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_PurchaseOrder Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "38" And pVal.ColUID = "14" Then
                                    oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                    If oApplication.Utilities.ValidateItemIdentifier(oApplication.Utilities.getMatrixValues(oMatrix, "1", pVal.Row)) = False Then
                                        BubbleEvent = False
                                        Exit Sub

                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "38" And pVal.ColUID = "14" And pVal.CharPressed <> 9 Then
                                    oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                    If oApplication.Utilities.ValidateItemIdentifier(oApplication.Utilities.getMatrixValues(oMatrix, "1", pVal.Row)) = False Then
                                        BubbleEvent = False
                                        Exit Sub

                                    End If
                                End If

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
                            Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If Not oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    If (pVal.ItemUID = "4" Or pVal.ItemUID = "46") And pVal.CharPressed = 9 Then
                                        If oForm.PaneLevel = 1 Then
                                            oForm.Freeze(True)
                                            changePrice(oForm)
                                            oForm.Freeze(False)
                                        End If
                                    ElseIf pVal.ItemUID = "38" And (pVal.ColUID = "1" Or pVal.ColUID = "3") Then 'pVal.ColUID = "3" Then
                                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            If Not IsNothing(oHTList) Then
                                                Dim key As ICollection = oHTList.Keys
                                                Dim k As DictionaryEntry
                                                Dim oDataView As DataView = SortHashtable(oHTList)
                                                For iRow As Long = 0 To oDataView.Count - 1
                                                    Dim sKey As String = oDataView(iRow)("key")
                                                    Dim sValue As String = oDataView(iRow)("value")
                                                    oForm.Freeze(True)
                                                    ' fillRProjectByRow(oForm, CInt(sKey))
                                                    changePrice(oForm, CInt(sKey))
                                                    oForm.Freeze(False)
                                                Next
                                                oHTList = Nothing
                                                'Dim key As ICollection = oHTList.Keys
                                                'Dim k As DictionaryEntry
                                                'For Each k In oHTList
                                                '    oForm.Freeze(True)
                                                '    fillRProjectByRow(oForm, CInt(k.Key))
                                                '    changePrice(oForm, CInt(k.Key))
                                                '    oForm.Freeze(False)
                                                'Next k
                                                'oHTList = Nothing
                                            End If
                                        End If
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                If pVal.ItemUID = "38" And (pVal.ColUID = "1" Or pVal.ColUID = "3") Then 'pVal.ColUID = "3" And pVal.Row > 0 Then
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                        If pVal.CharPressed = 9 Then
                                            Try
                                                changePrice(oForm, pVal.Row)
                                            Catch ex As Exception
                                                oForm.Freeze(False)
                                            End Try
                                        End If
                                    End If
                                End If
                                'Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                '    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                '    Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                '    Dim oCFL As SAPbouiCOM.ChooseFromList
                                '    Dim sCHFL_ID, val As String
                                '    Try
                                '        oCFLEvento = pVal
                                '        sCHFL_ID = oCFLEvento.ChooseFromListUID
                                '        oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                '        Dim oDataTable As SAPbouiCOM.DataTable
                                '        oDataTable = oCFLEvento.SelectedObjects
                                '        If pVal.ColUID = "31" And pVal.ItemUID = "38" Then
                                '            If IsNothing(oCFLEvento.SelectedObjects) Then
                                '                val = ""
                                '            Else
                                '                val = oDataTable.GetValue("PrjCode", 0)
                                '            End If
                                '            oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                '            If val = "" Then
                                '                Try
                                '                    'oApplication.Utilities.SetMatrixValues(oMatrix, "U_SPDocEty", pVal.Row, "")
                                '                    oApplication.Utilities.SetMatrixValues(oMatrix, "15", pVal.Row, 0)
                                '                    oMatrix.SetCellWithoutValidation(pVal.Row, "U_SPDocEty", "")
                                '                    oMatrix.SetCellWithoutValidation(pVal.Row, "U_PriceType", "S")
                                '                Catch ex As Exception
                                '                    'oApplication.Utilities.SetMatrixValues(oMatrix, "U_SPDocEty", pVal.Row, "")
                                '                    oApplication.Utilities.SetMatrixValues(oMatrix, "15", pVal.Row, 0)
                                '                    oMatrix.SetCellWithoutValidation(pVal.Row, "U_SPDocEty", "")
                                '                    oMatrix.SetCellWithoutValidation(pVal.Row, "U_PriceType", "S")
                                '                End Try
                                '            Else
                                '                If oCFL.ObjectType = "63" Then
                                '                    changePrice(oForm, pVal.Row, val)
                                '                End If
                                '            End If
                                '            If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                '                If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                '                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                '                End If
                                '            End If
                                '        ElseIf pVal.ItemUID = "38" And (pVal.ColUID = "1" Or pVal.ColUID = "3") Then
                                '            If Not IsNothing(oDataTable) Then
                                '                oHTList = New Hashtable(oDataTable.Rows.Count)
                                '                For index As Integer = 0 To oDataTable.Rows.Count - 1
                                '                    oHTList.Add((pVal.Row + index), oDataTable.GetValue("ItemCode", index))
                                '                Next
                                '            End If
                                '        End If
                                '    Catch ex As Exception
                                '        oForm.Freeze(False)
                                '    End Try
                        End Select
                End Select
            End If
        Catch ex As Exception
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

#Region "Data Events"

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
            Select Case BusinessObjectInfo.BeforeAction
                Case True

                Case False
                    Select Case BusinessObjectInfo.EventType
                        Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                            'oForm.Items.Item("_2").Enabled = False
                    End Select
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

#Region "Right Click"

    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
            If oForm.TypeEx = frm_PurchaseOrder Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                If (eventInfo.BeforeAction = True) Then
                    Try

                    Catch ex As Exception
                        MessageBox.Show(ex.Message)
                    End Try
                Else

                End If
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

#Region "Function"

    'Private Sub initializeControls(ByVal oForm As SAPbouiCOM.Form)
    '    Try

    '        oForm.Items.Item("156").Left = oForm.Items.Item("70").Left
    '        oForm.Items.Item("156").Top = oForm.Items.Item("70").Top + oForm.Items.Item("70").Height + 1
    '        oForm.Items.Item("157").Left = oForm.Items.Item("63").Left
    '        oForm.Items.Item("157").Top = oForm.Items.Item("63").Top + oForm.Items.Item("63").Height + 1

    '        oForm.Items.Item("156").FromPane = 0
    '        oForm.Items.Item("156").ToPane = 7

    '        oForm.Items.Item("157").FromPane = 0
    '        oForm.Items.Item("157").ToPane = 7

    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    Private Sub changePrice(ByVal oForm As SAPbouiCOM.Form)
        Try
            oMatrix = oForm.Items.Item("38").Specific
            Dim strItemCode, strCustomer, strpriceCurr As String
            Dim dblPrice As Double = 0
            For intRow As Integer = 1 To oMatrix.RowCount
                strItemCode = oMatrix.Columns.Item("1").Cells().Item(intRow).Specific.value
                strCustomer = oForm.Items.Item("4").Specific.Value
                getSpecialPrice(oForm, strCustomer, strItemCode, strpriceCurr, dblPrice)
                If dblPrice <> 0 Then
                    'Dim dblprice1 As Double = calculateUnitPrice(dblDiscount, oApplication.Utilities.getDocumentQuantity(strPrice))
                    oMatrix.Columns.Item("14").Cells().Item(intRow).Specific.value = (strpriceCurr & " " & dblPrice).ToString 'dblPrice
                End If
            Next

        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub changePrice(ByVal oForm As SAPbouiCOM.Form, ByVal intRow As Integer)
        Try
            oForm.Freeze(True)
            oMatrix = oForm.Items.Item("38").Specific
            Dim strItemCode, strCustomer, strPrice, strpriceCurr As String
            Dim dblPrice As Double
            
            strItemCode = oMatrix.Columns.Item("1").Cells().Item(intRow).Specific.value
            strCustomer = oForm.Items.Item("4").Specific.Value
            getSpecialPrice(oForm, strCustomer, strItemCode, strpriceCurr, dblPrice)
            If dblPrice <> 0 Then
                'Dim dblprice1 As Double = calculateUnitPrice(dblDiscount, oApplication.Utilities.getDocumentQuantity(strPrice))
                oMatrix.Columns.Item("14").Cells().Item(intRow).Specific.value = (strpriceCurr & " " & dblPrice).ToString 'dblPrice
            End If

            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub getSpecialPrice(ByVal oForm As SAPbouiCOM.Form, ByVal strCustomer As String, ByVal strItemCode As String, _
                                ByRef strPriceCurr As String, ByRef dblUnitPrice As Double)
        Try
            Dim oSPRecordSet As SAPbobsCOM.Recordset
            oSPRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select ""U_Z_CPrice"",""U_Z_CCurrency"" from OSCN where ""ItemCode""='" & strItemCode & "'"
            strQuery &= "and ""CardCode""='" & strCustomer & "'"
            oSPRecordSet.DoQuery(strQuery)
            If Not oSPRecordSet.EoF Then
                strPriceCurr = oSPRecordSet.Fields.Item("U_Z_CCurrency").Value
                dblUnitPrice = oSPRecordSet.Fields.Item("U_Z_CPrice").Value
            Else
                strPriceCurr = ""
                dblUnitPrice = 0
            End If
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Function SortHashtable(ByVal oHash As Hashtable) As DataView
        Dim oTable As New Data.DataTable
        oTable.Columns.Add(New Data.DataColumn("key"))
        oTable.Columns.Add(New Data.DataColumn("value"))

        For Each oEntry As Collections.DictionaryEntry In oHash
            Dim oDataRow As DataRow = oTable.NewRow()
            oDataRow("key") = oEntry.Key
            oDataRow("value") = oEntry.Value
            oTable.Rows.Add(oDataRow)
        Next

        Dim oDataView As DataView = New DataView(oTable)
        oDataView.Sort = "key ASC "

        Return oDataView
    End Function

    Private Function calculateUnitPrice(ByVal aDiscount As Double, ByVal aPrice As Double) As Double
        Dim dblTemp As Double
        Dim dblUnitprice As Double
        If aPrice = 0 Then
            Return 0
        End If
        dblTemp = aDiscount / 100
        dblTemp = 1 - dblTemp
        dblUnitprice = aPrice / dblTemp
        Return dblUnitprice
    End Function

#End Region

End Class
