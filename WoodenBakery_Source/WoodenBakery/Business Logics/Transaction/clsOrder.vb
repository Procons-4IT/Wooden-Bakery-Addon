Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsOrder
    Inherits clsBase
    Private oEditText As SAPbouiCOM.EditText
    Private oRecordSet As SAPbobsCOM.Recordset
    Private strQuery As String
    Private oMatrix As SAPbouiCOM.Matrix

    Public Sub New()
        MyBase.New()
    End Sub

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            Select Case pVal.MenuUID
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                Case mnu_ADD
                    oForm.Items.Item("38").Enabled = True
                    oForm.Items.Item("_2").Enabled = True
                Case mnu_CPRL_O
                    If Not oForm.Items.Item("4").Specific.value = "" Then
                        Dim objPromList As clsCustPromotionList
                        objPromList = New clsCustPromotionList
                        objPromList.LoadForm(oForm.Items.Item("4").Specific.value)
                    Else
                        oApplication.Utilities.Message("Select Customer to Get Promotion List...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    End If
                    Dim oMenuItem As SAPbouiCOM.MenuItem
                    oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                    If oMenuItem.SubMenus.Exists(pVal.MenuUID) Then
                        oApplication.SBO_Application.Menus.RemoveEx(pVal.MenuUID)
                    End If
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_SalesOrder Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_CLICK, SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("38").Specific
                                If pVal.ItemUID = "38" And (pVal.ColUID = "U_PrmApp" Or pVal.ColUID = "U_PrCode" Or pVal.ColUID = "U_SPDocEty" Or pVal.ColUID = "U_PrRef" Or pVal.ColUID = "U_PrLine") Then
                                    BubbleEvent = False
                                    Exit Sub
                                ElseIf pVal.ItemUID = "38" And pVal.Row > 0 Then
                                    If CType(oMatrix.Columns.Item("U_PrmApp").Cells().Item(pVal.Row).Specific, SAPbouiCOM.ComboBox).Value = "Y" Then
                                        If pVal.ColUID = "1" Or pVal.ColUID = "2" Or pVal.ColUID = "3" Or pVal.ColUID = "11" Then
                                            oApplication.Utilities.Message("Promotion details already applied for this Row. You should delete the  .", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                                            BubbleEvent = False
                                        End If
                                    ElseIf (CType(oMatrix.Columns.Item("U_PrCode").Cells().Item(pVal.Row).Specific, SAPbouiCOM.EditText).Value.Trim().Length > 0) Then
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
                                If pVal.ItemUID = "_2" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
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
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                initializeControls(oForm)
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
            If oForm.TypeEx = frm_SalesOrder Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                If (eventInfo.BeforeAction = True) Then
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

#Region "Function"

    Private Sub initializeControls(ByVal oForm As SAPbouiCOM.Form)
        Try
            oApplication.Utilities.AddControls(oForm, "_2", "2", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "RIGHT", 0, 0, "2", "Apply Promotion", 0, 0, 0, False)
            oForm.Items.Item("_2").Width = "140"
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
                    If CType(oMatrix.Columns.Item("U_PrCode").Cells().Item(intRowCount).Specific, SAPbouiCOM.EditText).Value.Trim().Length > 0 Then
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
                getFreeOfGoods(oForm, strCustomer, strDocDate, strItemCode, dblQty, intRow, strStatus, strUOM)
            Next

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub getFreeOfGoods(ByVal oForm As SAPbouiCOM.Form, ByVal strCustomer As String, ByVal strDocDate As String, ByVal strItemCode As String, _
  ByRef dblQuantity As Double, ByVal intRow As Integer, ByVal strStatus As String, ByVal strUOM As String)
        Try
            oMatrix = oForm.Items.Item("38").Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = " Select T1.""U_Z_OffCode"",T1.""U_Z_OQty"",T1.""U_Z_OUOMGroup"",T1.""U_Z_ODis"",T0.""U_PrCode"" "
            strQuery += " From "
            strQuery += " ""@Z_OPRM"" T0 "
            strQuery += " JOIN ""@Z_PRM1"" T1 On T0.""DocEntry"" = T1.""DocEntry"" "
            strQuery += " JOIN ""@Z_OCPR"" T2 On T0.""U_Z_PrCode"" = T1.""U_Z_PrCode"" "
            strQuery += " Where T2.""U_Z_CustCode"" = '" & strCustomer & "' "
            strQuery += " And T1.""U_Z_ItmCode"" = '" & strItemCode & "' "
            strQuery += " And T1.""U_Z_UOMGroup"" = '" & strUOM & "' "
            strQuery += " And T2.""U_Z_Active"" = 'Y' "

            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                oMatrix.AddRow(1, oMatrix.RowCount)

                Try
                    Dim strRef As String = String.Empty
                    oApplication.Utilities.addPromotionReference(strRef)

                    'Regular Item
                    CType(oMatrix.Columns.Item("U_PrmApp").Cells().Item(intRow).Specific, SAPbouiCOM.ComboBox).Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    CType(oMatrix.Columns.Item("U_IType").Cells().Item(intRow).Specific, SAPbouiCOM.ComboBox).Select("R", SAPbouiCOM.BoSearchKey.psk_ByValue)
                    oMatrix.Columns.Item("U_PrRef").Cells().Item(intRow).Specific.value = strRef

                    'Free Item
                    oMatrix.Columns.Item("1").Cells().Item(oMatrix.RowCount - 1).Specific.value = oRecordSet.Fields.Item("U_OffCode").Value
                    oMatrix.Columns.Item("11").Cells().Item(oMatrix.RowCount - 1).Specific.value = oRecordSet.Fields.Item("U_OQty").Value
                    oMatrix.Columns.Item("15").Cells().Item(oMatrix.RowCount - 1).Specific.value = oRecordSet.Fields.Item("U_ODis").Value
                    oMatrix.Columns.Item("1470002145").Cells().Item(oMatrix.RowCount - 1).Specific.value = oRecordSet.Fields.Item("U_Z_OUOMGroup").Value

                    oMatrix.Columns.Item("U_PrCode").Cells().Item(oMatrix.RowCount - 1).Specific.value = oRecordSet.Fields.Item("U_PrCode").Value
                    oMatrix.Columns.Item("U_PrRef").Cells().Item(oMatrix.RowCount - 1).Specific.value = strRef
                    CType(oMatrix.Columns.Item("U_IType").Cells().Item(oMatrix.RowCount - 1).Specific, SAPbouiCOM.ComboBox).Select("F", SAPbouiCOM.BoSearchKey.psk_ByValue)

                Catch ex As Exception

                End Try
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

End Class
