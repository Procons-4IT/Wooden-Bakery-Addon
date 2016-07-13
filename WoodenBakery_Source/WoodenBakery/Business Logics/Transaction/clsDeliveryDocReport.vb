Imports System.IO
Public Class clsDeliveryDocReport
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
    Private oComoColumn As SAPbouiCOM.ComboBoxColumn

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Private Sub LoadForm()
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_SOClosingRepot) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_Z_ODEL_R, frm_Z_ODEL_R)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        AddChooseFromList(oForm)
        oCombobox = oForm.Items.Item("7").Specific
        oForm.DataSources.UserDataSources.Add("IsSign", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oCombobox.DataBind.SetBound(True, "", "IsSign")
        oCombobox.ValidValues.Add("Y", "Yes")
        oCombobox.ValidValues.Add("N", "No")
        oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        oForm.Items.Item("7").DisplayDesc = True
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

        oCombobox = oForm.Items.Item("21").Specific
        oForm.DataSources.UserDataSources.Add("DocType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oCombobox.DataBind.SetBound(True, "", "DocType")
        oCombobox.ValidValues.Add("OINV", "Invoice")
        oCombobox.ValidValues.Add("ORIN", "Credit Note")
        oCombobox.ValidValues.Add("ORCT", "Incoming Payment")
        oCombobox.ValidValues.Add("OVPM", "Outgoing Payment")
        oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        oForm.Items.Item("21").DisplayDesc = True
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

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


            oCFLCreationParams.MultiSelection = False
            oCFL = oCFLs.Item("CFL_2")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

            oCFL = oCFLs.Item("CFL_3")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)
            oCon = oCons.Add()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#Region "Data Bind"
    Private Sub DataBind(aForm As SAPbouiCOM.Form)
        Dim strIsSigned, strReasonCode, strDocDate, strDocDate1, strCardcode, strCardCode1, strItemCode, strItemCode1, strCondition, strSQL, strItemGroup, strDocType, strDocument As String
        Dim dtDate, dtDate1 As Date
        Try
            aForm.Freeze(True)

            
            oCombobox = aForm.Items.Item("7").Specific
            strIsSigned = oCombobox.Selected.Value

            oCombobox = aForm.Items.Item("21").Specific
            strDocType = oCombobox.Selected.Value
            strDocument = oCombobox.Selected.Description
            strDocDate = oApplication.Utilities.getEdittextvalue(aForm, "11")
            strDocDate1 = oApplication.Utilities.getEdittextvalue(aForm, "13")
            strCardcode = oApplication.Utilities.getEdittextvalue(aForm, "15")
            strCardCode1 = oApplication.Utilities.getEdittextvalue(aForm, "17")


            '  strSQL = "SELECT  T0.""DocEntry"", T0.""DocNum"",  T0.""DocDate"", T0.""CardCode"", T0.""CardName"" , T0.""U_Z_DelDate"", T0.""U_Z_IsDel"" FROM OINV T0  "
            strSQL = "SELECT '" & strDocument & "' as ""Document Type"" , T0.""DocEntry"", T0.""DocNum"",  T0.""DocDate"", T0.""CardCode"", T0.""CardName"" ,T1.""U_Driver"" ""Driver Name"", T0.""U_Z_DelDate"",T0.""U_Z_ScnUser"", T0.""U_Z_IsDel"", T1.""U_WhseCode"" FROM " & strDocType & " T0 Left Outer Join OCRD T1 on T1.""CardCode""=T0.""CardCode"" "

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


            If strDocDate <> "" Then
                dtDate = oApplication.Utilities.GetDateTimeValue(strDocDate)
                strCondition = strCondition & " and ( T0.""U_Z_DelDate"">='" & dtDate.ToString("yyyy-MM-dd") & "'"
            Else
                strCondition = strCondition & " and ( 1=1 "
            End If
            If strDocDate1 <> "" Then
                dtDate1 = oApplication.Utilities.GetDateTimeValue(strDocDate1)
                strCondition = strCondition & " and T0.""U_Z_DelDate""<='" & dtDate1.ToString("yyyy-MM-dd") & "')"
            Else
                strCondition = strCondition & " and 1=1) "
            End If


            strDocDate = oApplication.Utilities.getEdittextvalue(aForm, "31")
            strDocDate1 = oApplication.Utilities.getEdittextvalue(aForm, "32")

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

            If strIsSigned = "Y" Then
                If blnIsHanaDB Then
                    strCondition = strCondition & " and  IFNULL(T0.""U_Z_IsDel"",'N')='Y'"
                Else
                    strCondition = strCondition & " and  ISNULL(T0.""U_Z_IsDel"",'N')='Y'"
                End If
            Else
                If blnIsHanaDB Then
                    strCondition = strCondition & " and  IFNULL(T0.""U_Z_IsDel"",'N')='N'"
                Else
                    strCondition = strCondition & " and  ISNULL(T0.""U_Z_IsDel"",'N')='N'"
                End If
            End If

            strSQL = strSQL & " Where  1 = 1 " & strCondition

            oGrid = aForm.Items.Item("22").Specific
            oGrid.DataTable.ExecuteQuery(strSQL)
            For intR As Integer = 0 To oGrid.Columns.Count - 1
                oGrid.Columns.Item(intR).Editable = False
            Next

            oEditTextColumn = oGrid.Columns.Item("DocEntry")
            oEditTextColumn.LinkedObjectType = "13"
            Select Case strDocType
                Case "OINV"
                    oEditTextColumn.LinkedObjectType = "13"
                Case "ORIN"
                    oEditTextColumn.LinkedObjectType = "14"
                Case "ORCT"
                    oEditTextColumn.LinkedObjectType = "24"
                Case "OVPM"
                    oEditTextColumn.LinkedObjectType = "46"
            End Select


            oEditTextColumn = oGrid.Columns.Item("CardCode")
            oEditTextColumn.LinkedObjectType = "2"
            oGrid.Columns.Item("U_Z_DelDate").TitleObject.Caption = "Signed Date"
            oGrid.Columns.Item("U_Z_IsDel").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oComoColumn = oGrid.Columns.Item("U_Z_IsDel")
            oComoColumn.ValidValues.Add("Y", "Yes")
            oComoColumn.ValidValues.Add("N", "No")
            oGrid.Columns.Item("U_Z_IsDel").Editable = False
            oComoColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
            oGrid.Columns.Item("U_Z_IsDel").TitleObject.Caption = "Is Signed"
            oGrid.Columns.Item("U_Z_ScnUser").TitleObject.Caption = "Scanned User"
            oGrid.Columns.Item("U_Z_ScnUser").Editable = False
            oGrid.Columns.Item("U_WhseCode").TitleObject.Caption = "Ware House"
            oGrid.AutoResizeColumns()
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None

            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Z_ODEL_R Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "3"
                                        oForm.PaneLevel = oForm.PaneLevel - 1
                                    Case "4"
                                        oForm.PaneLevel = oForm.PaneLevel + 1
                                        '  oForm.PaneLevel = oForm.PaneLevel + 1
                                        If oForm.PaneLevel = 3 Then
                                            DataBind(oForm)
                                        End If
                                        'Case "5"
                                        '    If oApplication.SBO_Application.MessageBox("Do you want to export the selected records into Tab Delimted file...", , "Continue", "Cancel") = 2 Then
                                        '        Exit Sub
                                        '    End If
                                        '    GenerateFile(oForm)
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
                                        If oCFL.ObjectType = "2" Then
                                            val = oDataTable.GetValue("CardCode", 0)
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
                Case mnu_Z_ODEL_R
                     LoadForm()
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
End Class
