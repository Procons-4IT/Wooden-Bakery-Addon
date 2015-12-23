Public Class clsMasters
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

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Private Function AddControls(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Try
            aForm.Freeze(True)
            If aForm.TypeEx = frm_BPMaster Then
                oApplication.Utilities.AddControls(aForm, "_1001", "2013", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", , , , "WeekEnd")
                oApplication.Utilities.AddControls(aForm, "_1002", "2014", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN", , , "_1001")

                oApplication.Utilities.AddControls(aForm, "_1003", "_1001", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", , , , "BP Category")
                oApplication.Utilities.AddControls(aForm, "_1004", "_1002", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN", , , "_1003")
               
            ElseIf aForm.TypeEx = frm_ItemMaster Then
                oApplication.Utilities.AddControls(aForm, "_1001", "1470002292", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", 6, 6, , "DeliveryDays")
                oApplication.Utilities.AddControls(aForm, "_1002", "1470002293", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN", 6, 6, "_1001")

                oApplication.Utilities.AddControls(aForm, "_1003", "_1001", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", 6, 6, , "Item Category")
                oApplication.Utilities.AddControls(aForm, "_1004", "_1002", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN", 6, 6, "_1003")

            ElseIf aForm.TypeEx = frm_Warehouse Then
                oApplication.Utilities.AddControls(aForm, "_1003", "42", SAPbouiCOM.BoFormItemTypes.it_STATIC, "DOWN", 1, 1, , "Category")
                oApplication.Utilities.AddControls(aForm, "_1004", "41", SAPbouiCOM.BoFormItemTypes.it_EDIT, "DOWN", 1, 1, "_1003")

            End If
            '    oApplication.Utilities.AddControls(aForm, "BtnAuto", "2", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "RIGHT", 0, 0, , "Auto Selection(FIFO)", 150)
            ' aForm.Items.Item("16").Enabled = False
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
        End Try
    End Function

    Private Sub AddChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition


            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)
            If objForm.TypeEx = frm_BPMaster Then
                oCFLCreationParams.MultiSelection = False
                oCFLCreationParams.ObjectType = "Z_OWEM"
                oCFLCreationParams.UniqueID = "CFL1"
                oCFL = oCFLs.Add(oCFLCreationParams)
                oCons = oCFL.GetConditions()
                oCon = oCons.Add()
                oCon.Alias = "U_Z_Active"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = "Y"
                oCFL.SetConditions(oCons)
                oCon = oCons.Add()

                oEditText = oForm.Items.Item("_1002").Specific
                oEditText.DataBind.SetBound(True, "OCRD", "U_Z_WeekEnd")
                oEditText.ChooseFromListUID = "CFL1"
                oEditText.ChooseFromListAlias = "U_Z_Code"


                oCFLCreationParams.MultiSelection = False
                oCFLCreationParams.ObjectType = "Z_OBPC"
                oCFLCreationParams.UniqueID = "CFL11"
                oCFL = oCFLs.Add(oCFLCreationParams)
                oCons = oCFL.GetConditions()
                oCon = oCons.Add()
                oCon.Alias = "U_Z_Active"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = "Y"
                oCFL.SetConditions(oCons)
                oCon = oCons.Add()

                oEditText = oForm.Items.Item("_1004").Specific
                oEditText.DataBind.SetBound(True, "OCRD", "U_Z_BPCCODE")
                oEditText.ChooseFromListUID = "CFL11"
                oEditText.ChooseFromListAlias = "U_Z_Code"


            ElseIf objForm.TypeEx = frm_ItemMaster Then
                'oCFLCreationParams.MultiSelection = False
                'oCFLCreationParams.ObjectType = "Z_HR_OBUOB"
                'oCFLCreationParams.UniqueID = "CFL1"
                'oCFL = oCFLs.Add(oCFLCreationParams)
                oEditText = oForm.Items.Item("_1002").Specific
                oEditText.DataBind.SetBound(True, "OITM", "U_Z_DelDays")


                oCFLCreationParams.MultiSelection = False
                oCFLCreationParams.ObjectType = "Z_OITC"
                oCFLCreationParams.UniqueID = "CFL11"
                oCFL = oCFLs.Add(oCFLCreationParams)
                oCons = oCFL.GetConditions()
                oCon = oCons.Add()
                oCon.Alias = "U_Z_Active"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = "Y"
                oCFL.SetConditions(oCons)
                oCon = oCons.Add()

                oEditText = oForm.Items.Item("_1004").Specific
                oEditText.DataBind.SetBound(True, "OITM", "U_Z_ITCCODE")
                oEditText.ChooseFromListUID = "CFL11"
                oEditText.ChooseFromListAlias = "U_Z_Code"
            ElseIf objForm.TypeEx = frm_Warehouse Then
                oCFLCreationParams.MultiSelection = False
                oCFLCreationParams.ObjectType = "Z_OWHC"
                oCFLCreationParams.UniqueID = "CFL11"
                oCFL = oCFLs.Add(oCFLCreationParams)
                oCons = oCFL.GetConditions()
                oCon = oCons.Add()
                oCon.Alias = "U_Z_Active"
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCon.CondVal = "Y"
                oCFL.SetConditions(oCons)
                oCon = oCons.Add()
                oEditText = oForm.Items.Item("_1004").Specific
                oEditText.DataBind.SetBound(True, "OWHS", "U_Z_WHSCODE")
                oEditText.ChooseFromListUID = "CFL11"
                oEditText.ChooseFromListAlias = "U_Z_Code"
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub DataBind(aform As SAPbouiCOM.Form)
        AddChooseFromList(aform)
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_BPMaster Or pVal.FormTypeEx = frm_ItemMaster Or pVal.FormTypeEx = frm_Warehouse Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                'If pVal.ItemUID = "_1004" Then
                                '    Dim oRec As SAPbobsCOM.Recordset
                                '    oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                '    If oApplication.Utilities.getEdittextvalue(oForm, pVal.ItemUID) <> "" Then
                                '        If oForm.TypeEx = frm_BPMaster Then
                                '            oRec.DoQuery("Select * from ""@Z_LUSR3"" where ""U_Z_Code""='" & oApplication.Utilities.getEdittextvalue(oForm, pVal.ItemUID) & "'")
                                '            If oRec.RecordCount > 0 Then
                                '                oApplication.Utilities.Message(" Category already mapped to user.", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                '                BubbleEvent = False
                                '                Exit Sub
                                '            End If
                                '        ElseIf oForm.TypeEx = frm_ItemMaster Then
                                '            oRec.DoQuery("Select * from ""@Z_LUSR2"" where ""U_Z_Code""='" & oApplication.Utilities.getEdittextvalue(oForm, pVal.ItemUID) & "'")
                                '            If oRec.RecordCount > 0 Then
                                '                oApplication.Utilities.Message(" Category already mapped to user.", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                '                BubbleEvent = False
                                '                Exit Sub
                                '            End If
                                '        ElseIf oForm.TypeEx = frm_Warehouse Then
                                '            oRec.DoQuery("Select * from ""@Z_LUSR4"" where ""U_Z_Code""='" & oApplication.Utilities.getEdittextvalue(oForm, pVal.ItemUID) & "'")
                                '            If oRec.RecordCount > 0 Then
                                '                oApplication.Utilities.Message("Category already mapped to user.", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                '                BubbleEvent = False
                                '                Exit Sub
                                '            End If
                                '        Else

                                '        End If
                                '    End If
                                'End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                AddControls(oForm)
                                DataBind(oForm)
                                ' AssignBatchNumber(oForm)
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
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
                                        If oCFL.ObjectType = "Z_OWEM" Or oCFL.ObjectType = "Z_OBPC" Or oCFL.ObjectType = "Z_OITC" Or oCFL.ObjectType = "Z_OWHC" Then
                                            val = oDataTable.GetValue("U_Z_Code", 0)
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
                Case "5896"
                    If pVal.BeforeAction = False Then
                        'oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        'AddControls(oForm)
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
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                oApplication.Utilities.UpdateCategores()
                '   oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
