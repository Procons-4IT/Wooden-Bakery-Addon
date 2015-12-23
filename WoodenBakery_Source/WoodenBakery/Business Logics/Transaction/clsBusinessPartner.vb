Public Class clsBusinessPartner
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



        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_BPMaster Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                AddChooseFromList(oForm)
                                'Dim frmUDFForm As SAPbouiCOM.Form
                                'Dim st As String = "-" & oForm.Type
                                'frmUDFForm = oApplication.SBO_Application.Forms.GetForm(st, oForm.TypeCount)
                                'oEditText = frmUDFForm.Items.Item("U_Z_WeekEnd").Specific
                                'oEditText.ChooseFromListUID = "CFL5"
                                'oEditText.ChooseFromListAlias = "U_Z_Code"
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
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

                                        If oCFL.ObjectType = "Z_OWEM" Then
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
            oForm = oApplication.SBO_Application.Forms.ActiveForm
            Select Case pVal.BeforeAction
                Case True

                Case False
                    Select Case pVal.MenuUID
                        Case mnu_InvSO
                        Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                        Case mnu_CPRL_C
                            If Not oForm.Items.Item("5").Specific.value = "" Then
                                Dim objPromList As clsCustPromotionList
                                objPromList = New clsCustPromotionList
                                objPromList.LoadForm(oForm.Items.Item("5").Specific.value)
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
            Select Case pVal.MenuUID
                Case mnu_InvSO
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

                        'Promotion List
                        oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                        If CType(oForm.Items.Item("40").Specific, SAPbouiCOM.ComboBox).Value = "C" Then
                            If Not oMenuItem.SubMenus.Exists(mnu_CPRL_C) Then
                                Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                                oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                                oCreationPackage.UniqueID = mnu_CPRL_C
                                oCreationPackage.String = "Promotion List"
                                oCreationPackage.Enabled = True
                                oMenus = oMenuItem.SubMenus
                                oMenus.AddEx(oCreationPackage)
                            End If
                        End If

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

                    oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                    If oMenuItem.SubMenus.Exists(mnu_CPRL_C) Then
                        oApplication.SBO_Application.Menus.RemoveEx(mnu_CPRL_C)
                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            End If
        End If
    End Sub

End Class
