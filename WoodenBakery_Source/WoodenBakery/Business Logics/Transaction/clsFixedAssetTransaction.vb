Imports System.IO
Public Class clsFixedAssetTransaction

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

#Region "Methods"
    Private Sub LoadForm()
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_FATransaction) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_FATransaction, frm_FATransaction)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.DataBrowser.BrowseBy = "5"
        AddChooseFromList(oForm)
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        oForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        oForm.Freeze(False)
    End Sub
    Public Sub LoadForm1(ByVal RequestCode As String)

        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_FATransaction) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_FATransaction, frm_FATransaction)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        AddChooseFromList(oForm)
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        oForm.Items.Item("5").Enabled = True
        oApplication.Utilities.setEdittextvalue(oForm, "5", RequestCode)
        oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE Then
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
        End If
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
        oForm.Freeze(False)
    End Sub
#End Region

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
            'oCFLCreationParams.ObjectType = "Z_HR_OBUOB"
            'oCFLCreationParams.UniqueID = "CFL1"
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
            'oCFL = oCFLs.Add(oCFLCreationParams)

            'oCFL = oCFLs.Item("CFL_2")
            'oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.Alias = "ItemType"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "F"
            'oCFL.SetConditions(oCons)

            addCFLCondition(objForm)

            'oCFL = oCFLs.Item("CFL_9")
            'oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.[Alias] = "U_Program"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "Y"
            'oCFL.SetConditions(oCons)


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub addCFLCondition(ByVal objForm As SAPbouiCOM.Form)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition

            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            oCFL = oCFLs.Item("CFL_2")
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.BracketOpenNum = 2
            oCon.Alias = "U_Z_USERCODE"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_CONTAIN
            oCon.CondVal = oApplication.Company.UserName
            oCon.BracketCloseNum = 1
            oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND

            oCon = oCons.Add
            oCon.BracketOpenNum = 1
            oCon.Alias = "ItemType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "f"
            oCon.BracketCloseNum = 2
            oCFL.SetConditions(oCons)

        Catch ex As Exception

        End Try
    End Sub

    Private Sub AddMode(ByVal aform As SAPbouiCOM.Form)
        Dim strCode As String
        Try
            aform.Freeze(True)
            'strCode = oApplication.Utilities.getMaxCode("@Z_OFATA", "DocEntry")
            aform.Items.Item("4").Enabled = True
            aform.Items.Item("7").Enabled = True
            ' oApplication.Utilities.setEdittextvalue(aform, "4", strCode)
            aform.Items.Item("7").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oApplication.Utilities.setEdittextvalue(aform, "7", "t")
            oApplication.SBO_Application.SendKeys("{TAB}")
            aform.Items.Item("31").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            Try
                aform.Items.Item("7").Enabled = False
            Catch ex As Exception

            End Try
            aform.Items.Item("1").Enabled = True
            aform.Freeze(False)
            oForm.Update()
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)

        End Try
    End Sub

    Private Sub fillopen()
        Dim mythr As New System.Threading.Thread(AddressOf ShowFileDialog)
        mythr.SetApartmentState(Threading.ApartmentState.STA)
        mythr.Start()
        mythr.Join()
    End Sub

    Private Sub ShowFileDialog()
        Dim oDialogBox As New OpenFileDialog
        Dim strFileName, strMdbFilePath As String
        Dim oProcesses() As Process
        Try
            Dim aform As New System.Windows.Forms.Form
            aform.TopMost = True

            oProcesses = Process.GetProcessesByName("SAP Business One")
            If oProcesses.Length <> 0 Then
                For i As Integer = 0 To oProcesses.Length - 1
                    Dim MyWindow As New clsListener.WindowWrapper(oProcesses(i).MainWindowHandle)
                    If oDialogBox.ShowDialog(MyWindow) = DialogResult.OK Then
                        strMdbFilePath = oDialogBox.FileName
                        strSelectedFilepath = oDialogBox.FileName
                        strFileName = strSelectedFilepath
                        strSelectedFolderPath = strFileName
                        If frmApprovalWOrksheetForm.TypeEx = frm_FATransaction Then
                            oApplication.Utilities.setEdittextvalue(frmApprovalWOrksheetForm, "21", strSelectedFolderPath)
                        End If
                        Exit For
                    Else
                        strSelectedFolderPath = ""
                        If frmApprovalWOrksheetForm.TypeEx = frm_FATransaction Then
                            oApplication.Utilities.setEdittextvalue(frmApprovalWOrksheetForm, "21", strSelectedFolderPath)
                        End If
                        Exit For
                    End If

                Next
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub

    Private Sub EnableControls(ByVal aform As SAPbouiCOM.Form)
        Try
            aform.Freeze(True)
            Select Case aform.Mode
                Case SAPbouiCOM.BoFormMode.fm_ADD_MODE
                    aform.Items.Item("4").Enabled = False
                    aform.Items.Item("6").Enabled = False
                Case SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    aform.Items.Item("4").Enabled = True
                    aform.Items.Item("6").Enabled = True
            End Select
            aform.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aform.Freeze(False)
        End Try
    End Sub

    Private Function Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        oCombobox = aForm.Items.Item("11").Specific
        If oCombobox.Selected.Value = "" Then
            oApplication.Utilities.Message("Transaction Type missing....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        If oApplication.Utilities.getEdittextvalue(aForm, "23") = "" Then
            oApplication.Utilities.Message("Fixed Asset code missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        oCombobox = aForm.Items.Item("11").Specific
        If oCombobox.Selected.Value <> "C" Then


            'If oApplication.Utilities.getEdittextvalue(aForm, "13") = "" Then
            '    oApplication.Utilities.Message("Transfer From missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If
            'If oApplication.Utilities.getEdittextvalue(aForm, "15") = "" Then
            '    oApplication.Utilities.Message("Transfer From  missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If
        Else
            If oApplication.Utilities.getEdittextvalue(aForm, "34") = "" Then
                oApplication.Utilities.Message("Valid From missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If oApplication.Utilities.getEdittextvalue(aForm, "36") = "" Then
                oApplication.Utilities.Message("Valid To  missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            If oApplication.Utilities.getEdittextvalue(aForm, "38") = "" Then
                oApplication.Utilities.Message("Distribution Rule is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        End If
        Return True
    End Function

    Private Sub CopyAttachment(ByVal Sfile As String)
        Try
            Dim oRec As SAPbobsCOM.Recordset

            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strQry = "Select AttachPath From OADP"
            oRec.DoQuery(strQry)
            Dim SPath As String = Sfile
            If SPath = "" Then
            Else
                Dim DPath As String = ""
                If Not oRec.EoF Then
                    DPath = oRec.Fields.Item("AttachPath").Value.ToString()
                End If
                If Not Directory.Exists(DPath) Then
                    Directory.CreateDirectory(DPath)
                End If
                Dim file = New FileInfo(SPath)
                Dim Filename As String = Path.GetFileName(SPath)
                Dim SavePath As String = Path.Combine(DPath, Filename)
                If System.IO.File.Exists(SavePath) Then
                Else
                    file.CopyTo(Path.Combine(DPath, file.Name), True)
                End If
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_FATransaction Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                If (pVal.ItemUID <> "1" And pVal.ItemUID <> "2") And (oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE) And (pVal.ItemUID <> "") Then
                                    Dim oDoc As SAPbouiCOM.DBDataSource
                                    oDoc = oForm.DataSources.DBDataSources.Item(0)
                                    Dim oItem As SAPbouiCOM.Item
                                    oItem = oForm.Items.Item(pVal.ItemUID)
                                    If oItem.Type <> SAPbouiCOM.BoFormItemTypes.it_BUTTON And oItem.Type <> SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON Then
                                        If (oDoc.GetValue("U_Z_DocStatus", 0).Trim <> "D") Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                    End If
                                    ' MsgBox(oDoc.GetValue("U_Z_DocStatus", 0))
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE) Then
                                    If Validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub

                                    Else
                                        CopyAttachment(oApplication.Utilities.getEdittextvalue(oForm, "21"))
                                    End If
                                End If
                                'If (pVal.ItemUID <> "1" And pVal.ItemUID <> "2") And (oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                                '    Dim oDoc As SAPbouiCOM.DBDataSource
                                '    oDoc = oForm.DataSources.DBDataSources.Item(0)
                                '    If (oDoc.GetValue("U_Z_DocStatus", 0) <> "D") Then
                                '        BubbleEvent = False
                                '        Exit Sub
                                '    End If

                                'End If
                                If pVal.ItemUID = "32" Then
                                    If oApplication.Utilities.getEdittextvalue(oForm, "21") <> "" Then
                                        oApplication.Utilities.LoadFiles(oApplication.Utilities.getEdittextvalue(oForm, "21"))
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                                If pVal.ItemUID = "26" Then
                                    Dim objHistory As New clsAppHistory
                                    objHistory.LoadForm(oForm, oApplication.Utilities.getEdittextvalue(oForm, "5"), HeaderDoctype.Fix)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "23" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                    'AddChooseFromList(oForm)
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oCombobox = oForm.Items.Item(pVal.ItemUID).Specific
                                If pVal.ItemUID = "11" Then
                                    If oCombobox.Selected.Value = "C" Then
                                        oForm.Items.Item("34").Enabled = True
                                        oForm.Items.Item("36").Enabled = True
                                        oForm.Items.Item("38").Enabled = True
                                        oForm.Items.Item("13").Enabled = False
                                        oForm.Items.Item("15").Enabled = False
                                    Else
                                        oForm.Items.Item("34").Enabled = False
                                        oForm.Items.Item("36").Enabled = False
                                        oForm.Items.Item("38").Enabled = False
                                        Dim oRec As SAPbobsCOM.Recordset
                                        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        Dim oTest As SAPbobsCOM.Recordset
                                        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        If oCombobox.Selected.Value = "L" Then
                                            oTest.DoQuery("SELECT T2.""Code"", T2.""Location"" FROM OITM  T1 INNER JOIN OLCT T2 ON T1.""Location"" = T2.""Code"" WHERE T1.""ItemCode""='" & oApplication.Utilities.getEdittextvalue(oForm, "23") & "'")
                                            If oTest.RecordCount > 0 Then
                                                oApplication.Utilities.setEdittextvalue(oForm, "13", oTest.Fields.Item(0).Value)
                                                oApplication.Utilities.setEdittextvalue(oForm, "28", oTest.Fields.Item(1).Value)
                                            Else
                                                oApplication.Utilities.setEdittextvalue(oForm, "13", "")
                                                oApplication.Utilities.setEdittextvalue(oForm, "28", "")
                                            End If

                                        Else
                                            oTest.DoQuery("SELECT T2.""empID"", T2.""firstName"" FROM OITM  T1 INNER JOIN OHEM T2 ON T1.""Employee"" = T2.""empID"" WHERE T1.""ItemCode""='" & oApplication.Utilities.getEdittextvalue(oForm, "23") & "'")
                                            If oTest.RecordCount > 0 Then
                                                oApplication.Utilities.setEdittextvalue(oForm, "13", oTest.Fields.Item(0).Value)
                                                oApplication.Utilities.setEdittextvalue(oForm, "28", oTest.Fields.Item(1).Value)
                                            Else
                                                oApplication.Utilities.setEdittextvalue(oForm, "13", "")
                                                oApplication.Utilities.setEdittextvalue(oForm, "28", "")
                                            End If
                                        End If
                                        oForm.Items.Item("15").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        oForm.Items.Item("13").Enabled = False
                                        oForm.Items.Item("15").Enabled = True
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "25" Then
                                    If oApplication.SBO_Application.MessageBox("Do you want to attach a file ", , "Continue", "Cancel") = 1 Then
                                        frmApprovalWOrksheetForm = oForm
                                        fillopen()
                                        strFilepath = oApplication.Utilities.getEdittextvalue(oForm, "21")
                                        If strFilepath = "" Then
                                            oApplication.Utilities.Message("Please Select a File", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            BubbleEvent = False
                                            Exit Sub
                                        Else
                                            ' strFilepath = ""
                                            ' oGrid.DataTable.Columns.Item("U_Z_Attachment").Cells.Item(pVal.Row).Value = strFilepath
                                        End If
                                    Else
                                        If strFilepath = "" Then
                                            strFilepath = ""
                                        End If

                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "11" Then
                                    oForm.Freeze(True)
                                    oDBSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_OFATA")
                                    If oDBSrc_Line.GetValue("U_Z_DocStatus", 0).ToString.Trim = "D" Then
                                        oApplication.Utilities.setEdittextvalue(oForm, "13", "")
                                        oApplication.Utilities.setEdittextvalue(oForm, "15", "")
                                        oApplication.Utilities.setEdittextvalue(oForm, "28", "")
                                        oApplication.Utilities.setEdittextvalue(oForm, "29", "")
                                    End If
                                    oForm.Freeze(False)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "38" And pVal.CharPressed = 9 Then
                                    Dim oDisClass As New clsDisRule


                                    oDisClass.SourceFormUID = oForm.UniqueID
                                    oDisClass.ItemUID = pVal.ItemUID
                                    oDisClass.strStaticValue = oApplication.Utilities.getEdittextvalue(oForm, pVal.ItemUID)
                                    oApplication.Utilities.LoadForm(xml_DisRule, frm_DisRule)
                                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                                    oDisClass.databound(oForm)
                                End If
                                If (pVal.ItemUID = "13" Or pVal.ItemUID = "15") And pVal.CharPressed = 9 And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                    Dim objChooseForm As SAPbouiCOM.Form
                                    Dim objChoose As New clsChooseFromList_Leave
                                    Dim strwhs, strProject, strGirdValue As String
                                    oCombobox = oForm.Items.Item("11").Specific
                                    Try
                                        strwhs = oCombobox.Selected.Value
                                    Catch ex As Exception
                                        strwhs = ""
                                    End Try
                                    If strwhs = "" Then
                                        Exit Sub
                                    End If
                                    strGirdValue = oApplication.Utilities.getEdittextvalue(oForm, pVal.ItemUID)
                                    Dim oTest As SAPbobsCOM.Recordset
                                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    If strGirdValue <> "" Then


                                        Select Case strwhs
                                            Case "L"
                                                oTest.DoQuery("SELECT T0.""Code"", T0.""Location"" FROM OLCT T0 where T0.""Location""='" & strGirdValue & "'")

                                            Case "E"
                                                oTest.DoQuery("SELECT T0.""empID"", T0.""firstName"" FROM OHEM T0 where T0.""empID""='" & strGirdValue & "'")
                                            Case "C"
                                                oTest.DoQuery("SELECT T0.""PrcCode"", T0.""PrcName"" FROM OPRC T0 where ""PrcCode""='" & strGirdValue & "'")

                                        End Select
                                        If oTest.RecordCount > 0 Then
                                            If pVal.ItemUID = "13" Then
                                                oApplication.Utilities.setEdittextvalue(oForm, "28", oTest.Fields.Item(1).Value)
                                            Else
                                                oApplication.Utilities.setEdittextvalue(oForm, "29", oTest.Fields.Item(1).Value)
                                            End If
                                            Exit Sub
                                        End If
                                    End If
                                    If strwhs <> "" Then
                                        objChoose.ItemUID = pVal.ItemUID
                                        objChoose.SourceFormUID = FormUID
                                        objChoose.SourceLabel = 0 'pVal.Row
                                        objChoose.CFLChoice = strwhs 'oCombo.Selected.Value
                                        objChoose.choice = "FATransfer"
                                        objChoose.ItemCode = strwhs
                                        objChoose.Documentchoice = "" ' oApplication.Utilities.GetDocType(oForm)
                                        If pVal.ItemUID = "13" Then
                                            objChoose.sourceColumID = "28"
                                        Else
                                            objChoose.sourceColumID = "29"
                                        End If

                                        objChoose.sourcerowId = 0 'pVal.Row
                                        objChoose.BinDescrUID = ""
                                        oApplication.Utilities.LoadForm("CFL_Leave.xml", frm_ChoosefromList_Leave)
                                        objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                        objChoose.databound(objChooseForm)
                                    End If
                                End If
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
                Case mnu_FATransaction
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If
                Case mnu_ADD
                    AddMode(oForm)
                Case mnu_FIND
                    If pVal.BeforeAction = False Then
                        oForm = oApplication.SBO_Application.Forms.ActiveForm()
                        EnableControls(oForm)
                    End If
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
        'If eventInfo.FormUID = "RightClk" Then
        If oForm.TypeEx = frm_FATransaction Then
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
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            End If
        End If
    End Sub

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                If BusinessObjectInfo.FormTypeEx = frm_FATransaction Then
                    '   MsgBox(BusinessObjectInfo.ObjectKey)

                    Dim s As String = oApplication.Company.GetNewObjectType()

                    Dim stXML As String = BusinessObjectInfo.ObjectKey
                    stXML = stXML.Replace("<?xml version=""1.0"" encoding=""UTF-16"" ?><Asset_TransactionParams><DocEntry>", "")
                    stXML = stXML.Replace("</DocEntry></Asset_TransactionParams>", "")
                    stXML = stXML.Replace("<?xml version=""1.0"" encoding=""UTF-16"" ?><Asset TransactionParams><DocEntry>", "")
                    stXML = stXML.Replace("</DocEntry></Asset TransactionParams>", "")
                    Dim otest As SAPbobsCOM.Recordset
                    otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    If stXML <> "" Then

                        otest.DoQuery("select * from ""@Z_OFATA""  where ""DocEntry""=" & stXML)
                        If otest.RecordCount > 0 Then
                            If otest.Fields.Item("U_Z_DocStatus").Value = "N" Then
                                Dim intTempID As String = oApplication.ApplProcedure.GetTemplateID(HeaderDoctype.Fix, oApplication.Company.UserName, "@Z_APPT1", "U_Z_EmpId") 'oApplication.Utilities.GetTemplateID(oForm, HeaderDoctype.Train, otest.Fields.Item("U_Z_HREmpID").Value)
                                If intTempID <> "0" Then
                                    oApplication.ApplProcedure.UpdateApprovalRequired("@Z_OFATA", "DocEntry", otest.Fields.Item("DocEntry").Value, "Y", intTempID, "P")
                                    Dim strMessage As String = "Fixed asset transaction need approval for the transaction id is : " & otest.Fields.Item("DocEntry").Value
                                    oApplication.ApplProcedure.InitialCurNextApprover("@Z_OFATA", "DocEntry", otest.Fields.Item("DocEntry").Value, intTempID, strMessage, "Fixed Asset Transaction Approval Notification")
                                    otest.DoQuery("Update ""@Z_OFATA"" set ""U_Z_DocStatus""='P' where ""DocEntry""=" & otest.Fields.Item("DocEntry").Value & "")
                                Else
                                    oApplication.ApplProcedure.UpdateApprovalRequired("@Z_OFATA", "DocEntry", otest.Fields.Item("DocEntry").Value, "N", intTempID, "A")
                                    oApplication.Utilities.UpdateFixedAsset(otest.Fields.Item("DocEntry").Value) 'Update Fixed Asset
                                    otest.DoQuery("Update ""@Z_OFATA"" set ""U_Z_DocStatus""='A' where ""DocEntry""=" & otest.Fields.Item("DocEntry").Value & "")

                                End If
                            End If
                        End If
                    End If
                End If
            End If
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD) Then
                If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD) Then
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    oForm.Items.Item("31").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    oForm.Items.Item("6").Enabled = True
                    oForm.Items.Item("4").Enabled = False
                    oDBSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_OFATA")
                    If oDBSrc_Line.GetValue("U_Z_DocStatus", 0).ToString.Trim <> "D" Then
                        '   oForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
                        oForm.Items.Item("1").Enabled = False
                    Else
                        oForm.Items.Item("1").Enabled = True
                        '  oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        ' oForm.Items.Item("70").Enabled = False
                    End If
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

End Class
