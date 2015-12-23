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
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                'oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                'If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                '    If pVal.ItemUID = "1" Then
                                '        oForm.Freeze(True)
                                '        fillRProject(oForm)
                                '        changePrice(oForm)
                                '        oForm.Freeze(False)
                                '    End If
                                'End If                                
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
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
