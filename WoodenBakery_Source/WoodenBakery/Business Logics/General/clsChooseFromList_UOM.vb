Public Class clsChooseFromList_UOM
    Inherits clsBase

#Region "Declarations"
    Public Shared ItemUID As String
    Public Shared SourceFormUID As String
    Public Shared SourceLabel As Integer
    Public Shared CFLChoice As String
    Public Shared ItemCode As String
    Public Shared sourceItemCode As String
    Public Shared choice As String
    Public Shared sourceColumID As String
    Public Shared sourcerowId As Integer
    Public Shared BinDescrUID As String
    Public Shared Documentchoice As String

    Private oDbDatasource As SAPbouiCOM.DBDataSource
    Private Ouserdatasource As SAPbouiCOM.UserDataSource
    Private oConditions As SAPbouiCOM.Conditions
    Private ocondition As SAPbouiCOM.Condition
    Private intRowId As Integer
    Private strRowNum As Integer
    Private i As Integer
    Private oedit As SAPbouiCOM.EditText
    '   Private oForm As SAPbouiCOM.Form
    Private objSoureceForm As SAPbouiCOM.Form
    Private objform As SAPbouiCOM.Form
    Private oMatrix As SAPbouiCOM.Grid
    Private osourcegrid As SAPbouiCOM.Matrix
    Private Const SEPRATOR As String = "~~~"
    Private SelectedRow As Integer
    Private sSearchColumn As String
    Private oItem As SAPbouiCOM.Item
    Public stritemid As SAPbouiCOM.Item
    Private intformmode As SAPbouiCOM.BoFormMode
    Private objGrid As SAPbouiCOM.Grid
    Private objSourcematrix As SAPbouiCOM.Matrix
    Private dtTemp As SAPbouiCOM.DataTable
    Private objStatic As SAPbouiCOM.StaticText
    Private inttable As Integer = 0
    Public strformid As String
    Public strStaticValue As String
    Public strSQL As String
    Private strSelectedItem1 As String
    Private strSelectedItem2 As String
    Private strSelectedItem3 As String
    Private strSelectedItem4, strSelectedItem5, strSelectedItem6 As String
    Private oRecSet As SAPbobsCOM.Recordset
    '   Private objSBOAPI As ClsSBO
    '   Dim objTransfer As clsTransfer
#End Region

#Region "New"
    '*****************************************************************
    'Type               : Constructor
    'Name               : New
    'Parameter          : 
    'Return Value       : 
    'Author             : DEV-2
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Create object for classes.
    '******************************************************************
    Public Sub New()
        '   objSBOAPI = New ClsSBO
        MyBase.New()
    End Sub
#End Region

#Region "Bind Data"
    '******************************************************************
    'Type               : Procedure
    'Name               : BindData
    'Parameter          : Form
    'Return Value       : 
    'Author             : DEV-2
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Binding the fields.
    '******************************************************************
    Public Sub databound(ByVal objform As SAPbouiCOM.Form)
        Try
            Dim strSQL As String = ""
            Dim ObjSegRecSet, oRecS As SAPbobsCOM.Recordset
            ObjSegRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objform.Freeze(True)
            objform.DataSources.DataTables.Add("dtLevel3")
            Ouserdatasource = objform.DataSources.UserDataSources.Add("dbFind", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 250)
            oedit = objform.Items.Item("etFind").Specific
            oedit.DataBind.SetBound(True, "", "dbFind")
            objGrid = objform.Items.Item("mtchoose").Specific
            dtTemp = objform.DataSources.DataTables.Item("dtLevel3")
            Dim strQuery As String = String.Empty
            oRecS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If choice = "PROMOTION" Then
                Select Case CFLChoice
                    Case "I"
                        strQuery = " SELECT  T2.""UomEntry"" as ""UomEntry"", T3.""UomCode"" as ""UomCode"" FROM OITM T0 INNER JOIN OUGP T1 ON T0.""UgpEntry"" = T1.""UgpEntry"" "
                        strQuery += " INNER JOIN UGP1 T2 ON T1.""UgpEntry"" = T2.""UgpEntry"" INNER JOIN OUOM T3 ON T3.""UomEntry"" = T2.""UomEntry"" "
                        strQuery += " Where T0.""ItemCode"" = '" & ItemCode & "' "
                        objform.Title = "Unit of Mesurement - Selection"
                    Case "O"
                        strQuery = " SELECT  T2.""UomEntry"" as ""UomEntry"", T3.""UomCode"" as ""UomCode"" FROM OITM T0 INNER JOIN OUGP T1 ON T0.""UgpEntry"" = T1.""UgpEntry"" "
                        strQuery += " INNER JOIN UGP1 T2 ON T1.""UgpEntry"" = T2.""UgpEntry"" INNER JOIN OUOM T3 ON T3.""UomEntry"" = T2.""UomEntry"" "
                        strQuery += " Where T0.""ItemCode"" = '" & ItemCode & "' "
                        objform.Title = "Unit of Mesurement - Selection"
                End Select

                If ItemCode <> "" And ItemCode <> "" Then
                    If Documentchoice = "" Then
                        dtTemp.ExecuteQuery(strQuery)
                        objGrid.DataTable = dtTemp
                    End If
                    objGrid.AutoResizeColumns()
                    objGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                End If
            ElseIf choice = "INVENTORY" Then
                Select Case CFLChoice
                    Case "I"
                        strQuery = " SELECT  T2.""UomEntry"" as ""UomEntry"", T3.""UomCode"" as ""UomCode"" FROM OITM T0 INNER JOIN OUGP T1 ON T0.""UgpEntry"" = T1.""UgpEntry"" "
                        strQuery += " INNER JOIN UGP1 T2 ON T1.""UgpEntry"" = T2.""UgpEntry"" INNER JOIN OUOM T3 ON T3.""UomEntry"" = T2.""UomEntry"" "
                        strQuery += " Where T0.""ItemCode"" = '" & ItemCode & "' "
                        objform.Title = "Unit of Mesurement - Selection"
                End Select

                If ItemCode <> "" And ItemCode <> "" Then
                    If Documentchoice = "" Then
                        dtTemp.ExecuteQuery(strQuery)
                        objGrid.DataTable = dtTemp
                    End If
                    objGrid.AutoResizeColumns()
                    objGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                End If
            End If

            objGrid.AutoResizeColumns()
            objGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            If objGrid.Rows.Count > 0 Then
                objGrid.Rows.SelectedRows.Add(0)
            End If
            objform.Freeze(False)
            objform.Update()
           
        Catch ex As Exception
            oApplication.SBO_Application.MessageBox(ex.Message)
            oApplication.SBO_Application.MessageBox(ex.Message)
        Finally
        End Try
    End Sub
#End Region

#Region "Get Form"
    '******************************************************************
    'Type               : Function
    'Name               : GetForm
    'Parameter          : FormUID
    'Return Value       : 
    'Author             : DEV-2
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : Get The Form
    '******************************************************************
    Public Function GetForm(ByVal FormUID As String) As SAPbouiCOM.Form
        Return oApplication.SBO_Application.Forms.Item(FormUID)
    End Function
#End Region

#Region "Item Event"

    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        BubbleEvent = True
        If pVal.FormTypeEx = frm_ChoosefromList_UOM Then
            If pVal.Before_Action = True Then
                If pVal.ItemUID = "mtchoose" Then
                    Try
                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.Row <> -1 Then
                            oForm = GetForm(pVal.FormUID)
                            oItem = CType(oForm.Items.Item(pVal.ItemUID), SAPbouiCOM.Item)
                            oMatrix = CType(oItem.Specific, SAPbouiCOM.Grid)
                            oMatrix.Rows.SelectedRows.Add(pVal.Row)
                        End If

                        If pVal.EventType = SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK And pVal.Row <> -1 Then
                            oForm = GetForm(pVal.FormUID)
                            oItem = oForm.Items.Item("mtchoose")
                            oMatrix = CType(oItem.Specific, SAPbouiCOM.Grid)
                            Dim inti As Integer
                            inti = 0
                            inti = 0
                            While inti <= oMatrix.DataTable.Rows.Count - 1
                                If oMatrix.Rows.IsSelected(inti) = True Then
                                    intRowId = inti
                                End If
                                System.Math.Min(System.Threading.Interlocked.Increment(inti), inti - 1)
                            End While
                            If CFLChoice <> "" Then

                                If intRowId = 0 Then
                                    strSelectedItem1 = Convert.ToString(oMatrix.DataTable.GetValue(0, intRowId))
                                    strSelectedItem2 = Convert.ToString(oMatrix.DataTable.GetValue(1, intRowId))
                                Else
                                    strSelectedItem1 = Convert.ToString(oMatrix.DataTable.GetValue(0, intRowId))
                                    strSelectedItem2 = Convert.ToString(oMatrix.DataTable.GetValue(1, intRowId))
                                End If

                                oForm.Close()
                                oForm = GetForm(SourceFormUID)
                                If choice = "PROMOTION" Then
                                    oForm.Freeze(True)
                                    'oMatrix = oForm.Items.Item("3").Specific
                                    Dim oBMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("3").Specific
                                    If CFLChoice = "I" Then
                                        oApplication.Utilities.SetMatrixValues(oBMatrix, "V_2_", sourcerowId, strSelectedItem2)
                                    ElseIf CFLChoice = "O" Then
                                        oApplication.Utilities.SetMatrixValues(oBMatrix, "V_5_", sourcerowId, strSelectedItem2)
                                    End If
                                    oForm.Freeze(False)
                                ElseIf choice = "INVENTORY" Then
                                    oForm.Freeze(True)
                                    'oMatrix = oForm.Items.Item("3").Specific
                                    Dim oBMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("3").Specific
                                    If CFLChoice = "I" Then
                                        oApplication.Utilities.SetMatrixValues(oBMatrix, "V_4", sourcerowId, strSelectedItem2)
                                    End If
                                    oForm.Freeze(False)
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        oApplication.SBO_Application.MessageBox(ex.Message)
                    End Try
                End If


                If pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN Then
                    Try
                        If pVal.ItemUID = "mtchoose" Then
                            oForm = GetForm(pVal.FormUID)
                            oItem = CType(oForm.Items.Item("mtchoose"), SAPbouiCOM.Item)
                            oMatrix = CType(oItem.Specific, SAPbouiCOM.Grid)
                            intRowId = pVal.Row - 1
                        End If
                        Dim inti As Integer
                        If pVal.CharPressed = 13 Then
                            inti = 0
                            inti = 0
                            oForm = GetForm(pVal.FormUID)
                            oItem = CType(oForm.Items.Item("mtchoose"), SAPbouiCOM.Item)

                            oMatrix = CType(oItem.Specific, SAPbouiCOM.Grid)
                            While inti <= oMatrix.DataTable.Rows.Count - 1
                                If oMatrix.Rows.IsSelected(inti) = True Then
                                    intRowId = inti
                                End If
                                System.Math.Min(System.Threading.Interlocked.Increment(inti), inti - 1)
                            End While
                            If CFLChoice <> "" Then
                                If intRowId = 0 Then
                                    strSelectedItem1 = Convert.ToString(oMatrix.DataTable.GetValue(0, intRowId))
                                    strSelectedItem2 = Convert.ToString(oMatrix.DataTable.GetValue(1, intRowId))
                                Else
                                    strSelectedItem1 = Convert.ToString(oMatrix.DataTable.GetValue(0, intRowId))
                                    strSelectedItem2 = Convert.ToString(oMatrix.DataTable.GetValue(1, intRowId))
                                End If
                                oForm.Close()
                                oForm = GetForm(SourceFormUID)
                                If choice = "PROMOTION" Then
                                    oForm.Freeze(True)
                                    '  oMatrix = oForm.Items.Item("3").Specific
                                    Dim oBMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("3").Specific
                                    If CFLChoice = "I" Then
                                        oApplication.Utilities.SetMatrixValues(oBMatrix, "V_2_", sourcerowId, strSelectedItem2)
                                    ElseIf CFLChoice = "O" Then
                                        oApplication.Utilities.SetMatrixValues(oBMatrix, "V_5_", sourcerowId, strSelectedItem2)
                                    End If
                                    oForm.Freeze(False)
                                ElseIf choice = "INVENTORY" Then
                                    oForm.Freeze(True)
                                    'oMatrix = oForm.Items.Item("3").Specific
                                    Dim oBMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("3").Specific
                                    If CFLChoice = "I" Then
                                        oApplication.Utilities.SetMatrixValues(oBMatrix, "V_4", sourcerowId, strSelectedItem2)
                                    End If
                                    oForm.Freeze(False)
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        oApplication.SBO_Application.MessageBox(ex.Message)
                    End Try
                End If


                If pVal.ItemUID = "btnChoose" AndAlso pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED Then
                    oForm = GetForm(pVal.FormUID)
                    oItem = oForm.Items.Item("mtchoose")
                    oMatrix = CType(oItem.Specific, SAPbouiCOM.Grid)
                    Dim inti As Integer
                    inti = 0
                    inti = 0
                    While inti <= oMatrix.DataTable.Rows.Count - 1
                        If oMatrix.Rows.IsSelected(inti) = True Then
                            intRowId = inti
                        End If
                        System.Math.Min(System.Threading.Interlocked.Increment(inti), inti - 1)
                    End While

                    If CFLChoice <> "" Then
                        If intRowId = 0 Then
                            strSelectedItem1 = Convert.ToString(oMatrix.DataTable.GetValue(0, intRowId))
                            strSelectedItem2 = Convert.ToString(oMatrix.DataTable.GetValue(1, intRowId))
                        Else
                            strSelectedItem1 = Convert.ToString(oMatrix.DataTable.GetValue(0, intRowId))
                            strSelectedItem2 = Convert.ToString(oMatrix.DataTable.GetValue(1, intRowId))
                        End If
                        oForm.Close()
                        oForm = GetForm(SourceFormUID)
                        If choice = "PROMOTION" Then
                            oForm.Freeze(True)
                            Dim oBMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("3").Specific
                            If CFLChoice = "I" Then
                                oApplication.Utilities.SetMatrixValues(oBMatrix, "V_2_", sourcerowId, strSelectedItem2)
                            ElseIf CFLChoice = "O" Then
                                oApplication.Utilities.SetMatrixValues(oBMatrix, "V_5_", sourcerowId, strSelectedItem2)
                            End If
                            oForm.Freeze(False)
                        ElseIf choice = "INVENTORY" Then
                            oForm.Freeze(True)
                            'oMatrix = oForm.Items.Item("3").Specific
                            Dim oBMatrix As SAPbouiCOM.Matrix = oForm.Items.Item("3").Specific
                            If CFLChoice = "I" Then
                                oApplication.Utilities.SetMatrixValues(oBMatrix, "V_4", sourcerowId, strSelectedItem2)
                            End If
                            oForm.Freeze(False)
                        End If
                    End If
                End If
            Else
                If pVal.BeforeAction = False Then
                    If (pVal.EventType = SAPbouiCOM.BoEventTypes.et_KEY_DOWN Or pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK) Then
                        BubbleEvent = False
                        Dim objGrid As SAPbouiCOM.Grid
                        Dim oedit As SAPbouiCOM.EditText
                        If pVal.ItemUID = "etFind" And pVal.CharPressed <> "13" Then
                            Dim i, j As Integer
                            Dim strItem As String
                            oForm = oApplication.SBO_Application.Forms.ActiveForm()
                            objGrid = oForm.Items.Item("mtchoose").Specific
                            oedit = oForm.Items.Item("etFind").Specific
                            For i = 0 To objGrid.DataTable.Rows.Count - 1
                                strItem = ""
                                strItem = objGrid.DataTable.GetValue(0, i)
                                For j = 1 To oedit.String.Length
                                    If oedit.String.Length <= strItem.Length Then
                                        If strItem.Substring(0, j).ToUpper = oedit.String.ToUpper Then
                                            objGrid.Rows.SelectedRows.Add(i)
                                            Exit Sub
                                        End If
                                    End If
                                Next
                            Next
                        End If
                    End If
                End If
            End If
        End If
    End Sub
#End Region

End Class
