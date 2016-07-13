Imports System.IO
Imports SAPbobsCOM

Public Class clsUtilities


    Private strThousSep As String = ","
    Private strDecSep As String = "."
    Private intQtyDec As Integer = 3
    Private FormNum As Integer
    Private oRecordSet As SAPbobsCOM.Recordset

    Public Sub New()
        MyBase.New()
        FormNum = 1
    End Sub

    Public Function GetData(ByVal oForm As SAPbouiCOM.Form, ByVal oLoadForm As SAPbouiCOM.Form, ByVal strID As String, ByVal oMatrix As SAPbouiCOM.Matrix, ByVal oDBDataSourceLines As SAPbouiCOM.DBDataSource) As Boolean
        Dim _retVal As Boolean
        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            Dim strPath As String = CType(oForm.Items.Item(strID).Specific, SAPbouiCOM.StaticText).Caption
            If strPath.Length > 0 Then
                If strPath.Length > 0 Then

                    Dim intCol As Integer = 0
                    Dim txtRows() As String
                    Dim fields() As String
                    txtRows = System.IO.File.ReadAllLines(strPath)
                    Dim intRow As Integer = 0
                    intRow = 0

                    oMatrix.Clear()
                    oMatrix.FlushToDataSource()
                    oMatrix.LoadFromDataSource()
                    Dim intAddRows As Integer = txtRows.Length - 1
                    If intAddRows > 1 Then
                        intAddRows -= 1
                        oMatrix.AddRow(intAddRows + 1, -1)
                    End If
                    oMatrix.FlushToDataSource()

                    For Each txtrow As String In txtRows
                        If intRow = 0 Then
                            fields = txtrow.Split(",")
                        ElseIf intRow > 0 Then
                            fields = txtrow.Split(",")
                            If fields.Length > 3 Then


                                'CType(oLoadForm.Items.Item("4").Specific, SAPbouiCOM.StaticText).Caption = "Importing Item Code : " + fields(0) + " And Record No : " + intRow.ToString() + ""
                                oDBDataSourceLines.SetValue("LineId", intRow - 1, (intRow + 1).ToString())
                                oDBDataSourceLines.SetValue("U_Z_ItmCode", intRow - 1, fields(0))
                                Dim strQry As String = " Select ""ItemName"" From OITM Where ""ItemCode"" = '" & fields(0) & "'"
                                Dim oUOMRS As SAPbobsCOM.Recordset
                                oUOMRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oUOMRS.DoQuery(strQry)
                                If Not oUOMRS.EoF Then
                                    oDBDataSourceLines.SetValue("U_Z_ItmName", intRow - 1, oUOMRS.Fields.Item(0).Value)
                                End If
                                oDBDataSourceLines.SetValue("U_Z_WareHouse", intRow - 1, fields(1))
                                oDBDataSourceLines.SetValue("U_Z_Status", intRow - 1, "O")
                                oDBDataSourceLines.SetValue("U_Z_Qty", intRow - 1, fields(3))

                                If fields(2) <> "" Then

                                    oDBDataSourceLines.SetValue("U_Z_UOM", intRow - 1, fields(2))
                                    Dim dblIQty As Double
                                    Dim strIUOM As String = String.Empty
                                    getInventoryQty(fields(0), fields(2), CDbl(fields(3)), strIUOM, dblIQty)
                                    oDBDataSourceLines.SetValue("U_Z_IUOM", intRow - 1, strIUOM)
                                    oDBDataSourceLines.SetValue("U_Z_IQty", intRow - 1, dblIQty)
                                Else

                                    oDBDataSourceLines.SetValue("U_Z_UOM", intRow - 1, "Manual")
                                    oDBDataSourceLines.SetValue("U_Z_IUOM", intRow - 1, "Manual")
                                    oDBDataSourceLines.SetValue("U_Z_IQty", intRow - 1, fields(3))

                                End If

                            End If

                            
                        End If
                        intRow = intRow + 1
                        _retVal = True
                    Next
                    oMatrix.LoadFromDataSource()
                    oMatrix.FlushToDataSource()
                End If
            End If
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function getInventoryQty(ByVal strItemCode As String, ByVal strRUOM As String, ByVal dblRQty As Double, ByRef strIUOM As String, ByRef dblIQty As Double)
        Try
            Dim oUOMRS As SAPbobsCOM.Recordset
            Dim strQry As String = String.Empty
            Dim intUOMEntry, intIUOMEntry As Integer
            strQry = " Select ""UomEntry"" From OUOM Where ""UomCode"" = '" & strRUOM & "'"
            oUOMRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oUOMRS.DoQuery(strQry)
            If Not oUOMRS.EoF Then
                intUOMEntry = oUOMRS.Fields.Item(0).Value
                If intUOMEntry <> -1 Then

                    strQry = " Select T1.""AltQty"",T1.""BaseQty"" From OITM T0 JOIN UGP1 T1 "
                    strQry += " On T0.""UgpEntry"" = T1.""UgpEntry"""
                    strQry += " Where T1.""UomEntry"" = '" & intUOMEntry & "'"
                    strQry += " And T0.""ItemCode"" = '" & strItemCode & "'"
                    oUOMRS.DoQuery(strQry)
                    If Not oUOMRS.EoF Then

                        Dim dblAltQty As Double = CDbl(oUOMRS.Fields.Item(0).Value)
                        Dim dblBaseQty As Double = CDbl(oUOMRS.Fields.Item(1).Value)

                        Dim dblRPInvQty As Double = (1 / ((dblAltQty / dblBaseQty))) * CDbl(dblRQty)
                        Dim dblInvQty As Double = 0

                        strQry = " Select T0.""IUoMEntry"" From OITM T0  "
                        strQry += " Where T0.""ItemCode"" = '" & strItemCode & "'"
                        oUOMRS.DoQuery(strQry)
                        If Not oUOMRS.EoF Then
                            intIUOMEntry = oUOMRS.Fields.Item(0).Value
                        End If

                        If intIUOMEntry <> intUOMEntry Then

                            strQry = " Select T1.""AltQty"",T1.""BaseQty"",T2.""UomCode"" From OITM T0 JOIN UGP1 T1 "
                            strQry += " On T0.""UgpEntry"" = T1.""UgpEntry"" JOIN OUOM T2 On T2.""UomEntry"" = T0.""IUoMEntry"" "
                            strQry += " Where T1.""UomEntry"" = '" & intIUOMEntry & "'"
                            strQry += " And T0.""ItemCode"" = '" & strItemCode & "'"
                            oUOMRS.DoQuery(strQry)
                            If Not oUOMRS.EoF Then

                                Dim dblAltQty1 As Double = CDbl(oUOMRS.Fields.Item(0).Value)
                                Dim dblBaseQty1 As Double = CDbl(oUOMRS.Fields.Item(1).Value)

                                dblInvQty = (dblAltQty1 / dblBaseQty1) * dblRPInvQty

                                ' oDBDataSourceLines.SetValue("U_Z_IUOM", intRow - 1, oUOMRS.Fields.Item("UomCode").Value)
                                ' oDBDataSourceLines.SetValue("U_Z_IQty", intRow - 1, dblInvQty)
                                strIUOM = oUOMRS.Fields.Item("UomCode").Value
                                dblIQty = dblInvQty

                            End If
                        Else

                            strQry = " Select T2.""UomCode"" From OITM T0 JOIN UGP1 T1 "
                            strQry += " On T0.""UgpEntry"" = T1.""UgpEntry"" JOIN OUOM T2 On T2.""UomEntry"" = T0.""IUoMEntry"" "
                            strQry += " Where T1.""UomEntry"" = '" & intIUOMEntry & "'"
                            strQry += " And T0.""ItemCode"" = '" & strItemCode & "'"
                            oUOMRS.DoQuery(strQry)
                            If Not oUOMRS.EoF Then

                                ' Dim dblIBaseQty1 As Double = CDbl(oUOMRS.Fields.Item(0).Value)
                                ' dblInvQty = CDbl(dblBInvQty) * CDbl(dblIBaseQty1) * CDbl(fields(3))

                                ' oDBDataSourceLines.SetValue("U_Z_IUOM", intRow - 1, oUOMRS.Fields.Item("UomCode").Value)
                                ' oDBDataSourceLines.SetValue("U_Z_IQty", intRow - 1, CDbl(Fields(3)))

                                strIUOM = oUOMRS.Fields.Item("UomCode").Value
                                dblIQty = dblRQty

                            End If
                        End If

                    End If
                Else
                    'oDBDataSourceLines.SetValue("U_Z_UOM", intRow - 1, "Manual")
                    'oDBDataSourceLines.SetValue("U_Z_IUOM", intRow - 1, "Manual")
                    'oDBDataSourceLines.SetValue("U_Z_IQty", intRow - 1, Fields(3))

                    strIUOM = "Manual"
                    dblIQty = dblRQty

                End If
            End If
        Catch ex As Exception

        End Try


    End Function

    Public Function LoadMessageForm(ByVal XMLFile As String, ByVal FormType As String) As SAPbouiCOM.Form
        LoadXMLFiles(XMLFile)
        Return Nothing
    End Function

    Public Function ValidateFile(ByVal oForm As SAPbouiCOM.Form, ByVal strID As String) As Boolean
        Dim _retVal As Boolean = True
        Try
            Dim strPath As String = CType(oForm.Items.Item(strID).Specific, SAPbouiCOM.StaticText).Caption
            If Path.GetExtension(strPath) <> ".csv" And Path.GetExtension(strPath) <> ".CSV" Then
                _retVal = False
                oApplication.Utilities.Message("In Valid File Format...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Else
                _retVal = True
            End If
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Sub OpenFileDialogBox(ByVal oForm As SAPbouiCOM.Form, ByVal strID As String)
        Dim _retVal As String = String.Empty
        Try
            FileOpen()
            CType(oForm.Items.Item(strID).Specific, SAPbouiCOM.StaticText).Caption = strFilepath
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#Region "FileOpen"
    Private Sub FileOpen()
        Try
            Dim mythr As New System.Threading.Thread(AddressOf ShowFileDialog)
            mythr.SetApartmentState(Threading.ApartmentState.STA)
            mythr.Start()
            mythr.Join()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub ShowFileDialog()
        Try
            Dim oDialogBox As New OpenFileDialog
            Dim strMdbFilePath As String
            Dim oProcesses() As Process
            Try
                oProcesses = Process.GetProcessesByName("SAP Business One")
                If oProcesses.Length <> 0 Then
                    For i As Integer = 0 To oProcesses.Length - 1
                        Dim MyWindow As New clsListener.WindowWrapper(oProcesses(i).MainWindowHandle)
                        oDialogBox.Filter = " Excel | *.csv;*.CSV"
                        If oDialogBox.ShowDialog(MyWindow) = DialogResult.OK Then
                            strMdbFilePath = oDialogBox.FileName
                            strFilepath = oDialogBox.FileName
                            Exit For
                        Else
                            Exit For
                        End If
                    Next
                End If
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Finally
            End Try
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region


    Public Function ValidateItemIdentifier(aItemCode As String) As Boolean
        Dim oItem As SAPbobsCOM.Recordset
        oItem = oApplication.Company.GetBusinessObject(BoObjectTypes.BoRecordset)
        If blnIsHanaDB = True Then
            oItem.DoQuery("Select ifnull(""U_Z_Identifier"",'F') from OITM where ""ItemCode""='" & aItemCode & "'")
        Else
            oItem.DoQuery("Select isnull(""U_Z_Identifier"",'F') from OITM where ""ItemCode""='" & aItemCode & "'")
        End If
        If oItem.Fields.Item(0).Value = "F" Then
            Return False
        Else
            Return True
        End If

    End Function

    Public Function UnlockSpecificDate(aForm As SAPbouiCOM.Form) As Boolean
        Dim oUser As String = oApplication.Company.UserName
        Dim oRec As SAPbobsCOM.Recordset
        Dim dtDate, dtDate1 As Date
        Dim stDate, DocType, strDocEntry As String
        Dim oItem As SAPbouiCOM.Item
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        ' oRec.DoQuery("Select * from ""@Z_OLUSR"" where ""U_Z_Code""='" & oUser & "'") 'Check Mapping for the user
        ' strDocEntry = oRec.Fields.Item("DocEntry").Value
        DocType = aForm.TypeEx
        oRec.DoQuery("Select * from ""@Z_ODOC"" where ""U_Z_FormUID""='" & DocType & "'") 'Check DocType is available 
        If oRec.RecordCount > 0 Then
            DocType = oRec.Fields.Item("U_Z_Code").Value
            oItem = aForm.Items.Item(oRec.Fields.Item("U_Z_FieldID").Value)
            If oItem.Type = SAPbouiCOM.BoFormItemTypes.it_EDIT Then
                stDate = oApplication.Utilities.getEdittextvalue(aForm, oRec.Fields.Item("U_Z_FieldID").Value)
            Else
                Return True
            End If

            If stDate = "" Then
                Return True
            Else
                dtDate1 = GetDateTimeValue(stDate)
            End If
        Else
            Return True
        End If

        If dtDate1 = Now.Date Then
            Return True
        Else

            Dim s As String = "Select * from ""@Z_LUSR1"" T0 Inner Join ""@Z_OLUSR"" T1 On T1.""DocEntry""=T0.""DocEntry"" where T1.""U_Z_Super""='Y' and   T1.""U_Z_UserCode""='" & oUser & "'"
            oRec.DoQuery(s) ' and ""DocEntry"" =" & strDocEntry) 'Get Unlock Date
            If oRec.RecordCount > 0 Then
                Return True
            End If

            s = "Select * from ""@Z_LUSR1"" T0 Inner Join ""@Z_OLUSR"" T1 On T1.""DocEntry""=T0.""DocEntry"" where T0.""U_Z_Date""='" & dtDate1.ToString("yyyyMMdd") & "' and  T1.""U_Z_UserCode""='" & oUser & "' and  T0.""U_Z_Active""='Y' and  T0.""U_Z_DocType""='" & DocType & "'"
            oRec.DoQuery(s) ' and ""DocEntry"" =" & strDocEntry) 'Get Unlock Date
            If oRec.RecordCount > 0 Then
                Return True
            Else
                Message("you are not able to post previous date postings...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                Return False
            End If
        End If
    End Function

    Public Sub PopulateDocTotaltoPaymentMeans(aform As SAPbouiCOM.Form, aForm1 As SAPbouiCOM.Form)
        If aform.TypeEx = frm_ARInvoicePayment Then
            Dim strDoctotal As String = getEdittextvalue(aform, "33")
            Dim dblDocTotal As Double = getDocumentQuantity(strDoctotal)
            'Dim oForm As SAPbouiCOM.Form
            'oForm = oApplication.SBO_Application.Forms.ActiveForm()
            If aForm1.TypeEx = frm_PaymentMeans Then
                blnInvoiceForm = True
                aForm1.Items.Item("6").Click()
                aForm1.Items.Item("38").Click()
                ' oApplication.SBO_Application.SendKeys("{^B}")
                setEdittextvalue(aForm1, "38", strDoctotal)
                Try
                    '   oApplication.SBO_Application.SendKeys("{TAB}")
                    'aForm1.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    'aForm1.Close()
                Catch ex As Exception

                End Try


            End If
        End If
    End Sub

    Public Function GetDeliveryDate(aCardCode As String, aItemCode As String, aPostingDate As Date) As Date
        Dim oRec, oRec1 As SAPbobsCOM.Recordset
        Dim intWeekFrom, intWeekTo As Integer
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec.DoQuery("Select * from OCRD where ""CardCode""='" & aCardCode & "'")
        If oRec.Fields.Item("U_Z_WeekEnd").Value <> "" Then
            oRec1.DoQuery("Select * from ""@Z_OWEM"" where ""U_Z_Code""='" & oRec.Fields.Item("U_Z_WeekEnd").Value & "'")
        Else
            oRec1.DoQuery("Select * from ""@Z_OWEM"" where ""U_Z_Default""='Y'")
        End If
        If oRec1.RecordCount > 0 Then
            intWeekFrom = oRec1.Fields.Item("U_Z_From").Value
            intWeekTo = oRec1.Fields.Item("U_Z_End").Value
        Else
            intWeekFrom = 7
            intWeekTo = 1
        End If

        oRec1.DoQuery("Select * from OITM where ""ItemCode""='" & aItemCode & "'")
        If oRec1.Fields.Item("U_Z_DelDays").Value > 0 Then
            aPostingDate = DateAdd("D", oRec1.Fields.Item("U_Z_DelDays").Value, aPostingDate)
        End If
        If intWeekFrom = "8" Or intWeekTo = "8" Then
            Return aPostingDate
        End If

        Dim intDay As Integer = aPostingDate.DayOfWeek
        If intDay = intWeekFrom Then
            aPostingDate = DateAdd("D", 1, aPostingDate)
        End If
        intDay = aPostingDate.DayOfWeek

        If intDay = intWeekTo Then
            aPostingDate = DateAdd("D", 1, aPostingDate)
        End If
        Return aPostingDate
    End Function


    Public Function getsalesOrderDeliveryDate(aCardCode As String, aPostingDate As Date) As Date
        Dim oRec, oRec1 As SAPbobsCOM.Recordset
        Dim intWeekFrom, intWeekTo As Integer
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec.DoQuery("Select * from OCRD where ""CardCode""='" & aCardCode & "'")
        If oRec.Fields.Item("U_Z_WeekEnd").Value <> "" Then
            oRec1.DoQuery("Select * from ""@Z_OWEM"" where ""U_Z_Code""='" & oRec.Fields.Item("U_Z_WeekEnd").Value & "'")
        Else
            oRec1.DoQuery("Select * from ""@Z_OWEM"" where ""U_Z_Default""='Y'")
        End If
        If oRec1.RecordCount > 0 Then
            intWeekFrom = oRec1.Fields.Item("U_Z_From").Value
            intWeekTo = oRec1.Fields.Item("U_Z_End").Value
        Else
            intWeekFrom = 7
            intWeekTo = 1
        End If

        'oRec1.DoQuery("Select * from OITM where ""ItemCode""='" & aItemCode & "'")
        'If oRec1.Fields.Item("U_Z_DelDays").Value > 0 Then
        '    aPostingDate = DateAdd("D", oRec1.Fields.Item("U_Z_DelDays").Value, aPostingDate)
        'End If
        If intWeekFrom = "8" Or intWeekTo = "8" Then
            Return aPostingDate
        End If

        Dim intDay As Integer = aPostingDate.DayOfWeek
        If intDay = intWeekFrom Then
            aPostingDate = DateAdd("D", 1, aPostingDate)
        End If
        intDay = aPostingDate.DayOfWeek

        If intDay = intWeekTo Then
            aPostingDate = DateAdd("D", 1, aPostingDate)
        End If
        Return aPostingDate
    End Function

    Public Sub UpdateFixedAsset(aDocEntry As String)
        Dim oRec As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec.DoQuery("Select * from ""@Z_OFATA"" where ""DocEntry""='" & aDocEntry & "'")
        If oRec.RecordCount > 0 Then
            ' Dim oItem As SAPbobsCOM.Items
            '   oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

            If oRec.Fields.Item("U_Z_TransType").Value = "L" Then
                Dim intToLocation As Integer = CInt(oRec.Fields.Item("U_Z_ToCode").Value)
                oRec.DoQuery("Update OITM set ""Location""='" & intToLocation & "' where ""ItemCode""='" & oRec.Fields.Item("U_Z_Code").Value & "'")

            ElseIf oRec.Fields.Item("U_Z_TransType").Value = "E" Then
                Dim intToLocation As Integer = CInt(oRec.Fields.Item("U_Z_ToCode").Value)
                oRec.DoQuery("Update OITM set ""Employee""='" & intToLocation & "' where ""ItemCode""='" & oRec.Fields.Item("U_Z_Code").Value & "'")
            ElseIf oRec.Fields.Item("U_Z_TransType").Value = "C" Then
                'Dim oItem As SAPbobsCOM.Items
                'oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                'If oItem.GetByKey(oRec.Fields.Item("U_Z_Code").Value) Then
                '    Dim oDis As SAPbobsCOM.ItemsDistributionRules
                '    oDis = oItem.DistributionRules

                '    MsgBox(oDis.Count)
                '    Dim strDisrule As String = oRec.Fields.Item("U_Z_DisRule").Value
                '    Dim strRule As String() = strDisrule.Split(";")
                '    Dim intCount As Integer = oItem.DistributionRules.Count - 1
                '    If intCount > 0 Then
                '        oItem.DistributionRules.Add()
                '    End If

                '    oItem.DistributionRules.SetCurrentLine(intCount + 1)
                '    oItem.DistributionRules.ValidFrom = oRec.Fields.Item("U_Z_FromDate").Value
                '    oItem.DistributionRules.ValidTo = oRec.Fields.Item("U_Z_ToDate").Value
                '    If strRule(0) <> "" Then
                '        oItem.DistributionRules.DistributionRule = strRule(0).Trim
                '    End If
                '    If strRule(1) <> "" Then
                '        oItem.DistributionRules.DistributionRule2 = strRule(1).Trim
                '    End If
                '    If strRule(2) <> "" Then
                '        oItem.DistributionRules.DistributionRule3 = strRule(2).Trim
                '    End If
                '    If strRule(3) <> "" Then
                '        oItem.DistributionRules.DistributionRule4 = strRule(3).Trim
                '    End If
                '    If strRule(4) <> "" Then
                '        oItem.DistributionRules.DistributionRule5 = strRule(4).Trim
                '    End If
                '    oItem.Update()


                'End If

            End If

            ''   MsgBox(oRec.Fields.Item("U_Z_Code").Value)
            'If oItem.GetByKey(oRec.Fields.Item("U_Z_Code").Value) Then
            '    If oItem.ItemType = SAPbobsCOM.ItemTypeEnum.itFixedAssets Then
            '        '  oItem.AssetItem = SAPbobsCOM.BoYesNoEnum.tYES
            '        If oItem.AssetItem = SAPbobsCOM.BoYesNoEnum.tYES Then
            '            If oRec.Fields.Item("U_Z_TransType").Value = "L" Then
            '                Dim intToLocation As Integer = CInt(oRec.Fields.Item("U_Z_ToCode").Value)
            '                oItem.Location = intToLocation ' oRec.Fields.Item("U_Z_ToCode").Value
            '                oRec.DoQuery("Update OHEM set ""Location""=" & intToLocation & "' where ""ItemCode""='" & oRec.Fields.Item("U_Z_Code").Value & "'")

            '            ElseIf oRec.Fields.Item("U_Z_TransType").Value = "E" Then
            '                Dim intToLocation As Integer = CInt(oRec.Fields.Item("U_Z_ToCode").Value)
            '                oRec.DoQuery("Update OHEM set ""Employee""=" & intToLocation & "' where ""ItemCode""='" & oRec.Fields.Item("U_Z_Code").Value & "'")

            '            End If
            '            'oItem.ItemName = "DFD"
            '            'oItem.Location = 3

            '            'If oItem.Update() <> 0 Then
            '            '    MsgBox(oApplication.Company.GetLastErrorDescription)
            '            'End If

            '        End If
            '    End If

            'End If

        End If

    End Sub

    Public Function validateAuthorization(ByVal aUserId As String, ByVal aFormUID As String) As Boolean
        Dim oAuth As SAPbobsCOM.Recordset
        oAuth = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim struserid As String
        '    Return False
        struserid = oApplication.Company.UserName
        oAuth.DoQuery("select * from UPT1 where ""FormId""='" & aFormUID & "'")
        If (oAuth.RecordCount <= 0) Then
            Return True
        Else
            Dim st As String
            st = oAuth.Fields.Item("PermId").Value
            st = "Select * from USR3 where ""PermId""='" & st & "' and ""UserLink""=" & aUserId
            oAuth.DoQuery(st)
            If oAuth.RecordCount > 0 Then
                If oAuth.Fields.Item("Permission").Value = "N" Then
                    Return False
                End If
                Return True
            Else
                Return True
            End If

        End If

        Return True

    End Function

    Public Sub UpdateCategores()
        Dim oRec, oRec1 As SAPbobsCOM.Recordset
        Dim strWarehouseUser, strBPUser, strItemUser As String
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strItemUser = ""
        strWarehouseUser = ""
        strBPUser = ""
        oRec.DoQuery("Select * from ""@Z_OITC""")
        For intRow As Integer = 0 To oRec.RecordCount - 1
            strItemUser = ""
            oRec1.DoQuery("SELECT T0.""U_Z_UserCode"" FROM ""@Z_OLUSR""  T0  Inner Join  ""@Z_LUSR2""  T1 on T1.""DocEntry""=T0.""DocEntry"" WHERE T1.""U_Z_Code"" ='" & oRec.Fields.Item("U_Z_Code").Value & "'")
            For intLoop As Integer = 0 To oRec1.RecordCount - 1
                If strItemUser = "" Then
                    strItemUser = oRec1.Fields.Item(0).Value
                Else
                    strItemUser = strItemUser & "," & oRec1.Fields.Item(0).Value
                End If
                oRec1.MoveNext()
            Next
            If 1 = 1 Then 'strItemUser <> "" Then
                oRec1.DoQuery("Update OITM set ""U_Z_USERCODE""='" & strItemUser & "' where ""U_Z_ITCCODE""='" & oRec.Fields.Item("U_Z_Code").Value & "'")
            End If
            oRec.MoveNext()
        Next
        'Update BP master
        oRec.DoQuery("Select * from ""@Z_OBPC""")
        For intRow As Integer = 0 To oRec.RecordCount - 1
            strItemUser = ""
            oRec1.DoQuery("SELECT T0.""U_Z_UserCode"" FROM ""@Z_OLUSR""  T0  Inner Join  ""@Z_LUSR3""  T1 on T1.""DocEntry""=T0.""DocEntry"" WHERE T1.""U_Z_Code"" ='" & oRec.Fields.Item("U_Z_Code").Value & "'")
            For intLoop As Integer = 0 To oRec1.RecordCount - 1
                If strItemUser = "" Then
                    strItemUser = oRec1.Fields.Item(0).Value
                Else
                    strItemUser = strItemUser & "," & oRec1.Fields.Item(0).Value
                End If
                oRec1.MoveNext()
            Next
            If 1 = 1 Then 'If strItemUser <> "" Then
                oRec1.DoQuery("Update OCRD set ""U_Z_USERCODE""='" & strItemUser & "' where ""U_Z_BPCCODE""='" & oRec.Fields.Item("U_Z_Code").Value & "'")
            End If
            oRec.MoveNext()
        Next

        'Update Warehouse master
        oRec.DoQuery("Select * from ""@Z_OWHC""")
        For intRow As Integer = 0 To oRec.RecordCount - 1
            strItemUser = ""
            oRec1.DoQuery("SELECT T0.""U_Z_UserCode"" FROM ""@Z_OLUSR""  T0  Inner Join  ""@Z_LUSR4""  T1 on T1.""DocEntry""=T0.""DocEntry"" WHERE T1.""U_Z_Code"" ='" & oRec.Fields.Item("U_Z_Code").Value & "'")
            For intLoop As Integer = 0 To oRec1.RecordCount - 1
                If strItemUser = "" Then
                    strItemUser = oRec1.Fields.Item(0).Value
                Else
                    strItemUser = strItemUser & "," & oRec1.Fields.Item(0).Value
                End If
                oRec1.MoveNext()
            Next
            If 1 = 1 Then 'If strItemUser <> "" Then
                oRec1.DoQuery("Update OWHS set ""U_Z_USERCODE""='" & strItemUser & "' where ""U_Z_WHSCODE""='" & oRec.Fields.Item("U_Z_Code").Value & "'")
            End If
            oRec.MoveNext()
        Next


    End Sub

    Public Sub filterProjectChooseFromList(ByVal oForm As SAPbouiCOM.Form, ByVal strCFLID As String)
        Try
            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList
            If strCFLID <> "" Then


                oCFLs = oForm.ChooseFromLists
                oCFL = oCFLs.Item(strCFLID)

                Dim strUserCode As String = oApplication.Company.UserName
                If oCFL.ObjectType = "2" Then 'BP Code
                    If oForm.TypeEx <> frm_Z_OVPL Or oForm.TypeEx = frm_Z_OCPR Or oForm.TypeEx <> frm_Z_OICT Then
                        oCons = oCFL.GetConditions()
                        If oCons.Count = 0 Then
                            oCon = oCons.Add()
                        Else
                            oCon = oCons.Item(0)
                        End If
                        oCon.Alias = "U_Z_USERCODE"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_CONTAIN
                        oCon.CondVal = strUserCode
                        oCFL.SetConditions(oCons)
                    End If
                End If

                If oCFL.ObjectType = "4" Then 'Item Code
                    If oForm.TypeEx <> frm_FATransaction Then
                        oCons = oCFL.GetConditions()

                        If oCons.Count = 0 Then
                            oCon = oCons.Add()
                        Else
                            oCon = oCons.Item(0)
                        End If

                        oCon.Alias = "U_Z_USERCODE"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_CONTAIN
                        oCon.CondVal = strUserCode
                        oCFL.SetConditions(oCons)
                    Else

                       

                    End If

                End If
                If oCFL.ObjectType = "64" Then 'Warehouse
                    If oForm.TypeEx <> frm_Z_OICT Then
                        oCons = oCFL.GetConditions()

                        If oCons.Count = 0 Then
                            oCon = oCons.Add()
                        Else
                            oCon = oCons.Item(0)
                        End If

                        oCon.Alias = "U_Z_USERCODE"
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_CONTAIN
                        oCon.CondVal = strUserCode
                        oCFL.SetConditions(oCons)
                    End If

                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub assignMatrixLineno(ByVal aGrid As SAPbouiCOM.Grid, ByVal aform As SAPbouiCOM.Form)
        aform.Freeze(True)
        Try
            For intNo As Integer = 0 To aGrid.DataTable.Rows.Count - 1
                aGrid.RowHeaders.SetText(intNo, intNo + 1)
            Next
        Catch ex As Exception
        End Try
        aGrid.Columns.Item("RowsHeader").TitleObject.Caption = "#"
        aform.Freeze(False)
    End Sub

#Region "Close Open Sales Order Lines"

    Private Sub WriteErrorlog(ByVal aMessage As String, ByVal aPath As String)
        Dim aSw As System.IO.StreamWriter
        Try
            If File.Exists(aPath) Then
            End If
            aSw = New StreamWriter(aPath, True)
            aSw.WriteLine(aMessage)
            aSw.Flush()
            aSw.Close()
            aSw.Dispose()
        Catch ex As Exception
            MsgBox("test")
        End Try
    End Sub


    Public Function GetDateTimeValue(ByVal DateString As String) As DateTime
        Dim objBridge As SAPbobsCOM.SBObob
        objBridge = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Return objBridge.Format_StringToDate(DateString).Fields.Item(0).Value
    End Function

    Public Sub createARINvoice()
        Dim strCardcode, stritemcode As String
        Dim intbaseEntry, intbaserow As Integer
        Dim oInv As SAPbobsCOM.Documents
        strCardcode = "C20000"
        intbaseEntry = 66
        intbaserow = 1
        oInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
        oInv.DocDate = Now.Date
        oInv.CardCode = strCardcode
        oInv.Lines.BaseType = 17
        oInv.Lines.BaseEntry = intbaseEntry
        oInv.Lines.BaseLine = intbaserow
        oInv.Lines.Quantity = 1
        If oInv.Add <> 0 Then
            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Else
            oApplication.Utilities.Message("AR Invoice added", SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        End If

    End Sub

    Public Sub CloseOpenSOLines()
        Try
            Dim oDoc As SAPbobsCOM.Documents
            Dim oTemp As SAPbobsCOM.Recordset
            Dim strSQL, strSQL1, spath As String
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            spath = System.Windows.Forms.Application.StartupPath & "\Sales Order Matching ErrorLog.txt"
            If File.Exists(spath) Then
                File.Delete(spath)
            End If
            blnError = False

            ' oTemp.DoQuery("Select DocEntry,LineNum from RDR1 where ifnull(trgetentry,0)=0 and  LineStatus='O' and Quantity = ifnull(U_RemQty,0) order by DocEntry,LineNum")
            '            oTemp.DoQuery("Select DocEntry,VisOrder,LineNum from RDR1 where ifnull(trgetentry,0)=0 and  LineStatus='O' and Quantity = ifnull(U_RemQty,0) order by DocEntry,LineNum")

            Dim strQuery As String = String.Empty
            If blnIsHanaDB Then
                strQuery = "Select DocEntry,VisOrder,LineNum from RDR1 where   LineStatus='O' and Quantity = ifnull(U_RemQty,0) order by DocEntry,LineNum"
            Else
                strQuery = "Select DocEntry,VisOrder,LineNum from RDR1 where   LineStatus='O' and Quantity = isnull(U_RemQty,0) order by DocEntry,LineNum"
            End If
            oTemp.DoQuery(strQuery)
            oApplication.Utilities.Message("Processing closing Sales order Lines", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Dim numb As Integer
            For introw As Integer = 0 To oTemp.RecordCount - 1
                oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                numb = oTemp.Fields.Item(1).Value
                '  numb = oTemp.Fields.Item(2).Value
                If oDoc.GetByKey(oTemp.Fields.Item("DocEntry").Value) Then
                    oApplication.Utilities.Message("Processing Sales order :" & oDoc.DocNum, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    oDoc.Comments = oDoc.Comments & "XXX1"
                    If oDoc.Update() <> 0 Then
                        WriteErrorlog(" Error in Closing Sales order Line  SO No :" & oDoc.DocNum & " : Line No : " & oDoc.Lines.LineNum & " : Error : " & oApplication.Company.GetLastErrorDescription, spath)
                        blnError = True
                    Else
                        oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
                        If oDoc.GetByKey(oTemp.Fields.Item("DocEntry").Value) Then
                            Dim strcomments As String
                            strcomments = oDoc.Comments
                            strcomments = strcomments.Replace("XXX1", "")
                            oDoc.Comments = strcomments
                            oDoc.Lines.SetCurrentLine(numb)
                            '  MsgBox(oDoc.Lines.VisualOrder)
                            If oDoc.Lines.LineStatus <> SAPbobsCOM.BoStatus.bost_Close Then
                                oDoc.Lines.LineStatus = SAPbobsCOM.BoStatus.bost_Close
                            End If
                            If oDoc.Update <> 0 Then
                                WriteErrorlog(" Error in Closing Sales order Line  SO No :" & oDoc.DocNum & " : Line No : " & oDoc.Lines.LineNum & " : Error : " & oApplication.Company.GetLastErrorDescription, spath)
                                blnError = True
                                'oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Else
                                WriteErrorlog(" Sales order Line  SO No :" & oDoc.DocNum & " : Line No : " & oDoc.Lines.LineNum & " : Closed successfully  ", spath)
                            End If
                        End If
                    End If

                End If
                oTemp.MoveNext()
            Next
            oApplication.Utilities.Message("Operation completed succesfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            blnError = True
            ' oApplication.SBO_Application.MessageBox("Error Occured...")\
            spath = System.Windows.Forms.Application.StartupPath & "\Sales Order Matching ErrorLog.txt"
            If File.Exists(spath) Then
                Dim x As System.Diagnostics.ProcessStartInfo
                x = New System.Diagnostics.ProcessStartInfo
                x.UseShellExecute = True

                x.FileName = spath
                System.Diagnostics.Process.Start(x)
                x = Nothing
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

    Public Function createHRMainAuthorization() As Boolean
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim mUserPermission As SAPbobsCOM.UserPermissionTree
        mUserPermission = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree)
        '//Mandatory field, which is the key of the object.
        '//The partner namespace must be included as a prefix followed by _
        mUserPermission.PermissionID = "Wooden"
        '//The Name value that will be displayed in the General Authorization Tree
        mUserPermission.Name = "Wooden Bakery Addon"
        '//The permission that this object can get
        mUserPermission.Options = SAPbobsCOM.BoUPTOptions.bou_FullReadNone
        '//In case the level is one, there Is no need to set the FatherID parameter.
        '   mUserPermission.Levels = 1
        RetVal = mUserPermission.Add
        If RetVal = 0 Or -2035 Then
            Return True
        Else
            MsgBox(oApplication.Company.GetLastErrorDescription)
            Return False
        End If


    End Function

    Public Function addChildAuthorization(ByVal aChildID As String, ByVal aChildiDName As String, ByVal aorder As Integer, ByVal aFormType As String, ByVal aParentID As String, ByVal Permission As SAPbobsCOM.BoUPTOptions) As Boolean
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim mUserPermission As SAPbobsCOM.UserPermissionTree
        mUserPermission = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree)

        mUserPermission.PermissionID = aChildID
        mUserPermission.Name = aChildiDName
        mUserPermission.Options = Permission ' SAPbobsCOM.BoUPTOptions.bou_FullReadNone

        '//For level 2 and up you must set the object's father unique ID
        'mUserPermission.Level
        mUserPermission.ParentID = aParentID
        mUserPermission.UserPermissionForms.DisplayOrder = aorder
        '//this object manages forms
        ' If aFormType <> "" Then
        mUserPermission.UserPermissionForms.FormType = aFormType
        ' End If

        RetVal = mUserPermission.Add
        If RetVal = 0 Or RetVal = -2035 Then
            Return True
        Else
            MsgBox(oApplication.Company.GetLastErrorDescription)
            Return False
        End If


    End Function

    Public Sub AuthorizationCreation()
        addChildAuthorization("WSetup", " Setup", 2, "", "Wooden", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
        addChildAuthorization("WTrans", "Transactions", 2, "", "Wooden", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
        'Setup

        '  addChildAuthorization("WUserSetup", "User Security Setup", 3, "", "WSetup", SAPbobsCOM.BoUPTOptions.bou_FullNone)
        addChildAuthorization("WReasonCode", "ReasonCode", 4, frm_ReasonCode, "WSetup", SAPbobsCOM.BoUPTOptions.bou_FullNone)
        addChildAuthorization("WWeekEnd", "WeekEnd Master", 4, frm_WeekEndMaster, "WSetup", SAPbobsCOM.BoUPTOptions.bou_FullNone)
        addChildAuthorization("WDocType", "Document Type", 4, frm_DocumentType, "WSetup", SAPbobsCOM.BoUPTOptions.bou_FullNone)
        addChildAuthorization("WUserSec", "Unlock Posting Date", 4, frm_UnLock, "WSetup", SAPbobsCOM.BoUPTOptions.bou_FullNone)
        addChildAuthorization("WItmCat", "Item Category", 4, frm_ItemCagetory, "WSetup", SAPbobsCOM.BoUPTOptions.bou_FullNone)
        addChildAuthorization("WBPCat", "BP Category", 4, frm_BPmCagetory, "WSetup", SAPbobsCOM.BoUPTOptions.bou_FullNone)
        addChildAuthorization("WPWCat", "Warehouse Category", 4, frm_WhsCagetory, "WSetup", SAPbobsCOM.BoUPTOptions.bou_FullNone)
        addChildAuthorization("WRPS", "Rebate Posting Setup", 4, frm_Rebate, "WSetup", SAPbobsCOM.BoUPTOptions.bou_FullNone)
        'Transactions

        'Self Request Approval
        addChildAuthorization("WAsset", "Fixed Asset", 3, "", "WTrans", SAPbobsCOM.BoUPTOptions.bou_FullNone)
        addChildAuthorization("WTransaction", "Asset Transactions ", 4, frm_FATransaction, "WAsset", SAPbobsCOM.BoUPTOptions.bou_FullNone)
        addChildAuthorization("WTransApp", "Asset Transactions Approval ", 4, frm_FATransactionApp, "WAsset", SAPbobsCOM.BoUPTOptions.bou_FullNone)

        addChildAuthorization("DocClose", "Document Closing", 3, "", "WTrans", SAPbobsCOM.BoUPTOptions.bou_FullNone)
        addChildAuthorization("WLineCls", "Line Closing", 4, frm_SOClosing, "DocClose", SAPbobsCOM.BoUPTOptions.bou_FullNone)
        addChildAuthorization("LineClsRpt", "Line Closing Report", 4, frm_SOClosingRepot, "DocClose", SAPbobsCOM.BoUPTOptions.bou_FullNone)

        addChildAuthorization("SupplierPrice", "SupplierPrice", 3, "", "WTrans", SAPbobsCOM.BoUPTOptions.bou_FullNone)
        addChildAuthorization("SuppPriceDocument", "Price Document", 4, frm_Z_OVPL, "SupplierPrice", SAPbobsCOM.BoUPTOptions.bou_FullNone)
        addChildAuthorization("SuppPriceApprval", "Supplier Price Approval", 4, frm_Z_OVPL_A, "SupplierPrice", SAPbobsCOM.BoUPTOptions.bou_FullNone)

    End Sub

#Region "AddControls"
    Public Sub AddControls(ByVal objForm As SAPbouiCOM.Form, ByVal ItemUID As String, ByVal SourceUID As String, ByVal ItemType As SAPbouiCOM.BoFormItemTypes, ByVal position As String, Optional ByVal fromPane As Integer = 1, Optional ByVal toPane As Integer = 1, Optional ByVal linkedUID As String = "", Optional ByVal strCaption As String = "", Optional ByVal dblWidth As Double = 0, Optional ByVal dblTop As Double = 0, Optional ByVal Hight As Double = 0, Optional ByVal Enable As Boolean = True)
        Dim objNewItem, objOldItem As SAPbouiCOM.Item
        Dim ostatic As SAPbouiCOM.StaticText
        Dim oButton As SAPbouiCOM.Button
        Dim oCheckbox As SAPbouiCOM.CheckBox
        Dim ofolder As SAPbouiCOM.Folder
        objOldItem = objForm.Items.Item(SourceUID)
        objNewItem = objForm.Items.Add(ItemUID, ItemType)
        With objNewItem
            If ItemType = SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON Then
                .Left = objOldItem.Left - 15
                .Top = objOldItem.Top + 1
                .LinkTo = linkedUID
            Else
                If position.ToUpper = "RIGHT" Then
                    .Left = objOldItem.Left + objOldItem.Width + 2
                    .Top = objOldItem.Top

                ElseIf position.ToUpper = "DOWN" Then
                    If ItemUID = "edWork" Then
                        .Left = objOldItem.Left + 40
                    Else
                        .Left = objOldItem.Left
                    End If
                    .Top = objOldItem.Top + objOldItem.Height + 20

                    .Width = objOldItem.Width
                    .Height = objOldItem.Height
                End If
            End If
            .FromPane = fromPane
            .ToPane = toPane
            If ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
                .LinkTo = linkedUID
            End If
            .LinkTo = linkedUID
        End With
        If (ItemType = SAPbouiCOM.BoFormItemTypes.it_EDIT Or ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC) Then
            objNewItem.Width = objOldItem.Width
        End If
        If ItemType = SAPbouiCOM.BoFormItemTypes.it_BUTTON Then
            objNewItem.Width = objOldItem.Width '+ 50
            oButton = objNewItem.Specific
            oButton.Caption = strCaption
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_FOLDER Then
            ofolder = objNewItem.Specific
            ofolder.Caption = strCaption
            ofolder.GroupWith(linkedUID)
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
            ostatic = objNewItem.Specific
            ostatic.Caption = strCaption
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX Then
            oCheckbox = objNewItem.Specific
            oCheckbox.Caption = strCaption

        End If
        If dblWidth <> 0 Then
            objNewItem.Width = dblWidth
        End If

        If dblTop <> 0 Then
            objNewItem.Top = objNewItem.Top + dblTop
        End If
        If Hight <> 0 Then
            objNewItem.Height = objNewItem.Height + Hight
        End If
    End Sub
#End Region

#Region "Connect to Company"
    Public Sub Connect()
        Dim strCookie As String
        Dim strConnectionContext As String

        Try
            strCookie = oApplication.Company.GetContextCookie
            strConnectionContext = oApplication.SBO_Application.Company.GetConnectionContext(strCookie)

            If oApplication.Company.SetSboLoginContext(strConnectionContext) <> 0 Then
                Throw New Exception("Wrong login credentials.")
            End If

            'Open a connection to company
            If oApplication.Company.Connect() <> 0 Then
                Throw New Exception("Cannot connect to company database. ")
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Genral Functions"

#Region "Get MaxCode"
    Public Function getMaxCode(ByVal sTable As String, ByVal sColumn As String) As String
        Dim oRS As SAPbobsCOM.Recordset
        Dim MaxCode As Integer
        Dim sCode As String
        Dim strSQL As String
        Try
            strSQL = "SELECT MAX(CAST(""" & sColumn & """ AS Numeric)) FROM """ & sTable & """"
            ExecuteSQL(oRS, strSQL)

            If Convert.ToString(oRS.Fields.Item(0).Value).Length > 0 Then
                MaxCode = oRS.Fields.Item(0).Value + 1
            Else
                MaxCode = 1
            End If

            sCode = Format(MaxCode, "00000000")
            Return sCode
        Catch ex As Exception
            Throw ex
        Finally
            oRS = Nothing
        End Try
    End Function

    Public Function getPromotionCode(ByVal strCardCode As String, ByVal strPrmCode As String) As String
        Dim oRS As SAPbobsCOM.Recordset
        Dim MaxCode As Integer
        Dim sCode As String
        Dim strSQL As String
        Try
            strSQL = " Select ""Code"" From ""@Z_OCPR"" Where ""U_Z_CustCode"" = '" & strCardCode & "' And ""U_Z_PrCode"" = '" & strPrmCode & "' "
            ExecuteSQL(oRS, strSQL)
            If oRS.RecordCount > 0 Then
                sCode = oRS.Fields.Item("Code").Value
            End If
            Return sCode
        Catch ex As Exception
            Throw ex
        Finally
            oRS = Nothing
        End Try
    End Function

#End Region

#Region "Status Message"
    Public Sub Message(ByVal sMessage As String, ByVal StatusType As SAPbouiCOM.BoStatusBarMessageType)
        oApplication.SBO_Application.StatusBar.SetText(sMessage, SAPbouiCOM.BoMessageTime.bmt_Short, StatusType)
    End Sub
#End Region

#Region "Add Choose from List"
    Public Sub AddChooseFromList(ByVal FormUID As String, ByVal CFL_Text As String, ByVal CFL_Button As String, _
                                        ByVal ObjectType As SAPbouiCOM.BoLinkedObject, _
                                            Optional ByVal AliasName As String = "", Optional ByVal CondVal As String = "", _
                                                    Optional ByVal Operation As SAPbouiCOM.BoConditionOperation = SAPbouiCOM.BoConditionOperation.co_EQUAL)

        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        Try
            oCFLs = oApplication.SBO_Application.Forms.Item(FormUID).ChooseFromLists
            oCFLCreationParams = oApplication.SBO_Application.CreateObject( _
                                    SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            If ObjectType = SAPbouiCOM.BoLinkedObject.lf_Items Then
                oCFLCreationParams.MultiSelection = True
            Else
                oCFLCreationParams.MultiSelection = False
            End If

            oCFLCreationParams.ObjectType = ObjectType
            oCFLCreationParams.UniqueID = CFL_Text

            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL1

            oCons = oCFL.GetConditions()

            If Not AliasName = "" Then
                oCon = oCons.Add()
                oCon.Alias = AliasName
                oCon.Operation = Operation
                oCon.CondVal = CondVal
                oCFL.SetConditions(oCons)
            End If

            oCFLCreationParams.UniqueID = CFL_Button
            oCFL = oCFLs.Add(oCFLCreationParams)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Get Linked Object Type"
    Public Function getLinkedObjectType(ByVal Type As SAPbouiCOM.BoLinkedObject) As String
        Return CType(Type, String)
    End Function

#End Region

#Region "Execute Query"
    Public Sub ExecuteSQL(ByRef oRecordSet As SAPbobsCOM.Recordset, ByVal SQL As String)
        Try
            If oRecordSet Is Nothing Then
                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            End If

            oRecordSet.DoQuery(SQL)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Get Application path"
    Public Function getApplicationPath() As String

        Return Application.StartupPath.Trim

        'Return IO.Directory.GetParent(Application.StartupPath).ToString
    End Function
#End Region

#Region "Date Manipulation"

#Region "Convert SBO Date to System Date"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	ConvertStrToDate
    'Parameter          	:   ByVal oDate As String, ByVal strFormat As String
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	07/12/05
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To convert Date according to current culture info
    '********************************************************************
    Public Function ConvertStrToDate(ByVal strDate As String, ByVal strFormat As String) As DateTime
        Try
            Dim oDate As DateTime
            Dim ci As New System.Globalization.CultureInfo("en-GB", False)
            Dim newCi As System.Globalization.CultureInfo = CType(ci.Clone(), System.Globalization.CultureInfo)

            System.Threading.Thread.CurrentThread.CurrentCulture = newCi
            oDate = oDate.ParseExact(strDate, strFormat, ci.DateTimeFormat)

            Return oDate
        Catch ex As Exception
            Throw ex
        End Try

    End Function
#End Region

#Region " Get SBO Date Format in String (ddmmyyyy)"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	StrSBODateFormat
    'Parameter          	:   none
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To get date Format(ddmmyy value) as applicable to SBO
    '********************************************************************
    Public Function StrSBODateFormat() As String
        Try
            Dim rsDate As SAPbobsCOM.Recordset
            Dim strsql As String, GetDateFormat As String
            Dim DateSep As Char

            strsql = "Select DateFormat,DateSep from OADM"
            rsDate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsDate.DoQuery(strsql)
            DateSep = rsDate.Fields.Item(1).Value

            Select Case rsDate.Fields.Item(0).Value
                Case 0
                    GetDateFormat = "dd" & DateSep & "MM" & DateSep & "yy"
                Case 1
                    GetDateFormat = "dd" & DateSep & "MM" & DateSep & "yyyy"
                Case 2
                    GetDateFormat = "MM" & DateSep & "dd" & DateSep & "yy"
                Case 3
                    GetDateFormat = "MM" & DateSep & "dd" & DateSep & "yyyy"
                Case 4
                    GetDateFormat = "yyyy" & DateSep & "dd" & DateSep & "MM"
                Case 5
                    GetDateFormat = "dd" & DateSep & "MMM" & DateSep & "yyyy"
            End Select
            Return GetDateFormat

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "Get SBO date Format in Number"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	IntSBODateFormat
    'Parameter          	:   none
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To get date Format(integer value) as applicable to SBO
    '********************************************************************
    Public Function NumSBODateFormat() As String
        Try
            Dim rsDate As SAPbobsCOM.Recordset
            Dim strsql As String
            Dim DateSep As Char

            strsql = "Select DateFormat,DateSep from OADM"
            rsDate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsDate.DoQuery(strsql)
            DateSep = rsDate.Fields.Item(1).Value

            Select Case rsDate.Fields.Item(0).Value
                Case 0
                    NumSBODateFormat = 3
                Case 1
                    NumSBODateFormat = 103
                Case 2
                    NumSBODateFormat = 1
                Case 3
                    NumSBODateFormat = 120
                Case 4
                    NumSBODateFormat = 126
                Case 5
                    NumSBODateFormat = 130
            End Select
            Return NumSBODateFormat

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#End Region

#Region "Get Rental Period"
    Public Function getRentalDays(ByVal Date1 As String, ByVal Date2 As String, ByVal IsWeekDaysBilling As Boolean) As Integer
        Dim TotalDays, TotalDaysincSat, TotalBillableDays As Integer
        Dim TotalWeekEnds As Integer
        Dim StartDate As Date
        Dim EndDate As Date
        Dim oRecordset As SAPbobsCOM.Recordset

        StartDate = CType(Date1.Insert(4, "/").Insert(7, "/"), Date)
        EndDate = CType(Date2.Insert(4, "/").Insert(7, "/"), Date)

        TotalDays = DateDiff(DateInterval.Day, StartDate, EndDate)

        If IsWeekDaysBilling Then
            strSQL = " select dbo.WeekDays('" & Date1 & "','" & Date2 & "')"
            oApplication.Utilities.ExecuteSQL(oRecordset, strSQL)
            If oRecordset.RecordCount > 0 Then
                TotalBillableDays = oRecordset.Fields.Item(0).Value
            End If
            Return TotalBillableDays
        Else
            Return TotalDays + 1
        End If

    End Function

    Public Function WorkDays(ByVal dtBegin As Date, ByVal dtEnd As Date) As Long
        Try
            Dim dtFirstSunday As Date
            Dim dtLastSaturday As Date
            Dim lngWorkDays As Long

            ' get first sunday in range
            dtFirstSunday = dtBegin.AddDays((8 - Weekday(dtBegin)) Mod 7)

            ' get last saturday in range
            dtLastSaturday = dtEnd.AddDays(-(Weekday(dtEnd) Mod 7))

            ' get work days between first sunday and last saturday
            lngWorkDays = (((DateDiff(DateInterval.Day, dtFirstSunday, dtLastSaturday)) + 1) / 7) * 5

            ' if first sunday is not begin date
            If dtFirstSunday <> dtBegin Then

                ' assume first sunday is after begin date
                ' add workdays from begin date to first sunday
                lngWorkDays = lngWorkDays + (7 - Weekday(dtBegin))

            End If

            ' if last saturday is not end date
            If dtLastSaturday <> dtEnd Then

                ' assume last saturday is before end date
                ' add workdays from last saturday to end date
                lngWorkDays = lngWorkDays + (Weekday(dtEnd) - 1)

            End If

            WorkDays = lngWorkDays
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Function

#End Region

#Region "Get Item Price with Factor"
    Public Function getPrcWithFactor(ByVal CardCode As String, ByVal ItemCode As String, ByVal RntlDays As Integer, ByVal Qty As Double) As Double
        Dim oItem As SAPbobsCOM.Items
        Dim Price, Expressn As Double
        Dim oDataSet, oRecSet As SAPbobsCOM.Recordset

        oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        oApplication.Utilities.ExecuteSQL(oDataSet, "Select U_RentFac, U_NumDys From [@REN_FACT] order by U_NumDys ")
        If oItem.GetByKey(ItemCode) And oDataSet.RecordCount > 0 Then

            oApplication.Utilities.ExecuteSQL(oRecSet, "Select ListNum from OCRD where CardCode = '" & CardCode & "'")
            oItem.PriceList.SetCurrentLine(oRecSet.Fields.Item(0).Value - 1)
            Price = oItem.PriceList.Price
            Expressn = 0
            oDataSet.MoveFirst()

            While RntlDays > 0

                If oDataSet.EoF Then
                    oDataSet.MoveLast()
                End If

                If RntlDays < oDataSet.Fields.Item(1).Value Then
                    Expressn += (oDataSet.Fields.Item(0).Value * RntlDays * Price * Qty)
                    RntlDays = 0
                    Exit While
                End If
                Expressn += (oDataSet.Fields.Item(0).Value * oDataSet.Fields.Item(1).Value * Price * Qty)
                RntlDays -= oDataSet.Fields.Item(1).Value
                oDataSet.MoveNext()

            End While

        End If
        If oItem.UserFields.Fields.Item("U_Rental").Value = "Y" Then
            Return CDbl(Expressn / Qty)
        Else
            Return Price
        End If


    End Function
#End Region

#Region "Get WareHouse List"
    Public Function getUsedWareHousesList(ByVal ItemCode As String, ByVal Quantity As Double) As DataTable
        Dim oDataTable As DataTable
        Dim oRow As DataRow
        Dim rswhs As SAPbobsCOM.Recordset
        Dim LeftQty As Double
        Try
            oDataTable = New DataTable
            oDataTable.Columns.Add(New System.Data.DataColumn("ItemCode"))
            oDataTable.Columns.Add(New System.Data.DataColumn("WhsCode"))
            oDataTable.Columns.Add(New System.Data.DataColumn("Quantity"))

            strSQL = "Select WhsCode, ItemCode, (OnHand + OnOrder - IsCommited) As Available From OITW Where ItemCode = '" & ItemCode & "' And " & _
                        "WhsCode Not In (Select Whscode From OWHS Where U_Reserved = 'Y' Or U_Rental = 'Y') Order By (OnHand + OnOrder - IsCommited) Desc "

            ExecuteSQL(rswhs, strSQL)
            LeftQty = Quantity

            While Not rswhs.EoF
                oRow = oDataTable.NewRow()

                oRow.Item("WhsCode") = rswhs.Fields.Item("WhsCode").Value
                oRow.Item("ItemCode") = rswhs.Fields.Item("ItemCode").Value

                LeftQty = LeftQty - CType(rswhs.Fields.Item("Available").Value, Double)

                If LeftQty <= 0 Then
                    oRow.Item("Quantity") = CType(rswhs.Fields.Item("Available").Value, Double) + LeftQty
                    oDataTable.Rows.Add(oRow)
                    Exit While
                Else
                    oRow.Item("Quantity") = CType(rswhs.Fields.Item("Available").Value, Double)
                End If

                oDataTable.Rows.Add(oRow)
                rswhs.MoveNext()
                oRow = Nothing
            End While

            'strSQL = ""
            'For count As Integer = 0 To oDataTable.Rows.Count - 1
            '    strSQL += oDataTable.Rows(count).Item("WhsCode") & " : " & oDataTable.Rows(count).Item("Quantity") & vbNewLine
            'Next
            'MessageBox.Show(strSQL)

            Return oDataTable

        Catch ex As Exception
            Throw ex
        Finally
            oRow = Nothing
        End Try
    End Function
#End Region

#Region "GetDocumentQuantity"
    Public Function getDocumentQuantity(ByVal strQuantity As String) As Double
        Dim dblQuant As Double
        Dim strTemp, strTemp1 As String
        Dim oRec As SAPbobsCOM.Recordset
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec.DoQuery("Select ""CurrCode""  from OCRN")
        For intRow As Integer = 0 To oRec.RecordCount - 1
            strQuantity = strQuantity.Replace(oRec.Fields.Item(0).Value, "")
            oRec.MoveNext()
        Next
        strTemp1 = strQuantity
        strTemp = CompanyDecimalSeprator
        If CompanyDecimalSeprator <> "." Then
            If CompanyThousandSeprator <> strTemp Then
            End If
            strQuantity = strQuantity.Replace(".", ",")
        End If
        If strQuantity = "" Then
            Return 0
        End If
        Try
            dblQuant = Convert.ToDouble(strQuantity)
        Catch ex As Exception
            dblQuant = Convert.ToDouble(strTemp1)
        End Try

        Return dblQuant
    End Function
#End Region


#Region "Set / Get Values from Matrix"
    Public Function getMatrixValues(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal coluid As String, ByVal intRow As Integer) As String
        Return aMatrix.Columns.Item(coluid).Cells.Item(intRow).Specific.value
    End Function
    Public Sub SetMatrixValues(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal coluid As String, ByVal intRow As Integer, ByVal strvalue As String)
        aMatrix.Columns.Item(coluid).Cells.Item(intRow).Specific.value = strvalue
    End Sub
#End Region

#Region "Get Edit Text"
    Public Function getEdittextvalue(ByVal aform As SAPbouiCOM.Form, ByVal UID As String) As String
        Dim objEdit As SAPbouiCOM.EditText
        objEdit = aform.Items.Item(UID).Specific
        Return objEdit.String
    End Function
    Public Sub setEdittextvalue(ByVal aform As SAPbouiCOM.Form, ByVal UID As String, ByVal newvalue As String)
        Dim objEdit As SAPbouiCOM.EditText
        objEdit = aform.Items.Item(UID).Specific
        objEdit.String = newvalue
    End Sub
#End Region

#End Region

    Public Sub LoadFiles(ByVal aFileName As String)


        Dim strFilename, strFilePath As String
        strFilename = aFileName
        Dim Filename As String = Path.GetFileName(strFilename)
        strFilePath = aFileName

        If File.Exists(strFilePath) = False Then
            Dim oRec As SAPbobsCOM.Recordset
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strQry = "Select ""AttachPath"" From OADP"
            oRec.DoQuery(strQry)
            strFilePath = oRec.Fields.Item(0).Value

            If Filename = "" Then
                strFilePath = strFilePath
            Else
                strFilePath = strFilePath & Filename
            End If
            If File.Exists(strFilePath) = False Then
                oApplication.Utilities.Message("File does not exists ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Exit Sub
            End If
            strFilename = strFilePath
        Else
            strFilename = strFilePath
        End If
        Dim x As System.Diagnostics.ProcessStartInfo
        x = New System.Diagnostics.ProcessStartInfo
        x.UseShellExecute = True
        x.FileName = strFilename
        System.Diagnostics.Process.Start(x)
        x = Nothing
        Exit Sub


    End Sub

    Public Function getAccountCode(ByVal aCode As String) As String
        Dim oRS As SAPbobsCOM.Recordset
        oRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRS.DoQuery("Select ""AcctCode"" from OACT where ""FormatCode""='" & aCode & "'")
        If oRS.RecordCount > 0 Then
            '    MsgBox(oRS.Fields.Item(0).Value)
            Return oRS.Fields.Item(0).Value
        Else
            Return ""
        End If
    End Function

    Public Sub setUserDatabind(ByVal aForm As SAPbouiCOM.Form, ByVal UID As String, ByVal strDBID As String)
        Dim objEdit As SAPbouiCOM.EditText
        objEdit = aForm.Items.Item(UID).Specific
        objEdit.DataBind.SetBound(True, "", strDBID)
    End Sub

    Public Function GetCode(ByVal sTableName As String) As String
        Dim oRecSet As SAPbobsCOM.Recordset
        Dim sQuery As String
        oRecSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        sQuery = "SELECT Top 1 ""DocEntry"" FROM " & sTableName + " ORDER BY Convert(Int,""DocEntry"") desc"
        oRecSet.DoQuery(sQuery)
        If Not oRecSet.EoF Then
            GetCode = Convert.ToInt32(oRecSet.Fields.Item(0).Value.ToString()) + 1
        Else
            GetCode = "1"
        End If
    End Function

#Region "Functions related to Load XML"

#Region "Add/Remove Menus "
    Public Sub AddRemoveMenus(ByVal sFileName As String)
        Dim oXMLDoc As New Xml.XmlDocument
        Dim sFilePath As String
        Try
            sFilePath = getApplicationPath() & "\XML Files\" & sFileName
            oXMLDoc.Load(sFilePath)
            oApplication.SBO_Application.LoadBatchActions(oXMLDoc.InnerXml)
        Catch ex As Exception
            Throw ex
        Finally
            oXMLDoc = Nothing
        End Try
    End Sub
#End Region

#Region "Load XML File "
    Private Function LoadXMLFiles(ByVal sFileName As String) As String
        Dim oXmlDoc As Xml.XmlDocument
        Dim oXNode As Xml.XmlNode
        Dim oAttr As Xml.XmlAttribute
        Dim sPath As String
        Dim FrmUID As String
        Try
            oXmlDoc = New Xml.XmlDocument

            sPath = getApplicationPath() & "\XML Files\" & sFileName

            oXmlDoc.Load(sPath)
            oXNode = oXmlDoc.GetElementsByTagName("form").Item(0)
            oAttr = oXNode.Attributes.GetNamedItem("uid")
            oAttr.Value = oAttr.Value & FormNum
            FormNum = FormNum + 1
            oApplication.SBO_Application.LoadBatchActions(oXmlDoc.InnerXml)
            FrmUID = oAttr.Value

            Return FrmUID

        Catch ex As Exception
            Throw ex
        Finally
            oXmlDoc = Nothing
        End Try
    End Function
#End Region

    Public Function LoadForm(ByVal XMLFile As String, ByVal FormType As String) As SAPbouiCOM.Form
        'Return LoadForm(XMLFile, FormType.ToString(), FormType & "_" & oApplication.SBO_Application.Forms.Count.ToString)
        LoadXMLFiles(XMLFile)
        Return Nothing
    End Function

    '*****************************************************************
    'Type               : Function   
    'Name               : LoadForm
    'Parameter          : XmlFile,FormType,FormUID
    'Return Value       : SBO Form
    'Author             : Senthil Kumar B Senthil Kumar B
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Load XML file 
    '*****************************************************************

    Public Function LoadForm(ByVal XMLFile As String, ByVal FormType As String, ByVal FormUID As String) As SAPbouiCOM.Form

        Dim oXML As System.Xml.XmlDocument
        Dim objFormCreationParams As SAPbouiCOM.FormCreationParams
        Try
            oXML = New System.Xml.XmlDocument
            oXML.Load(XMLFile)
            objFormCreationParams = (oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams))
            objFormCreationParams.XmlData = oXML.InnerXml
            objFormCreationParams.FormType = FormType
            objFormCreationParams.UniqueID = FormUID
            Return oApplication.SBO_Application.Forms.AddEx(objFormCreationParams)
        Catch ex As Exception
            Throw ex

        End Try

    End Function



#Region "Load Forms"
    Public Sub LoadForm(ByRef oObject As Object, ByVal XmlFile As String)
        Try
            oObject.FrmUID = LoadXMLFiles(XmlFile)
            oObject.Form = oApplication.SBO_Application.Forms.Item(oObject.FrmUID)
            If Not oApplication.Collection.ContainsKey(oObject.FrmUID) Then
                oApplication.Collection.Add(oObject.FrmUID, oObject)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#End Region

#Region "Functions related to System Initilization"

#Region "Create Tables"
    Public Sub CreateTables()
        Dim oCreateTable As clsTable
        Try
            oCreateTable = New clsTable
            oCreateTable.CreateTables()
        Catch ex As Exception
            Throw ex
        Finally
            oCreateTable = Nothing
        End Try
    End Sub
#End Region

#Region "Notify Alert"
    Public Sub NotifyAlert()
        'Dim oAlert As clsPromptAlert

        'Try
        '    oAlert = New clsPromptAlert
        '    oAlert.AlertforEndingOrdr()
        'Catch ex As Exception
        '    Throw ex
        'Finally
        '    oAlert = Nothing
        'End Try

    End Sub
#End Region

#End Region

#Region "Function related to Quantities"

#Region "Get Available Quantity"
    Public Function getAvailableQty(ByVal ItemCode As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset

        strSQL = "Select SUM(T1.OnHand + T1.OnOrder - T1.IsCommited) From OITW T1 Left Outer Join OWHS T3 On T3.Whscode = T1.WhsCode " & _
                    "Where T1.ItemCode = '" & ItemCode & "'"
        Me.ExecuteSQL(rsQuantity, strSQL)

        If rsQuantity.Fields.Item(0) Is System.DBNull.Value Then
            Return 0
        Else
            Return CLng(rsQuantity.Fields.Item(0).Value)
        End If

    End Function
#End Region

#Region "Get Rented Quantity"
    Public Function getRentedQty(ByVal ItemCode As String, ByVal StartDate As String, ByVal EndDate As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset
        Dim RentedQty As Long

        strSQL = " select Sum(U_ReqdQty) from [@REN_RDR1] Where U_ItemCode = '" & ItemCode & "' " & _
                    " And DocEntry IN " & _
                    " (Select DocEntry from [@REN_ORDR] Where U_Status = 'R') " & _
                    " and '" & StartDate & "' between [@REN_RDR1].U_ShipDt1 and [@REN_RDR1].U_ShipDt2 "
        '" and [@REN_RDR1].U_ShipDt1 between '" & StartDate & "' and '" & EndDate & "'"

        ExecuteSQL(rsQuantity, strSQL)
        If Not rsQuantity.Fields.Item(0).Value Is System.DBNull.Value Then
            RentedQty = rsQuantity.Fields.Item(0).Value
        End If

        Return RentedQty

    End Function
#End Region

#Region "Get Reserved Quantity"
    Public Function getReservedQty(ByVal ItemCode As String, ByVal StartDate As String, ByVal EndDate As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset
        Dim ReservedQty As Long

        strSQL = " select Sum(U_ReqdQty) from [@REN_QUT1] Where U_ItemCode = '" & ItemCode & "' " & _
                    " And DocEntry IN " & _
                    " (Select DocEntry from [@REN_OQUT] Where U_Status = 'R' And Status = 'O') " & _
                    " and '" & StartDate & "' between [@REN_QUT1].U_ShipDt1 and [@REN_QUT1].U_ShipDt2"

        ExecuteSQL(rsQuantity, strSQL)
        If Not rsQuantity.Fields.Item(0).Value Is System.DBNull.Value Then
            ReservedQty = rsQuantity.Fields.Item(0).Value
        End If

        Return ReservedQty

    End Function
#End Region

#End Region

#Region "Functions related to Tax"

#Region "Get Tax Codes"
    Public Sub getTaxCodes(ByRef oCombo As SAPbouiCOM.ComboBox)
        Dim rsTaxCodes As SAPbobsCOM.Recordset

        strSQL = "Select Code, Name From OVTG Where Category = 'O' Order By Name"
        Me.ExecuteSQL(rsTaxCodes, strSQL)

        oCombo.ValidValues.Add("", "")
        If rsTaxCodes.RecordCount > 0 Then
            While Not rsTaxCodes.EoF
                oCombo.ValidValues.Add(rsTaxCodes.Fields.Item(0).Value, rsTaxCodes.Fields.Item(1).Value)
                rsTaxCodes.MoveNext()
            End While
        End If
        oCombo.ValidValues.Add("Define New", "Define New")
        'oCombo.Select("")
    End Sub
#End Region

#Region "Get Applicable Code"

    Public Function getApplicableTaxCode1(ByVal CardCode As String, ByVal ItemCode As String, ByVal Shipto As String) As String
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oItem As SAPbobsCOM.Items
        Dim rsExempt As SAPbobsCOM.Recordset
        Dim TaxGroup As String
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        If oBP.GetByKey(CardCode.Trim) Then
            If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
                If oBP.VatGroup.Trim <> "" Then
                    TaxGroup = oBP.VatGroup.Trim
                Else
                    strSQL = "select LicTradNum from CRD1 where Address ='" & Shipto & "' and CardCode ='" & CardCode & "'"
                    Me.ExecuteSQL(rsExempt, strSQL)
                    If rsExempt.RecordCount > 0 Then
                        rsExempt.MoveFirst()
                        TaxGroup = rsExempt.Fields.Item(0).Value
                    Else
                        TaxGroup = ""
                    End If
                    'TaxGroup = oBP.FederalTaxID
                End If
            ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
                strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
                Me.ExecuteSQL(rsExempt, strSQL)
                If rsExempt.RecordCount > 0 Then
                    rsExempt.MoveFirst()
                    TaxGroup = rsExempt.Fields.Item(0).Value
                Else
                    TaxGroup = ""
                End If
            End If
        End If




        Return TaxGroup

    End Function


    Public Function getApplicableTaxCode(ByVal CardCode As String, ByVal ItemCode As String) As String
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oItem As SAPbobsCOM.Items
        Dim rsExempt As SAPbobsCOM.Recordset
        Dim TaxGroup As String
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        If oBP.GetByKey(CardCode.Trim) Then
            If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
                If oBP.VatGroup.Trim <> "" Then
                    TaxGroup = oBP.VatGroup.Trim
                Else
                    TaxGroup = oBP.FederalTaxID
                End If
            ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
                strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
                Me.ExecuteSQL(rsExempt, strSQL)
                If rsExempt.RecordCount > 0 Then
                    rsExempt.MoveFirst()
                    TaxGroup = rsExempt.Fields.Item(0).Value
                Else
                    TaxGroup = ""
                End If
            End If
        End If

        'If oBP.GetByKey(CardCode.Trim) Then
        '    If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
        '        If oBP.VatGroup.Trim <> "" Then
        '            TaxGroup = oBP.VatGroup.Trim
        '        Else
        '            oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        '            If oItem.GetByKey(ItemCode.Trim) Then
        '                TaxGroup = oItem.SalesVATGroup.Trim
        '            End If
        '        End If
        '    ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
        '        strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
        '        Me.ExecuteSQL(rsExempt, strSQL)
        '        If rsExempt.RecordCount > 0 Then
        '            rsExempt.MoveFirst()
        '            TaxGroup = rsExempt.Fields.Item(0).Value
        '        Else
        '            TaxGroup = ""
        '        End If
        '    End If
        'End If
        Return TaxGroup

    End Function
#End Region

#End Region

#Region "Log Transaction"
    Public Sub LogTransaction(ByVal DocNum As Integer, ByVal ItemCode As String, _
                                    ByVal FromWhs As String, ByVal TransferedQty As Double, ByVal ProcessDate As Date)
        Dim sCode As String
        Dim sColumns As String
        Dim sValues As String
        Dim rsInsert As SAPbobsCOM.Recordset

        sCode = Me.getMaxCode("@REN_PORDR", "Code")

        sColumns = "Code, Name, U_DocNum, U_WhsCode, U_ItemCode, U_Quantity, U_RetQty, U_Date"
        sValues = "'" & sCode & "','" & sCode & "'," & DocNum & ",'" & FromWhs & "','" & ItemCode & "'," & TransferedQty & ", 0, Convert(DateTime,'" & ProcessDate.ToString("yyyyMMdd") & "')"

        strSQL = "Insert into [@REN_PORDR] (" & sColumns & ") Values (" & sValues & ")"
        oApplication.Utilities.ExecuteSQL(rsInsert, strSQL)

    End Sub

    Public Sub LogCreatedDocument(ByVal DocNum As Integer, ByVal CreatedDocType As SAPbouiCOM.BoLinkedObject, ByVal CreatedDocNum As String, ByVal sCreatedDate As String)
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim sCode As String
        Dim CreatedDate As DateTime
        Try
            oUserTable = oApplication.Company.UserTables.Item("REN_DORDR")

            sCode = Me.getMaxCode("@REN_DORDR", "Code")

            If Not oUserTable.GetByKey(sCode) Then
                oUserTable.Code = sCode
                oUserTable.Name = sCode

                With oUserTable.UserFields.Fields
                    .Item("U_DocNum").Value = DocNum
                    .Item("U_DocType").Value = CInt(CreatedDocType)
                    .Item("U_DocEntry").Value = CInt(CreatedDocNum)

                    If sCreatedDate <> "" Then
                        CreatedDate = CDate(sCreatedDate.Insert(4, "/").Insert(7, "/"))
                        .Item("U_Date").Value = CreatedDate
                    Else
                        .Item("U_Date").Value = CDate(Format(Now, "Long Date"))
                    End If

                End With

                If oUserTable.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If

        Catch ex As Exception
            Throw ex
        Finally
            oUserTable = Nothing
        End Try
    End Sub
#End Region

    Public Function FormatDataSourceValue(ByVal Value As String) As Double
        Dim NewValue As Double

        If Value <> "" Then
            If Value.IndexOf(".") > -1 Then
                Value = Value.Replace(".", CompanyDecimalSeprator)
            End If

            If Value.IndexOf(CompanyThousandSeprator) > -1 Then
                Value = Value.Replace(CompanyThousandSeprator, "")
            End If
        Else
            Value = "0"

        End If

        ' NewValue = CDbl(Value)
        NewValue = Val(Value)

        Return NewValue


        'Dim dblValue As Double
        'Value = Value.Replace(CompanyThousandSeprator, "")
        'Value = Value.Replace(CompanyDecimalSeprator, System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
        'dblValue = Val(Value)
        'Return dblValue
    End Function

    Public Function FormatScreenValues(ByVal Value As String) As Double
        Dim NewValue As Double

        If Value <> "" Then
            If Value.IndexOf(".") > -1 Then
                Value = Value.Replace(".", CompanyDecimalSeprator)
            End If
        Else
            Value = "0"
        End If

        'NewValue = CDbl(Value)
        NewValue = Val(Value)

        Return NewValue

        'Dim dblValue As Double
        'Value = Value.Replace(CompanyThousandSeprator, "")
        'Value = Value.Replace(CompanyDecimalSeprator, System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
        'dblValue = Val(Value)
        'Return dblValue

    End Function

    Public Function SetScreenValues(ByVal Value As String) As String

        If Value.IndexOf(CompanyDecimalSeprator) > -1 Then
            Value = Value.Replace(CompanyDecimalSeprator, ".")
        End If

        Return Value

    End Function

    Public Function SetDBValues(ByVal Value As String) As String

        If Value.IndexOf(CompanyDecimalSeprator) > -1 Then
            Value = Value.Replace(CompanyDecimalSeprator, ".")
        End If

        Return Value

    End Function

    Public Sub UpdateSupplierCatelog(aDocEntry As String)
        Try
            Dim oRec As SAPbobsCOM.Recordset
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oRec.DoQuery("Select T0.""CreateDate"" ""CreateDt"", * from ""@Z_OVPL"" T0 inner Join OUSR T1 on T0.""UserSign""=T1.""INTERNAL_K""   where ""DocEntry""='" & aDocEntry & "'")
            Dim strQuery1 As String = String.Empty
            If blnIsHanaDB Then
                strQuery1 = " Select T0.""DocNum"",T0.""CreateDate"" ""CreateDt"","
                strQuery1 &= " (Case When ifnull(""U_Z_CPrice"",0) = 0 Then ""U_Z_UPrice"" Else ""U_Z_CPrice"" End ) As ""U_Z_CPrice"", "
                strQuery1 &= " (Case When ifnull(""U_Z_CCurrency"",'') = '' Then ""U_Z_UCurrency"" Else ""U_Z_CCurrency"" End) As ""U_Z_CCurrency"", "
                strQuery1 &= " ""U_Z_UPrice"",""U_Z_UCurrency"",""U_NAME"",""U_Z_ItemCode"",""U_Z_CardCode"",""U_Z_AppStatus"" "
                strQuery1 &= " from ""@Z_OVPL"" T0 inner Join OUSR T1 on T0.""UserSign""=T1.""INTERNAL_K"" where ""DocEntry""='" & aDocEntry & "'"
            Else
                ' strQuery1 = "Select T0.""CreateDate"" ""CreateDt"",isnull(""U_Z_UPrice"",""U_Z_CPrice"") As ""U_Z_CPrice"",isnull(""U_Z_CCurrency"",""U_Z_CCurrency"") As ""U_Z_CCurrency"", * from ""@Z_OVPL"" T0 inner Join OUSR T1 on T0.""UserSign""=T1.""INTERNAL_K"" where ""DocEntry""='" & aDocEntry & "'"
            End If
            oRec.DoQuery(strQuery1)
            If oRec.RecordCount > 0 Then
                If oRec.Fields.Item("U_Z_AppStatus").Value = "A" Then
                    Dim dtDate As Date = oRec.Fields.Item("CreateDt").Value
                    Dim strQuery As String = "Update OSCN "
                    strQuery &= " set ""U_Z_UPrice""='" & oRec.Fields.Item("U_Z_CPrice").Value & "' "
                    strQuery &= " , ""U_Z_UCurrency""='" & oRec.Fields.Item("U_Z_CCurrency").Value & "' "
                    strQuery &= " , ""U_Z_CPrice""='" & oRec.Fields.Item("U_Z_UPrice").Value & "' "
                    strQuery &= " , ""U_Z_CCurrency""='" & oRec.Fields.Item("U_Z_UCurrency").Value & "' "
                    strQuery &= " , ""U_Z_ReqBy""='" & oRec.Fields.Item("U_NAME").Value & "' "
                    strQuery &= " , ""U_Z_ReqDate""='" & dtDate.ToString("yyyy-MM-dd") & "'"
                    strQuery &= " , ""U_Z_AppBy""='" & oApplication.Company.UserName & "' "
                    strQuery &= " , ""U_Z_AppDate""='" & Now.Date.ToString("yyyy-MM-dd") & "'"
                    strQuery &= " , ""U_Z_DocNum""='" & oRec.Fields.Item("DocNum").Value & "'"
                    strQuery &= "where ""ItemCode""='" & oRec.Fields.Item("U_Z_ItemCode").Value & "'"
                    strQuery &= "and ""CardCode""='" & oRec.Fields.Item("U_Z_CardCode").Value & "'"

                    oRec.DoQuery(strQuery)
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub GetItemPriceSupplierCatelog(strItemCode As String, strCardCode As String, ByRef strCurrency As String, ByRef dblPrice As Double)
        Try
            Dim oRec As SAPbobsCOM.Recordset
            oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strQuery As String
            strQuery = "Select ""U_Z_CPrice"",""U_Z_CCurrency"" from OSCN where ""ItemCode""='" & strItemCode & "'"
            strQuery &= "and ""CardCode""='" & strCardCode.Trim() & "'"
            oRec.DoQuery(strQuery)
            If oRec.RecordCount > 0 Then
                strCurrency = oRec.Fields.Item("U_Z_CCurrency").Value
                dblPrice = oRec.Fields.Item("U_Z_CPrice").Value
            Else
                strCurrency = ""
                dblPrice = 0
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub addPromotionReference(ByRef strCode As String)
        Try
            Dim oUDT As SAPbobsCOM.UserTable
            oUDT = oApplication.Company.UserTables.Item("Z_OPRF")
            Dim intCode As Integer = getMaxCode("@Z_OPRF", "Code")
            oUDT.Code = String.Format("{0:000000000}", intCode)
            oUDT.Name = String.Format("{0:000000000}", intCode)
            Dim intStatus As Integer = oUDT.Add()
            If intStatus = 0 Then
                strCode = String.Format("{0:000000000}", intCode)
            End If
        Catch ex As Exception
            oApplication.SBO_Application.SetStatusBarMessage(oApplication.Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, True)
        End Try
    End Sub

    Public Function CreateInventoryCountDocument(ByVal oForm As SAPbouiCOM.Form, ByVal strRef As String)
        Dim _retVal As Boolean = False
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim strQuery As String = String.Empty
        Try

            Dim oUpdateRecord As SAPbobsCOM.Recordset
            oUpdateRecord = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)


            If IO.Directory.Exists(System.Windows.Forms.Application.StartupPath.ToString() & "\Log") = False Then
                IO.Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath.ToString() & "\Log")
            End If
            Dim strFile As String = "\Log\Inventor_Creation_" + System.DateTime.Now.ToString("yyyyMMddmmss") + ".txt"

            strQuery = "Select T1.""U_Z_CntDate"",T0.""U_Z_ItmCode"",SUM(T0.""U_Z_IQty"") ""U_Z_IQty"",T0.""U_Z_IUOM"",T0.""U_Z_WareHouse"" From ""@Z_ICT1"" T0 JOIN ""@Z_OICT"" T1 On T0.""DocEntry"" = T1.""DocEntry"" "
            strQuery += " Where T1.""DocEntry"" = '" & strRef & "'"
            If Not blnIsHanaDB Then
                strQuery += " And ISNULL(""U_Z_ICRef"",'') = '' "
            Else
                strQuery += " And IFNULL(""U_Z_ICRef"",'') = '' "
            End If
            strQuery += " And T0.""U_Z_ItmCode"" Is Not Null Group By T0.""U_Z_ItmCode"",T1.""U_Z_CntDate"",T0.""U_Z_IUOM"",T0.""U_Z_WareHouse"" "

            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                Dim blnRowExist As Boolean = False

                Dim oCS As SAPbobsCOM.CompanyService = oApplication.Company.GetCompanyService()
                Dim oICS As SAPbobsCOM.InventoryCountingsService = oCS.GetBusinessService(ServiceTypes.InventoryCountingsService)
                Dim oIC As SAPbobsCOM.InventoryCounting = oICS.GetDataInterface(InventoryCountingsServiceDataInterfaces.icsInventoryCounting)

                oIC.CountDate = oRecordSet.Fields.Item("U_Z_CntDate").Value
                oIC.SingleCounterType = CounterTypeEnum.ctUser
                oIC.SingleCounterID = oApplication.Company.UserSignature
                oIC.UserFields.Item("U_Z_ICTREF").Value = strRef
                'oIC.UserFields.Item("U_Z_ICTREFL").Value = oRecordSet.Fields.Item("LineId").Value.ToString
                Dim oICLS As SAPbobsCOM.InventoryCountingLines = oIC.InventoryCountingLines

                While Not oRecordSet.EoF
                    If 1 = 1 Then

                        blnRowExist = True
                        Dim intRow As Integer = 0
                        blnRowExist = True

                        Dim oICL As SAPbobsCOM.InventoryCountingLine = oICLS.Add

                        oICL.ItemCode = oRecordSet.Fields.Item("U_Z_ItmCode").Value
                        oICL.Counted = BoYesNoEnum.tYES
                        oICL.CountedQuantity = oRecordSet.Fields.Item("U_Z_IQty").Value
                        oICL.WarehouseCode = oRecordSet.Fields.Item("U_Z_WareHouse").Value
                        oICL.UoMCode = oRecordSet.Fields.Item("U_Z_IUOM").Value

                    End If
                    oRecordSet.MoveNext()
                End While


                If blnRowExist Then
                    Try
                        Dim oICP As SAPbobsCOM.InventoryCountingParams = oICS.Add(oIC)
                        If oICP.DocumentEntry > 0 Then
                            strQuery = "Update ""@Z_ICT1"" Set ""U_Z_ICRef"" = '" & oICP.DocumentEntry.ToString & "'"
                            strQuery += ",""U_Z_Status"" = 'C' "
                            strQuery += " Where ""DocEntry"" = '" & strRef & "'"
                            'strQuery += " And ""LineId"" = '" & oRecordSet.Fields.Item("LineId").Value.ToString() & "'"
                            oUpdateRecord.DoQuery(strQuery)
                            Trace_ProcessCall("Inventory Count Ref : " & strRef & " -->Success", strFile)
                        Else
                            Trace_ProcessCall("Inventory Count Ref : " & strRef & "-->ERROR ERRORCODE :" & oApplication.Company.GetLastErrorCode().ToString() & " ERRORDESC : " & oApplication.Company.GetLastErrorDescription().ToString(), strFile)
                        End If
                    Catch ex As Exception
                        Trace_ProcessCall("Inventory Count Ref : " & strRef & "-->ERROR : " & ex.Message, strFile)
                    End Try
                End If

                Dim strPath As String = System.Windows.Forms.Application.StartupPath.ToString() & "\Log" & strFile
                If (File.Exists(strPath)) Then
                    System.Diagnostics.Process.Start(strPath)
                End If

            End If
            Return _retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function checkAllDocumentStatus(ByVal oForm As SAPbouiCOM.Form, ByVal strRef As String)
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oUpdateRecord As SAPbobsCOM.Recordset
            oUpdateRecord = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strQuery As String = String.Empty
            strQuery = "Select Count(""DocEntry"") From ""@Z_ICT1"" "
            strQuery += " Where ""DocEntry"" = '" & strRef & "'"
            strQuery += " And ""U_Z_Status"" = 'O' "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                Dim intCount As Integer = CInt(oRecordSet.Fields.Item(0).Value)
                If intCount = 0 Then
                    strQuery = "Update ""@Z_OICT"" Set ""U_Z_Status"" = 'C' "
                    strQuery += " Where ""DocEntry"" = '" & strRef & "'"
                    oUpdateRecord.DoQuery(strQuery)
                End If
            End If
        Catch ex As Exception

        End Try
    End Function

    Public Sub Trace_ProcessCall(ByVal strContent As String, ByVal strFile As String)
        Try
            Dim strPath As String = System.Windows.Forms.Application.StartupPath.ToString() + strFile
            If Not File.Exists(strPath) Then
                Dim fileStream As FileStream
                fileStream = New FileStream(strPath, FileMode.Create, FileAccess.Write)
                Dim sw As New StreamWriter(fileStream)
                sw.BaseStream.Seek(0, SeekOrigin.End)
                sw.WriteLine(strContent)
                sw.Flush()
                sw.Close()
            Else
                Dim fileStream As FileStream
                fileStream = New FileStream(strPath, FileMode.Append, FileAccess.Write)
                Dim sw As New StreamWriter(fileStream)
                sw.BaseStream.Seek(0, SeekOrigin.End)
                sw.WriteLine(strContent)
                sw.Flush()
                sw.Close()
            End If
        Catch ex As Exception
            'Throw ex
        End Try
    End Sub

    Public Function changeStatus(ByVal oForm As SAPbouiCOM.Form, ByVal strStatus As String) As Boolean
        Dim _retVal As Boolean = False
        Try
            Dim strDocEntry As String = CType(oForm.Items.Item("6_").Specific, SAPbouiCOM.EditText).Value
            Dim oGeneralService As SAPbobsCOM.GeneralService
            Dim oGeneralData As SAPbobsCOM.GeneralData
            Dim oCompanyService As SAPbobsCOM.CompanyService
            Dim oGeneralDataParams As SAPbobsCOM.GeneralDataParams
            Dim strQuery As String = String.Empty
            oCompanyService = oApplication.Company.GetCompanyService()
            Try
                oGeneralService = oCompanyService.GetGeneralService("Z_OICT")
                oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
                oGeneralDataParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
                oGeneralDataParams.SetProperty("DocEntry", strDocEntry)
                oGeneralData = oGeneralService.GetByParams(oGeneralDataParams)
                oGeneralData.SetProperty("U_Z_Status", strStatus)
                oGeneralService.Update(oGeneralData)
                _retVal = True
            Catch ex As Exception
                Throw ex
            End Try
        Catch ex As Exception
            Throw ex
        End Try
        Return _retVal
    End Function

End Class
