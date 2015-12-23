﻿Public Class clsRebatePosting
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
    Private InvBaseDocNo, strQuery As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Invoice Then
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
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                If oForm.TypeEx = frm_Invoice Then
                    Dim oobj As SAPbobsCOM.Documents
                    oobj = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                    If oobj.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                        If oobj.Cancelled = SAPbobsCOM.BoYesNoEnum.tYES Then
                            ' CreditAPCreditNote(oobj.DocEntry)
                            CancelJournal(oobj.DocEntry)
                        Else
                            ' CreateAPInvoice(oobj.DocEntry)
                            CreateJournal(oobj.DocEntry)
                        End If
                    End If

                End If

                If oForm.TypeEx = frm_ARCreditNote Then
                    Dim oobj As SAPbobsCOM.Documents
                    oobj = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                    If oobj.Browser.GetByKeys(BusinessObjectInfo.ObjectKey) Then
                        If oobj.Cancelled = SAPbobsCOM.BoYesNoEnum.tYES Then
                            '  CreditAPCreditNote_ARCreditNote(oobj.DocEntry)
                            CancelJournal_CreditNote(oobj.DocEntry)
                        Else
                            CreateJournal_CreditMemo(oobj.DocEntry)
                        End If
                    End If

                End If
            End If

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub ApplyRebateAmount(ByVal aDocEntry As Integer)
        Dim oAPInv As SAPbobsCOM.Documents
        Dim oTest, otest1, oRec, oTemp1, oTest2 As SAPbobsCOM.Recordset
        Dim strCardCode, strItemCode, strQuery, strDistRule As String
        Dim dblLineTotal, dblCommission, dblMarketing, dblComPercentage, dblMarketingPercentage, dblSupComm, dblSupCommPreLicense, dblItemCost, dblUnitPrice, dblQty, dblHubTotal As Double
        Dim dtPostingDate As Date
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oAPInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        Dim blnHubWhs As Boolean = False
        ' oRec.DoQuery("Select T1.""CardCode"" ""Card"", * from INV1 T0 Inner Join OINV T1 on T1.""DocEntry""=T0.""DocEntry"" inner Join OITM T2 on T2.""ItemCode""=T0.""ItemCode"" where T0.""DocEntry""=" & aDocEntry)
        oRec.DoQuery("Select T0.""BaseCard"" ""Card"", * from INV1 T0 Inner Join OINV T1 on T1.""DocEntry""=T0.""DocEntry"" where T0.""DocEntry""=" & aDocEntry)
        For intRow As Integer = 0 To oRec.RecordCount - 1
            strCardCode = oRec.Fields.Item("Card").Value
            strItemCode = oRec.Fields.Item("ItemCode").Value
            dtPostingDate = oRec.Fields.Item("DocDate").Value
            dblUnitPrice = oRec.Fields.Item("PriceBefDi").Value
            dblQty = oRec.Fields.Item("Quantity").Value

            oTemp1.DoQuery("Select * from OITM where ""ItemCode""='" & strItemCode & "'")
            dblSupComm = oTemp1.Fields.Item("U_Z_SupCom").Value
            dblSupCommPreLicense = oTemp1.Fields.Item("U_Z_SupComPre").Value

            If blnIsHanaDB Then
                strQuery = "Select ifnull(""U_Z_Type"",'R') from OWHS where ""WhsCode""='" & oRec.Fields.Item("WhsCode").Value & "'"
            Else
                strQuery = "Select isnull(""U_Z_Type"",'R') from OWHS where ""WhsCode""='" & oRec.Fields.Item("WhsCode").Value & "'"
            End If
            otest1.DoQuery(strQuery)
            If otest1.Fields.Item(0).Value = "H" Then
                blnHubWhs = True
                otest1.DoQuery("Select ""AvgPrice"" from OITW where ""WhsCode""='" & oRec.Fields.Item("WhsCode").Value & "' and ""ItemCode""='" & strItemCode & "'")
                'If otest1.RecordCount > 0 Then
                '    dblItemCost = otest1.Fields.Item(0).Value
                'Else
                '    dblItemCost = 0
                'End If
                dblItemCost = oRec.Fields.Item("StockPrice").Value
            Else
                dblItemCost = 0

            End If

            strDistRule = oRec.Fields.Item("OcrCode").Value ' & ";" & oRec.Fields.Item("OcrCode2").Value & ";" & oRec.Fields.Item("OcrCode3").Value & ";" & oRec.Fields.Item("OcrCode4").Value & ";" & oRec.Fields.Item("OcrCode5").Value
            If blnIsHanaDB Then
                strQuery = "Select T0.""ItemCode"",T1.""U_Z_Comm"",T1.""U_Z_MarkReb"",T1.""U_Z_RegStatus"" From OSPP T0 LEFT OUTER JOIN SPP1 T1 On T0.""ItemCode"" = T1.""ItemCode"" And T0.""CardCode"" = T1.""CardCode""  Where T1.""U_Z_OcrCode""='" & strDistRule & "' and T0.""CardCode"" = '" & strCardCode & "' And T0.""ItemCode"" = '" & strItemCode & "' and '" & dtPostingDate.ToString("yyyy-MM-dd") & "' between  ifnull(T1.""FromDate"",NOW()) and ifnull(T1.""ToDate"",NOW())"
            Else
                strQuery = "Select T0.""ItemCode"",T1.""U_Z_Comm"",T1.""U_Z_MarkReb"",T1.""U_Z_RegStatus"" From OSPP T0 LEFT OUTER JOIN SPP1 T1 On T0.""ItemCode"" = T1.""ItemCode"" And T0.""CardCode"" = T1.""CardCode""  Where T1.""U_Z_OcrCode""='" & strDistRule & "' and T0.""CardCode"" = '" & strCardCode & "' And T0.""ItemCode"" = '" & strItemCode & "' and '" & dtPostingDate.ToString("yyyy-MM-dd") & "' between  isnull(T1.""FromDate"",GetDate()) and isnull(T1.""ToDate"",GetDate())"
            End If

            otest1.DoQuery(strQuery)
            If otest1.RecordCount > 0 Then
                Dim oTe As SAPbobsCOM.Recordset
                oTe = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If blnIsHanaDB Then
                    strQuery = "Select ifnull(""U_Z_ComBase"",0) from OCRD where  ""CardCode""='" & strCardCode & "'"
                Else
                    strQuery = "Select isnull(""U_Z_ComBase"",0) from OCRD where  ""CardCode""='" & strCardCode & "'"
                End If
                oTe.DoQuery(strQuery)
                If oTe.Fields.Item(0).Value > 0 Then
                    dblUnitPrice = (dblUnitPrice * oTe.Fields.Item(0).Value / 100)
                    dblLineTotal = dblUnitPrice * dblQty
                Else
                    dblLineTotal = oRec.Fields.Item("LineTotal").Value
                End If
                '   dblLineTotal = oRec.Fields.Item("LineTotal").Value
                dblCommission = otest1.Fields.Item(1).Value
                dblMarketing = otest1.Fields.Item(2).Value
                dblComPercentage = dblCommission
                dblMarketingPercentage = dblMarketing
                If dblCommission <> 0 Then
                    dblCommission = dblLineTotal * dblCommission / 100
                Else
                    dblCommission = 0
                    dblComPercentage = 0
                End If
                If dblMarketing <> 0 Then
                    dblMarketing = dblLineTotal * dblMarketing / 100
                Else
                    dblMarketing = 0
                    dblMarketingPercentage = 0
                End If
                If blnHubWhs = True Then
                    If otest1.Fields.Item("U_Z_RegStatus").Value = "R" Or otest1.Fields.Item("U_Z_RegStatus").Value = "N" Then
                        Dim s As String = "Update INV1 set ""U_Z_ItemCost""='" & dblItemCost & "', ""U_Z_SupCom1""='" & dblSupComm & "', ""U_Z_RegStatus""='" & otest1.Fields.Item("U_Z_RegStatus").Value & "', ""U_Z_Comm_Per""='" & dblComPercentage & "',""U_Z_MarkReb_Per""='" & dblMarketingPercentage & "', ""U_Z_Comm""='" & dblCommission & "',""U_Z_MarkReb""='" & dblMarketing & "', ""U_Z_IsComm""='Y' where ""DocEntry""=" & oRec.Fields.Item("DocEntry").Value & " and ""LineNum""=" & oRec.Fields.Item("LineNum").Value
                        dblHubTotal = (dblUnitPrice * dblQty * dblSupComm / 100)
                        dblHubTotal = dblHubTotal - (dblQty * dblItemCost)

                        oTemp1.DoQuery("Update INV1 set ""U_Z_Accrual""='" & dblHubTotal & "', ""U_Z_ItemCost""='" & dblItemCost & "', ""U_Z_SupCom1""='" & dblSupComm & "', ""U_Z_RegStatus""='" & otest1.Fields.Item("U_Z_RegStatus").Value & "', ""U_Z_Comm_Per""='" & dblComPercentage & "',""U_Z_MarkReb_Per""='" & dblMarketingPercentage & "', ""U_Z_Comm""='" & dblCommission & "',""U_Z_MarkReb""='" & dblMarketing & "', ""U_Z_IsComm""='Y' where ""DocEntry""=" & oRec.Fields.Item("DocEntry").Value & " and ""LineNum""=" & oRec.Fields.Item("LineNum").Value)
                    ElseIf otest1.Fields.Item("U_Z_RegStatus").Value = "L" Then
                        dblHubTotal = (dblUnitPrice * dblQty * dblSupCommPreLicense / 100)
                        dblHubTotal = dblHubTotal - (dblQty * dblItemCost)
                        oTemp1.DoQuery("Update INV1 set ""U_Z_Accrual""='" & dblHubTotal & "',  ""U_Z_ItemCost""='" & dblItemCost & "', ""U_Z_SupCom1""='" & dblSupCommPreLicense & "', ""U_Z_RegStatus""='" & otest1.Fields.Item("U_Z_RegStatus").Value & "', ""U_Z_Comm_Per""='" & dblComPercentage & "',""U_Z_MarkReb_Per""='" & dblMarketingPercentage & "', ""U_Z_Comm""='" & dblCommission & "',""U_Z_MarkReb""='" & dblMarketing & "', ""U_Z_IsComm""='Y' where ""DocEntry""=" & oRec.Fields.Item("DocEntry").Value & " and ""LineNum""=" & oRec.Fields.Item("LineNum").Value)
                    Else
                        oTemp1.DoQuery("Update INV1 set  ""U_Z_Comm_Per""='" & dblComPercentage & "',""U_Z_MarkReb_Per""='" & dblMarketingPercentage & "', ""U_Z_Comm""='" & dblCommission & "',""U_Z_MarkReb""='" & dblMarketing & "', ""U_Z_IsComm""='Y' where ""DocEntry""=" & oRec.Fields.Item("DocEntry").Value & " and ""LineNum""=" & oRec.Fields.Item("LineNum").Value)
                    End If
                Else
                    oTemp1.DoQuery("Update INV1 set  ""U_Z_Comm_Per""='" & dblComPercentage & "',""U_Z_MarkReb_Per""='" & dblMarketingPercentage & "', ""U_Z_Comm""='" & dblCommission & "',""U_Z_MarkReb""='" & dblMarketing & "', ""U_Z_IsComm""='Y' where ""DocEntry""=" & aDocEntry & " and ""LineNum""=" & oRec.Fields.Item("LineNum").Value)
                End If
            Else
                dblCommission = 0
                dblMarketing = 0
                oTemp1.DoQuery("Update INV1 set  ""U_Z_Comm_Per""='" & dblComPercentage & "',""U_Z_MarkReb_Per""='" & dblMarketingPercentage & "',""U_Z_Comm""='" & dblCommission & "',""U_Z_MarkReb""='" & dblMarketing & "', ""U_Z_IsComm""='N' where ""DocEntry""=" & oRec.Fields.Item("DocEntry").Value & " and ""LineNum""=" & oRec.Fields.Item("LineNum").Value)
            End If
            oRec.MoveNext()
        Next
    End Sub

    Public Sub ApplyRebateAmount_CreditNote(ByVal aDocEntry As Integer)
        Dim oAPInv As SAPbobsCOM.Documents
        Dim oTest, otest1, oRec, oTemp1, oTest2 As SAPbobsCOM.Recordset
        Dim strCardCode, strItemCode, strQuery, strDistRule As String
        Dim dblLineTotal, dblCommission, dblMarketing, dblComPercentage, dblMarketingPercentage As Double
        Dim dtPostingDate As Date
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oAPInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        oRec.DoQuery("Select T0.""BaseCard"" ""Card"", * from RIN1 T0 Inner Join ORIN T1 on T1.""DocEntry""=T0.""DocEntry"" where T0.""DocEntry""=" & aDocEntry)
        For intRow As Integer = 0 To oRec.RecordCount - 1
            strCardCode = oRec.Fields.Item("Card").Value
            strItemCode = oRec.Fields.Item("ItemCode").Value
            dtPostingDate = oRec.Fields.Item("DocDate").Value
            strDistRule = oRec.Fields.Item("OcrCode").Value ' & ";" & oRec.Fields.Item("OcrCode2").Value & ";" & oRec.Fields.Item("OcrCode3").Value & ";" & oRec.Fields.Item("OcrCode4").Value & ";" & oRec.Fields.Item("OcrCode5").Value
            'strQuery = "Select T0.""ItemCode"",T1.""U_Z_Comm"",T1.""U_Z_MarkReb"" From OSPP T0 LEFT OUTER JOIN SPP1 T1 On T0.""ItemCode"" = T1.""ItemCode"" And T0.""CardCode"" = T1.""CardCode""  Where T1.""U_Z_OcrCode""='" & strDistRule & "' and T0.""CardCode"" = '" & strCardCode & "' And T0.""ItemCode"" = '" & strItemCode & "' and '" & dtPostingDate.ToString("yyyy-MM-dd") & "' between  ifnull(T1.""FromDate"",NOW()) and ifnull(T1.""ToDate"",NOW())"

            If blnIsHanaDB Then
                strQuery = "Select T0.""ItemCode"",T1.""U_Z_Comm"",T1.""U_Z_MarkReb"" From OSPP T0 LEFT OUTER JOIN SPP1 T1 On T0.""ItemCode"" = T1.""ItemCode"" And T0.""CardCode"" = T1.""CardCode""  Where T1.""U_Z_OcrCode""='" & strDistRule & "' and T0.""CardCode"" = '" & strCardCode & "' And T0.""ItemCode"" = '" & strItemCode & "' and '" & dtPostingDate.ToString("yyyy-MM-dd") & "' between  ifnull(T1.""FromDate"",NOW()) and ifnull(T1.""ToDate"",NOW())"
            Else
                strQuery = "Select T0.""ItemCode"",T1.""U_Z_Comm"",T1.""U_Z_MarkReb"" From OSPP T0 LEFT OUTER JOIN SPP1 T1 On T0.""ItemCode"" = T1.""ItemCode"" And T0.""CardCode"" = T1.""CardCode""  Where T1.""U_Z_OcrCode""='" & strDistRule & "' and T0.""CardCode"" = '" & strCardCode & "' And T0.""ItemCode"" = '" & strItemCode & "' and '" & dtPostingDate.ToString("yyyy-MM-dd") & "' between  isnull(T1.""FromDate"",GetDate()) and isnull(T1.""ToDate"",GetDate())"
            End If

            'otest1.DoQuery(strQuery)
            If oRec.Fields.Item("U_Z_IsComm").Value = "Y" And (oRec.Fields.Item("U_Z_Comm_Per").Value <> 0 Or oRec.Fields.Item("U_Z_MarkReb_Per").Value <> 0) Then
                dblLineTotal = oRec.Fields.Item("LineTotal").Value
                dblCommission = oRec.Fields.Item("U_Z_Comm_Per").Value
                dblMarketing = oRec.Fields.Item("U_Z_MarkReb_Per").Value
                dblComPercentage = dblCommission
                dblMarketingPercentage = dblMarketing
                If dblCommission <> 0 Then
                    dblCommission = dblLineTotal * dblCommission / 100
                Else
                    dblCommission = 0
                    dblComPercentage = 0
                End If
                If dblMarketing <> 0 Then
                    dblMarketing = dblLineTotal * dblMarketing / 100
                Else
                    dblMarketing = 0
                    dblMarketingPercentage = 0
                End If
                otest1.DoQuery("Update RIN1 set ""U_Z_Comm_Per""='" & dblComPercentage & "',""U_Z_MarkReb_Per""='" & dblMarketingPercentage & "', ""U_Z_Comm""='" & dblCommission & "',""U_Z_MarkReb""='" & dblMarketing & "', ""U_Z_IsComm""='Y' where ""DocEntry""=" & oRec.Fields.Item("DocEntry").Value & " and ""LineNum""=" & oRec.Fields.Item("LineNum").Value)
            Else
                dblCommission = 0
                dblMarketing = 0
                otest1.DoQuery("Update RIN1 set ""U_Z_Comm_Per""='" & dblComPercentage & "',""U_Z_MarkReb_Per""='" & dblMarketingPercentage & "',""U_Z_Comm""='" & dblCommission & "',""U_Z_MarkReb""='" & dblMarketing & "', ""U_Z_IsComm""='N' where ""DocEntry""=" & oRec.Fields.Item("DocEntry").Value & " and ""LineNum""=" & oRec.Fields.Item("LineNum").Value)
            End If
            oRec.MoveNext()
        Next
    End Sub

    Public Function CreateJournal_CreditMemo(ByVal DocNum As String) As Boolean
        Dim oAPInv As SAPbobsCOM.JournalEntries
        Dim oTest, otest1, oRec, oTemp1, oTest2 As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oAPInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
        Dim strCreditAc, strDebitAc As String

        Try
            strCreditAc = ""
            strDebitAc = ""
            Dim strDocNum, strCardcode As String
            If 1 = 1 Then 'strCreditAc <> "" And strDebitAc <> "" Then
                strQuery = "Select * from ORIN where ""DocEntry""=" & DocNum & " and ""U_Z_OYVP""='Y'"
                oTest.DoQuery(strQuery)
                If oTest.RecordCount > 0 Then
                    Dim dblPercentage, dblDocTotal As Double
                    strQuery = "Select * from OCRD where ""CardCode""='" & oTest.Fields.Item("CardCode").Value & "' and ""U_Z_OYVP"">0"
                    otest1.DoQuery(strQuery)
                    If otest1.RecordCount > 0 Then
                        dblPercentage = otest1.Fields.Item("U_Z_OYVP").Value
                        strDocNum = oTest.Fields.Item("DocNum").Value
                        strCardcode = oTest.Fields.Item("CardCode").Value
                        oAPInv.TaxDate = oTest.Fields.Item("DocDate").Value
                        oAPInv.DueDate = oTest.Fields.Item("DocDueDate").Value
                        oAPInv.ReferenceDate = oTest.Fields.Item("DocDate").Value
                        oAPInv.UserFields.Fields.Item("U_Z_ARInvoice").Value = oTest.Fields.Item("DocNum").Value.ToString
                        oAPInv.UserFields.Fields.Item("U_Z_BaseEntry").Value = DocNum
                        oAPInv.Memo = "Rebate Posting Based on A/R CreditNote  : " & oTest.Fields.Item("DocNum").Value.ToString
                        Dim blnLineExists As Boolean = False

                        Dim dbLComm, dblMarketing As Double
                        Dim dblDebit As Double = 0
                        Dim strCountry As String = ""
                        Dim intLineCount As Integer = 0
                        'For intloop As Integer = 0 To oTemp1.RecordCount - 1
                        dbLComm = 0
                        dblMarketing = 0
                        ' dbLComm = oTemp1.Fields.Item("U_Z_Comm").Value
                        dblDocTotal = oTest.Fields.Item("DocTotal").Value
                        dblDocTotal = dblDocTotal * dblPercentage / 100
                        If dblDocTotal <> 0 Then
                            oTest.DoQuery("Select * from ""@Z_OYVP""")
                            dbLComm = dblDocTotal
                            If oTest.RecordCount > 0 Then
                                strCreditAc = oTest.Fields.Item("U_Z_Debit").Value
                                strDebitAc = oTest.Fields.Item("U_Z_Credit").Value
                            Else
                                strDebitAc = ""
                                strCreditAc = ""
                            End If
                            If strDebitAc <> "" And strDebitAc <> "" Then
                                'Credit Entry
                                If intLineCount > 0 Then
                                    oAPInv.Lines.Add()
                                End If
                                oAPInv.Lines.SetCurrentLine(intLineCount)
                                dblDebit = dblDebit + dbLComm
                                oAPInv.Lines.AccountCode = oApplication.Utilities.getAccountCode(strCreditAc)
                                oAPInv.Lines.Credit = dbLComm
                                oAPInv.Lines.Reference1 = strDocNum
                                oAPInv.Lines.Reference2 = strCardcode
                                intLineCount = intLineCount + 1
                                blnLineExists = True
                                'Debit Entry
                                If intLineCount > 0 Then
                                    oAPInv.Lines.Add()
                                End If
                                oAPInv.Lines.SetCurrentLine(intLineCount)
                                oAPInv.Lines.AccountCode = oApplication.Utilities.getAccountCode(strDebitAc)
                                oAPInv.Lines.Debit = dbLComm
                                oAPInv.Lines.Reference1 = strDocNum
                                oAPInv.Lines.Reference2 = strCardcode
                                intLineCount = intLineCount + 1
                                blnLineExists = True
                            End If
                        End If
                        '  oTemp1.MoveNext()
                        ' Next
                        If blnLineExists = True Then
                            If oAPInv.Add <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            Else
                                oTest2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim strDocNum1 As String
                                oApplication.Company.GetNewObjectCode(strDocNum1)
                                oAPInv.GetByKey(CInt(strDocNum1))
                                strDocNum1 = oAPInv.JdtNum
                                strQuery = "Update ORIN set ""U_Z_JournalRef""='" & strDocNum1 & "' where ""DocEntry""=" & DocNum
                                oTest2.DoQuery(strQuery)
                            End If
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Public Function CreateJournal(ByVal DocNum As String) As Boolean
        Dim oAPInv As SAPbobsCOM.JournalEntries
        Dim oTest, otest1, oRec, oTemp1, oTest2 As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oAPInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
        Dim strCreditAc, strDebitAc, strTaxDebitAC As String

        Try
            strCreditAc = ""
            strDebitAc = ""
            strTaxDebitAC = ""
            Dim strDocNum, strCardcode As String
            If 1 = 1 Then 'strCreditAc <> "" And strDebitAc <> "" Then
                strQuery = "Select * from OINV where ""DocEntry""=" & DocNum & " and ""U_Z_OYVP""='Y'"
                oTest.DoQuery(strQuery)
                If oTest.RecordCount > 0 Then
                    Dim dblPercentage, dblDocTotal As Double
                    strQuery = "Select * from OCRD where ""CardCode""='" & oTest.Fields.Item("CardCode").Value & "' and ""U_Z_OYVP"">0"
                    otest1.DoQuery(strQuery)
                    If otest1.RecordCount > 0 Then
                        dblPercentage = otest1.Fields.Item("U_Z_OYVP").Value
                        strDocNum = oTest.Fields.Item("DocNum").Value
                        strCardcode = oTest.Fields.Item("CardCode").Value
                        oAPInv.TaxDate = oTest.Fields.Item("DocDate").Value
                        oAPInv.DueDate = oTest.Fields.Item("DocDueDate").Value
                        oAPInv.ReferenceDate = oTest.Fields.Item("DocDate").Value
                        oAPInv.UserFields.Fields.Item("U_Z_ARInvoice").Value = oTest.Fields.Item("DocNum").Value.ToString
                        oAPInv.UserFields.Fields.Item("U_Z_BaseEntry").Value = DocNum
                        oAPInv.Memo = "Rebate Posting Based on A/R Invoice  : " & oTest.Fields.Item("DocNum").Value.ToString
                        Dim blnLineExists As Boolean = False
                      
                        Dim dbLComm, dblMarketing, dblTax, dblTotal As Double
                        Dim dblDebit As Double = 0
                        Dim strCountry As String = ""
                        Dim intLineCount As Integer = 0
                        'For intloop As Integer = 0 To oTemp1.RecordCount - 1
                        dbLComm = 0
                        dblMarketing = 0
                        ' dbLComm = oTemp1.Fields.Item("U_Z_Comm").Value
                        dblTotal = oTest.Fields.Item("DocTotal").Value
                        dblTax = oTest.Fields.Item("VatSum").Value
                        dblDocTotal = dblTotal - dblTax ' oTest.Fields.Item("DocTotal").Value
                        dblDocTotal = dblDocTotal * dblPercentage / 100
                        dblTax = dblTax * dblPercentage / 100

                        If dblDocTotal <> 0 Then
                            oTest.DoQuery("Select * from ""@Z_OYVP""")
                            dbLComm = dblDocTotal
                            If oTest.RecordCount > 0 Then
                                strDebitAc = oTest.Fields.Item("U_Z_Debit").Value
                                strCreditAc = oTest.Fields.Item("U_Z_Credit").Value
                                strTaxDebitAC = oTest.Fields.Item("U_Z_TaxDebit").Value
                            Else
                                strDebitAc = ""
                                strCreditAc = ""
                            End If
                            If strDebitAc <> "" And strDebitAc <> "" Then
                                'Credit Entry
                                If intLineCount > 0 Then
                                    oAPInv.Lines.Add()
                                End If
                                oAPInv.Lines.SetCurrentLine(intLineCount)
                                dblDebit = dblDebit + dbLComm
                                oAPInv.Lines.AccountCode = oApplication.Utilities.getAccountCode(strCreditAc)
                                oAPInv.Lines.Credit = dbLComm
                                oAPInv.Lines.Reference1 = strDocNum
                                oAPInv.Lines.Reference2 = strCardcode
                                intLineCount = intLineCount + 1
                                blnLineExists = True
                                'Debit Entry
                                If intLineCount > 0 Then
                                    oAPInv.Lines.Add()
                                End If
                                oAPInv.Lines.SetCurrentLine(intLineCount)
                                oAPInv.Lines.AccountCode = oApplication.Utilities.getAccountCode(strDebitAc)
                                oAPInv.Lines.Debit = dbLComm
                                oAPInv.Lines.Reference1 = strDocNum
                                oAPInv.Lines.Reference2 = strCardcode
                                intLineCount = intLineCount + 1
                                blnLineExists = True
                            End If
                        End If

                        'tax posting


                        If dblTax <> 0 Then
                            oTest.DoQuery("Select * from ""@Z_OYVP""")
                            dbLComm = dblTax
                            If oTest.RecordCount > 0 Then
                                strDebitAc = oTest.Fields.Item("U_Z_Debit").Value
                                strCreditAc = oTest.Fields.Item("U_Z_Credit").Value
                                strTaxDebitAC = oTest.Fields.Item("U_Z_TaxDebit").Value
                            Else
                                strDebitAc = ""
                                strCreditAc = ""
                                strTaxDebitAC = ""
                            End If

                            If strTaxDebitAC <> "" And strCreditAc <> "" Then
                                'Credit Entry
                                If intLineCount > 0 Then
                                    oAPInv.Lines.Add()
                                End If
                                oAPInv.Lines.SetCurrentLine(intLineCount)
                                dblDebit = dblTax
                                oAPInv.Lines.AccountCode = oApplication.Utilities.getAccountCode(strCreditAc)
                                oAPInv.Lines.Credit = dblTax
                                oAPInv.Lines.Reference1 = strDocNum
                                oAPInv.Lines.Reference2 = strCardcode
                                intLineCount = intLineCount + 1
                                blnLineExists = True

                                'Debit Entry
                                If intLineCount > 0 Then
                                    oAPInv.Lines.Add()
                                End If
                                oAPInv.Lines.SetCurrentLine(intLineCount)
                                oAPInv.Lines.AccountCode = oApplication.Utilities.getAccountCode(strTaxDebitAC)
                                oAPInv.Lines.Debit = dblTax
                                oAPInv.Lines.Reference1 = strDocNum
                                oAPInv.Lines.Reference2 = strCardcode
                                intLineCount = intLineCount + 1
                                blnLineExists = True
                            End If
                        End If
                        '  oTemp1.MoveNext()
                        ' Next
                        If blnLineExists = True Then
                            If oAPInv.Add <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            Else
                                oTest2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim strDocNum1 As String
                                oApplication.Company.GetNewObjectCode(strDocNum1)
                                oAPInv.GetByKey(CInt(strDocNum1))
                                strDocNum1 = oAPInv.JdtNum
                                strQuery = "Update OINV set ""U_Z_JournalRef""='" & strDocNum1 & "' where ""DocEntry""=" & DocNum
                                oTest2.DoQuery(strQuery)
                            End If
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Public Function CreateAPInvoice(ByVal DocNum As String) As Boolean
        Dim oAPInv As SAPbobsCOM.Documents
        Dim oTest, otest1, oRec, oTemp1, oTest2 As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oAPInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        Try
            ApplyRebateAmount(DocNum)
            strQuery = "Select * from OINV where ""DocEntry""=" & DocNum & ""
            oTest.DoQuery(strQuery)
            If oTest.RecordCount > 0 Then
                strQuery = "Select * from OCRD where ""CardCode""='" & oTest.Fields.Item("CardCode").Value & "' and ""U_Z_ComRePay""<>''"
                otest1.DoQuery(strQuery)
                If otest1.RecordCount > 0 Then
                    strQuery = "Select * from OCRD where ""CardCode""='" & otest1.Fields.Item("U_Z_ComRePay").Value & "' and ""CardType""='S'"
                    oRec.DoQuery(strQuery)
                    If oRec.RecordCount > 0 Then
                        oAPInv.DocDate = oTest.Fields.Item("DocDate").Value
                        oAPInv.DocDueDate = oTest.Fields.Item("DocDueDate").Value
                        oAPInv.CardCode = otest1.Fields.Item("U_Z_ComRePay").Value
                        oAPInv.UserFields.Fields.Item("U_Z_ARInvoice").Value = oTest.Fields.Item("DocNum").Value.ToString
                        oAPInv.UserFields.Fields.Item("U_Z_BaseEntry").Value = DocNum
                        oAPInv.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                        oAPInv.NumAtCard = "AR Invoice No : " & oTest.Fields.Item("DocNum").Value.ToString
                        oAPInv.Comments = "Rebate/Commission posting based on A/R Invoice No  : " & oTest.Fields.Item("DocNum").Value.ToString
                        Dim blnLineExists As Boolean = False
                        '   strQuery = "Select sum(""U_Z_Comm"") 'U_Z_Comm',sum(""U_Z_MarkReb"") 'U_Z_MarkReb',""OcrCode"",""OcrCode2"",""OcrCode3"",""OcrCode4"",""OcrCode5"" from INV1 where ""DocEntry""='" & DocNum & "' and ""U_Z_IsComm""='Y' group by ""OcrCode"",""OcrCode2"",""OcrCode3"",""OcrCode4"",""OcrCode5"""
                        '  strQuery = "Select sum(""U_Z_Comm"") ""U_Z_Comm"",sum(""U_Z_MarkReb"") ""U_Z_MarkReb"",""OcrCode"" from INV1 where ""DocEntry""='" & DocNum & "' and ""U_Z_IsComm""='Y' group by ""OcrCode"""
                        strQuery = "Select (""U_Z_Comm"") ""U_Z_Comm"",(""U_Z_MarkReb"") ""U_Z_MarkReb"",""OcrCode"",""OcrCode2"",""OcrCode3"",""OcrCode4"",""OcrCode5"" from INV1 where ""DocEntry""='" & DocNum & "' and ""U_Z_IsComm""='Y' "
                        oTemp1.DoQuery(strQuery)
                        Dim dbLComm, dblMarketing As Double
                        Dim intLineCount As Integer = 0
                        For intloop As Integer = 0 To oTemp1.RecordCount - 1
                            dbLComm = 0
                            dblMarketing = 0
                            dbLComm = oTemp1.Fields.Item("U_Z_Comm").Value
                            dblMarketing = oTemp1.Fields.Item("U_Z_MarkReb").Value
                            If intLineCount > 0 Then
                                oAPInv.Lines.Add()
                            End If
                            oAPInv.Lines.SetCurrentLine(intLineCount)
                            oTest.DoQuery("Select * from ""@Z_OCRE""")
                            If dbLComm > 0 Then
                                oAPInv.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTest.Fields.Item("U_Z_ProReb").Value) 'oTest.Fields.Item("U_Z_MarkReb").Value
                                intLineCount = intLineCount + 1
                                oAPInv.Lines.LineTotal = dbLComm
                                oAPInv.Lines.ItemDescription = "Commission Rebate"
                                blnLineExists = True
                                If oTemp1.Fields.Item("OcrCode").Value <> "" Then
                                    oAPInv.Lines.CostingCode = oTemp1.Fields.Item("OcrCode").Value
                                End If
                                If oTemp1.Fields.Item("OcrCode2").Value <> "" Then
                                    oAPInv.Lines.CostingCode2 = oTemp1.Fields.Item("OcrCode2").Value
                                End If
                                If oTemp1.Fields.Item("OcrCode3").Value <> "" Then
                                    oAPInv.Lines.CostingCode3 = oTemp1.Fields.Item("OcrCode3").Value
                                End If
                                If oTemp1.Fields.Item("OcrCode4").Value <> "" Then
                                    oAPInv.Lines.CostingCode4 = oTemp1.Fields.Item("OcrCode4").Value
                                End If
                                If oTemp1.Fields.Item("OcrCode5").Value <> "" Then
                                    oAPInv.Lines.CostingCode5 = oTemp1.Fields.Item("OcrCode5").Value
                                End If
                            End If
                            If intLineCount > 0 Then
                                oAPInv.Lines.Add()
                            End If
                            oAPInv.Lines.SetCurrentLine(intLineCount)
                            If dblMarketing > 0 Then
                                oAPInv.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTest.Fields.Item("U_Z_MarkReb").Value) 'oTest.Fields.Item("U_Z_MarkReb").Value
                                intLineCount = intLineCount + 1
                                oAPInv.Lines.LineTotal = dblMarketing
                                oAPInv.Lines.ItemDescription = "Marketing  Rebate"
                                blnLineExists = True
                                If oTemp1.Fields.Item("OcrCode").Value <> "" Then
                                    oAPInv.Lines.CostingCode = oTemp1.Fields.Item("OcrCode").Value
                                End If
                                If oTemp1.Fields.Item("OcrCode2").Value <> "" Then
                                    oAPInv.Lines.CostingCode2 = oTemp1.Fields.Item("OcrCode2").Value
                                End If
                                If oTemp1.Fields.Item("OcrCode3").Value <> "" Then
                                    oAPInv.Lines.CostingCode3 = oTemp1.Fields.Item("OcrCode3").Value
                                End If
                                If oTemp1.Fields.Item("OcrCode4").Value <> "" Then
                                    oAPInv.Lines.CostingCode4 = oTemp1.Fields.Item("OcrCode4").Value
                                End If
                                If oTemp1.Fields.Item("OcrCode5").Value <> "" Then
                                    oAPInv.Lines.CostingCode5 = oTemp1.Fields.Item("OcrCode5").Value
                                End If
                            End If
                            oTemp1.MoveNext()
                        Next
                        If blnLineExists = True Then
                            If oAPInv.Add <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            Else
                                oTest2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim strDocNum As String
                                oApplication.Company.GetNewObjectCode(strDocNum)
                                oAPInv.GetByKey(CInt(strDocNum))
                                strDocNum = oAPInv.DocNum
                                strQuery = "Update OINV set ""U_Z_BaseEntry""='" & oAPInv.DocEntry.ToString & "', ""U_Z_APInvoice""='" & strDocNum & "' where ""DocEntry""=" & DocNum
                                oTest2.DoQuery(strQuery)
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Public Function CreditAPCreditNote(ByVal DocNum As String) As Boolean
        Dim oAPInv, oAPInv2 As SAPbobsCOM.Documents
        Dim oAPInv1 As SAPbobsCOM.Documents
        Dim oTest, otest1, oRec, oTemp1, oTest2 As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oAPInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        oAPInv2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
        oAPInv1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
        Try
            '  ApplyRebateAmount(DocNum)
            strQuery = "Select * from OINV where ""DocEntry""=" & DocNum & ""
            oTest.DoQuery(strQuery)
            If oTest.RecordCount > 0 Then
                If oTest.Fields.Item("U_Z_APInvoice").Value <> "" Then
                    oRec.DoQuery("Select * from OPCH where ""DocNum""=" & oTest.Fields.Item("U_Z_APInvoice").Value)
                    If oRec.RecordCount <= 0 Then
                        Return True
                    Else
                        '  MsgBox(oRec.Fields.Item("DocEntry").Value)
                    End If

                    If oAPInv.GetByKey(oRec.Fields.Item("DocEntry").Value) Then
                        If oAPInv.DocumentStatus = SAPbobsCOM.BoStatus.bost_Open Then
                            oAPInv2 = oAPInv.CreateCancellationDocument()
                            If oAPInv2.Add() <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                        Else
                            oAPInv1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
                            oAPInv1.DocDate = oTest.Fields.Item("DocDate").Value
                            oAPInv1.DocDueDate = oTest.Fields.Item("DocDueDate").Value
                            oAPInv1.CardCode = oAPInv.CardCode
                            oAPInv1.NumAtCard = oAPInv.NumAtCard
                            oAPInv1.UserFields.Fields.Item("U_Z_ARInvoice").Value = oTest.Fields.Item("DocNum").Value.ToString
                            oAPInv1.UserFields.Fields.Item("U_Z_BaseEntry").Value = DocNum
                            oAPInv1.Comments = "Rebate/Commission posting canceled based on A/P Invoice  : " & oTest.Fields.Item("DocNum").Value.ToString
                            oAPInv1.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                            For intLoop As Integer = 0 To oAPInv.Lines.Count - 1
                                If intLoop > 0 Then
                                    oAPInv1.Lines.Add()
                                    oAPInv1.Lines.SetCurrentLine(intLoop)
                                End If
                                oAPInv.Lines.SetCurrentLine(intLoop)
                                oAPInv1.Lines.AccountCode = oAPInv.Lines.AccountCode
                                oAPInv1.Lines.ItemDescription = oAPInv.Lines.ItemDescription
                                oAPInv1.Lines.LineTotal = oAPInv.Lines.LineTotal
                            Next
                            If oAPInv1.Add <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            Else
                                oTest2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim strDocNum As String
                                oApplication.Company.GetNewObjectCode(strDocNum)
                                oAPInv1.GetByKey(CInt(strDocNum))
                                strDocNum = oAPInv1.DocNum
                                strQuery = "Update OINV set ""U_Z_BaseEntry""='" & oAPInv1.DocEntry.ToString & "', ""U_Z_APInvoice""='" & strDocNum & "' where ""DocEntry""=" & DocNum
                                oTest2.DoQuery(strQuery)
                            End If
                        End If
                    End If

                End If
                Return True
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Public Function CancelJournal_CreditNote(ByVal DocNum As String) As Boolean
        Dim oAPInv, oAPInv2 As SAPbobsCOM.JournalEntries
        Dim oAPInv1 As SAPbobsCOM.Documents
        Dim oTest, otest1, oRec, oTemp1, oTest2 As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oAPInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
        oAPInv2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
        oAPInv1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
        Try
            '  ApplyRebateAmount(DocNum)
            strQuery = "Select * from ORIN where ""DocEntry""=" & DocNum & ""
            oTest.DoQuery(strQuery)
            If oTest.RecordCount > 0 Then
                If oTest.Fields.Item("U_Z_JournalRef").Value <> "" Then
                    oRec.DoQuery("Select * from OJDT where ""TransId""=" & oTest.Fields.Item("U_Z_JournalRef").Value)
                    If oRec.RecordCount <= 0 Then
                        Return True
                    Else
                        '  MsgBox(oRec.Fields.Item("DocEntry").Value)
                    End If

                    If oAPInv.GetByKey(oRec.Fields.Item("TransId").Value) Then
                        If 1 = 1 Then 'oAPInv.DocumentStatu = SAPbobsCOM.BoStatus.bost_Open Then
                            If oAPInv.Cancel <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                        End If

                    End If

                End If

            End If
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Public Function CancelJournal(ByVal DocNum As String) As Boolean
        Dim oAPInv, oAPInv2 As SAPbobsCOM.JournalEntries
        Dim oAPInv1 As SAPbobsCOM.Documents
        Dim oTest, otest1, oRec, oTemp1, oTest2 As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oAPInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
        oAPInv2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
        oAPInv1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
        Try
            '  ApplyRebateAmount(DocNum)
            strQuery = "Select * from OINV where ""DocEntry""=" & DocNum & ""
            oTest.DoQuery(strQuery)
            If oTest.RecordCount > 0 Then
                If oTest.Fields.Item("U_Z_JournalRef").Value <> "" Then
                    oRec.DoQuery("Select * from OJDT where ""TransId""=" & oTest.Fields.Item("U_Z_JournalRef").Value)
                    If oRec.RecordCount <= 0 Then
                        Return True
                    Else
                        '  MsgBox(oRec.Fields.Item("DocEntry").Value)
                    End If

                    If oAPInv.GetByKey(oRec.Fields.Item("TransId").Value) Then
                        If 1 = 1 Then 'oAPInv.DocumentStatu = SAPbobsCOM.BoStatus.bost_Open Then
                            If oAPInv.Cancel <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                        End If

                    End If

                End If

            End If
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Public Function CancelARCreditNoe(ByVal DocNum As String) As Boolean
        Dim oAPInv, oAPInv2 As SAPbobsCOM.Documents
        Dim oAPInv1 As SAPbobsCOM.Documents
        Dim oTest, otest1, oRec, oTemp1, oTest2 As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oAPInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
        oAPInv2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
        oAPInv1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
        Try
            '  ApplyRebateAmount(DocNum)
            strQuery = "Select * from ORIN where ""DocEntry""=" & DocNum & ""
            oTest.DoQuery(strQuery)
            If oTest.RecordCount > 0 Then
                If oTest.Fields.Item("U_Z_APInvoice").Value <> "" Then
                    oRec.DoQuery("Select * from ORPC where ""DocNum""=" & oTest.Fields.Item("U_Z_APInvoice").Value)
                    If oRec.RecordCount <= 0 Then
                        Return True
                    Else
                        '  MsgBox(oRec.Fields.Item("DocEntry").Value)
                    End If

                    If oAPInv.GetByKey(oRec.Fields.Item("DocEntry").Value) Then
                        If oAPInv.DocumentStatus = SAPbobsCOM.BoStatus.bost_Open Then
                            oAPInv2 = oAPInv.CreateCancellationDocument()
                            If oAPInv2.Add() <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            End If
                        Else
                            oAPInv1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
                            oAPInv1.DocDate = oTest.Fields.Item("DocDate").Value
                            oAPInv1.DocDueDate = oTest.Fields.Item("DocDueDate").Value
                            oAPInv1.CardCode = oAPInv.CardCode
                            oAPInv1.UserFields.Fields.Item("U_Z_ARInvoice").Value = oTest.Fields.Item("DocNum").Value.ToString
                            oAPInv1.UserFields.Fields.Item("U_Z_BaseEntry").Value = DocNum
                            oAPInv1.NumAtCard = oAPInv.NumAtCard
                            oAPInv1.Comments = "Rebate/Commission posting -canceled Based on A/R Credit Note  : " & oTest.Fields.Item("DocNum").Value.ToString
                            oAPInv1.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                            For intLoop As Integer = 0 To oAPInv.Lines.Count - 1
                                If intLoop > 0 Then
                                    oAPInv1.Lines.Add()
                                    oAPInv1.Lines.SetCurrentLine(intLoop)
                                End If
                                oAPInv.Lines.SetCurrentLine(intLoop)
                                oAPInv1.Lines.AccountCode = oAPInv.Lines.AccountCode
                                oAPInv1.Lines.ItemDescription = oAPInv.Lines.ItemDescription
                                oAPInv1.Lines.LineTotal = oAPInv.Lines.LineTotal
                            Next
                            If oAPInv1.Add <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            Else
                                oTest2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim strDocNum As String
                                oApplication.Company.GetNewObjectCode(strDocNum)
                                oAPInv1.GetByKey(CInt(strDocNum))
                                strDocNum = oAPInv1.DocNum
                                strQuery = "Update OINV set ""U_Z_BaseEntry""='" & oAPInv1.DocEntry.ToString & "', ""U_Z_APInvoice""='" & strDocNum & "' where ""Docentry""=" & DocNum
                                oTest2.DoQuery(strQuery)
                            End If
                        End If
                    End If

                End If
                Return True
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Public Function CreditAPCreditNote_ARCreditNote(ByVal DocNum As String) As Boolean
        Dim oAPInv, oAPInv2 As SAPbobsCOM.Documents
        Dim oAPInv1 As SAPbobsCOM.Documents
        Dim oTest, otest1, oRec, oTemp1, oTest2 As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oAPInv = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
        oAPInv2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
        oAPInv1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
        Try
            ApplyRebateAmount_CreditNote(DocNum)
            strQuery = "Select * from ORIN where ""DocEntry""=" & DocNum & ""
            oTest.DoQuery(strQuery)
            If oTest.RecordCount > 0 Then
                If 1 = 1 Then ' oAPInv.GetByKey(oTest.Fields.Item("DocEntry").Value) Then
                    strQuery = "Select * from OCRD where ""CardCode""='" & oTest.Fields.Item("CardCode").Value & "' and ""U_Z_ComRePay""<>''"
                    otest1.DoQuery(strQuery)
                    If otest1.RecordCount <= 0 Then 'oAPInv.DocumentStatus = SAPbobsCOM.BoStatus.bost_Open Then
                        'oAPInv2 = oAPInv.CreateCancellationDocument()
                        'If oAPInv2.Add() <> 0 Then
                        '    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        '    Return False
                        'End If
                    Else
                        oAPInv1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
                        oAPInv1.DocDate = oTest.Fields.Item("DocDate").Value
                        oAPInv1.DocDueDate = oTest.Fields.Item("DocDueDate").Value
                        oAPInv1.CardCode = otest1.Fields.Item("U_Z_ComRePay").Value
                        oAPInv1.UserFields.Fields.Item("U_Z_ARInvoice").Value = oTest.Fields.Item("DocNum").Value.ToString
                        oAPInv1.UserFields.Fields.Item("U_Z_BaseEntry").Value = DocNum
                        oAPInv1.NumAtCard = "AR Credit Memo No : " & oTest.Fields.Item("DocNum").Value.ToString
                        oAPInv1.Comments = "Rebate/Commission posting based on A/R Credit Memo  : " & oTest.Fields.Item("DocNum").Value.ToString
                        oAPInv1.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                        strQuery = "Select sum(""U_Z_Comm"") ""U_Z_Comm"",sum(""U_Z_MarkReb"") ""U_Z_MarkReb"",""OcrCode"" from RIN1 where ""DocEntry""='" & DocNum & "' and ""U_Z_IsComm""='Y' group by ""OcrCode"""
                        oTemp1.DoQuery(strQuery)
                        Dim dbLComm, dblMarketing As Double
                        Dim blnLineExists As Boolean = False
                        Dim intLineCount As Integer = 0
                        For intloop As Integer = 0 To oTemp1.RecordCount - 1
                            dbLComm = 0
                            dblMarketing = 0
                            dbLComm = oTemp1.Fields.Item("U_Z_Comm").Value
                            dblMarketing = oTemp1.Fields.Item("U_Z_MarkReb").Value

                            oTest.DoQuery("Select * from ""@Z_OCRE""")
                            If dbLComm > 0 Then
                                If intLineCount > 0 Then
                                    oAPInv1.Lines.Add()
                                End If
                                oAPInv1.Lines.SetCurrentLine(intLineCount)
                                oAPInv1.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTest.Fields.Item("U_Z_ProReb").Value) 'oTest.Fields.Item("U_Z_MarkReb").Value
                                oAPInv1.Lines.LineTotal = dbLComm
                                oAPInv1.Lines.ItemDescription = "Commission Rebate"
                                blnLineExists = True
                                If oTemp1.Fields.Item("OcrCode").Value <> "" Then
                                    oAPInv1.Lines.CostingCode = oTemp1.Fields.Item("OcrCode").Value
                                End If
                                intLineCount = intLineCount + 1
                            End If
                            If dblMarketing > 0 Then
                                If intLineCount > 0 Then
                                    oAPInv1.Lines.Add()
                                End If
                                '  oAPInv.Lines.SetCurrentLine(intLineCount)
                                oAPInv1.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTest.Fields.Item("U_Z_MarkReb").Value) 'oTest.Fields.Item("U_Z_MarkReb").Value

                                oAPInv1.Lines.LineTotal = dblMarketing
                                oAPInv1.Lines.ItemDescription = "Marketing  Rebate"
                                blnLineExists = True
                                If oTemp1.Fields.Item("OcrCode").Value <> "" Then
                                    oAPInv1.Lines.CostingCode = oTemp1.Fields.Item("OcrCode").Value
                                End If
                                intLineCount = intLineCount + 1
                            End If
                            oTemp1.MoveNext()
                        Next
                        If blnLineExists = True Then
                            If oAPInv1.Add <> 0 Then
                                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                Return False
                            Else
                                oTest2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                Dim strDocNum As String
                                oApplication.Company.GetNewObjectCode(strDocNum)
                                oAPInv1.GetByKey(CInt(strDocNum))
                                strDocNum = oAPInv1.DocNum
                                strQuery = "Update ORIN set ""U_Z_BaseEntry""='" & oAPInv1.DocEntry.ToString & "', ""U_Z_APInvoice""='" & strDocNum & "' where ""DocEntry""=" & DocNum
                                oTest2.DoQuery(strQuery)
                            End If
                        End If
                    End If

                End If
                Return True


            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

End Class
