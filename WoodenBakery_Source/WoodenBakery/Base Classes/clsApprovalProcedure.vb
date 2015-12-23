Imports System.IO
Imports System.Net.Mail
Imports System.Collections.Specialized
Imports System.Security.Cryptography
Imports System.Text

Public Class clsApprovalProcedure
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
    Dim oRec, oRecordSet As SAPbobsCOM.Recordset
    Dim SmtpServer As New Net.Mail.SmtpClient()
    Dim mail As New Net.Mail.MailMessage
    Dim mailServer As String
    Dim mailPort As String
    Dim mailId As String
    Dim mailUser As String
    Dim mailPwd As String
    Dim mailSSL As String
    Dim toID As String
    Dim ccID As String
    Dim mType As String
    Private FormNum As Integer
    Dim oCombo, oCombobox1, oCombobox2 As SAPbouiCOM.ComboBox
    Dim oEdit As SAPbouiCOM.EditText
    Dim oExEdit As SAPbouiCOM.EditText
    Dim StrMailMessage, strSubject As String

    Public Sub New()
        MyBase.New()
        FormNum = 1
    End Sub

    Public Function GetTemplateID(ByVal DocType As modVariables.HeaderDoctype, ByVal OrginatorId As String, ByVal ChildTbl As String, ByVal ChildColumn As String) As String
        Try
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim Status As String = ""
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If blnIsHanaDB Then
                strQuery = "Select * from ""@Z_OAPPT"" T0 left join """ & ChildTbl & """ T1 on T0.""DocEntry""=T1.""DocEntry"" "
                strQuery += " where IFNULL(T0.""U_Z_Active"",'N')='Y' and T0.""U_Z_DocType""='" & DocType.ToString() & "' and T1.""" & ChildColumn & """='" & OrginatorId.Trim() & "' "
            Else
                strQuery = "Select * from ""@Z_OAPPT"" T0 left join """ & ChildTbl & """ T1 on T0.""DocEntry""=T1.""DocEntry"" "
                strQuery += " where ISNULL(T0.""U_Z_Active"",'N')='Y' and T0.""U_Z_DocType""='" & DocType.ToString() & "' and T1.""" & ChildColumn & """='" & OrginatorId.Trim() & "' "
            End If
            oRecordSet.DoQuery(strQuery)
            If oRecordSet.RecordCount > 0 Then
                Status = oRecordSet.Fields.Item("DocEntry").Value
            Else
                Status = "0"
            End If
            Return Status
        Catch ex As Exception
            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

    Public Function DocApproval(ByVal DocType As modVariables.HeaderDoctype, ByVal OrginatorId As String, ByVal ChildTbl As String, ByVal ChildColumn As String) As String
        Try
            Dim Status As String = ""
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select * from ""@Z_HR_OAPPT"" T0 left join """ & ChildTbl & """ T1 on T0.""DocEntry""=T1.""DocEntry"" "
            strQuery += " where T0.""U_Z_Active""='Y' and T0.""U_Z_DocType""='" & DocType.ToString() & "' and  T1.""" & ChildColumn & """='" & OrginatorId.Trim() & "' "
            oRecordSet.DoQuery(strQuery)
            If oRecordSet.RecordCount > 0 Then
                Status = "P"
            Else
                Status = "A"
            End If
            Return Status
        Catch ex As Exception
            MsgBox(oApplication.Company.GetLastErrorDescription)
            Return False
        End Try
    End Function

    Public Sub UpdateApprovalRequired(ByVal strTable As String, ByVal sColumn As String, ByVal StrCode As String, ByVal ReqValue As String, ByVal AppTempId As String, ByVal status As String)
        Try
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If blnIsHanaDB Then
                strQuery = "Update """ & strTable & """ set ""U_Z_IsApp""='" & ReqValue & "',""U_Z_AppReqDate""= NOW(),""U_Z_ApproveId""='" & AppTempId & "',""U_Z_AppStatus""='" & status & "'"
                strQuery += " where """ & sColumn & """='" & StrCode & "'"
            Else
                strQuery = "Update """ & strTable & """ set ""U_Z_IsApp""='" & ReqValue & "',""U_Z_AppReqDate""= GetDate(),""U_Z_ApproveId""='" & AppTempId & "',""U_Z_AppStatus""='" & status & "'"
                strQuery += " where """ & sColumn & """='" & StrCode & "'"
            End If
            
            oRecordSet.DoQuery(strQuery)
        Catch ex As Exception
            MsgBox(oApplication.Company.GetLastErrorDescription)
        End Try
    End Sub

    Public Sub InitialCurNextApprover(ByVal strTable As String, ByVal sColumn As String, ByVal StrCode As String, ByVal strTemplateNo As String, ByVal strMessage As String, ByVal strSubject As String)
        Try
            Dim strMessageUser As String
            Dim oRecordSet, oTemp As SAPbobsCOM.Recordset
            Dim oCmpSrv As SAPbobsCOM.CompanyService
            Dim oMessageService As SAPbobsCOM.MessagesService
            Dim oMessage As SAPbobsCOM.Message
            oCmpSrv = oApplication.Company.GetCompanyService()
            oMessageService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService)
            oMessage = oMessageService.GetDataInterface(SAPbobsCOM.MessagesServiceDataInterfaces.msdiMessage)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If blnIsHanaDB Then
                strQuery = "Select Top 1 ""U_Z_AUser"" From ""@Z_APPT2"" Where ""DocEntry"" = '" + strTemplateNo + "'  and IFNULL(""U_Z_AMan"",'')='Y' Order By ""LineId"" Asc "
            Else
                strQuery = "Select Top 1 ""U_Z_AUser"" From ""@Z_APPT2"" Where ""DocEntry"" = '" + strTemplateNo + "'  and ISNULL(""U_Z_AMan"",'')='Y' Order By ""LineId"" Asc "
            End If
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                strMessageUser = oRecordSet.Fields.Item(0).Value
                strQuery = "Update """ & strTable & """ set ""U_Z_CurrApprover""='" & strMessageUser & "',""U_Z_NextApprover""='" & strMessageUser & "' where """ & sColumn & """='" & StrCode & "'"
                oTemp.DoQuery(strQuery)
                SAPUserMessage(strMessage, StrCode, strMessageUser, strSubject)
                SendMail_Approval(strMessage, strSubject, getEmailid(strMessageUser))
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Function updateFinalStatus(ByVal aForm As SAPbouiCOM.Form, ByVal UserName As String, ByVal strStatus As String, ByVal strTemplateNo As String, ByVal enDocType As modVariables.HeaderDoctype, ByVal crUser As String, ByVal strDocEntry As String) As Boolean
        Try
            Dim oRecordSet, oTemp As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If strStatus = "A" Then
                strQuery = "Select ""DocEntry"" From ""@Z_APPT2"" Where ""DocEntry"" = '" + strTemplateNo + "'  and  ""U_Z_AFinal"" = 'Y' and ""U_Z_AUser"" = '" & UserName & "'"
                oRecordSet.DoQuery(strQuery)
                If Not oRecordSet.EoF Then
                    Select Case enDocType
                        Case HeaderDoctype.Fix 'Fixed Asset Transaction
                            strQuery = "Update ""@Z_OFATA"" Set ""U_Z_AppStatus"" = 'A',""U_Z_DocStatus"" = 'A',""U_Z_AppRemarks""='" & oApplication.Utilities.getEdittextvalue(aForm, "10") & "' Where ""DocEntry"" = '" + strDocEntry + "'"
                            oRecordSet.DoQuery(strQuery)
                            oApplication.Utilities.UpdateFixedAsset(strDocEntry) 'Update Fixed Asset
                            StrMailMessage = "Fised Asset transaction has been approved for the Transaction number :" & CInt(strDocEntry)
                            strSubject = "Fixed asset transaction approval notification."
                            SAPUserMessage(StrMailMessage, strDocEntry, crUser, strSubject)
                        Case HeaderDoctype.Spl 'Fixed Asset Transaction
                            strQuery = "Update ""@Z_OVPL"" Set ""U_Z_AppStatus"" = 'A',""U_Z_DocStatus"" = 'A',""U_Z_AppRemarks""='" & oApplication.Utilities.getEdittextvalue(aForm, "10") & "' Where ""DocEntry"" = '" + strDocEntry + "'"
                            oRecordSet.DoQuery(strQuery)
                            oApplication.Utilities.UpdateSupplierCatelog(strDocEntry)
                            StrMailMessage = "Supplier Price transaction has been approved for the Transaction number :" & CInt(strDocEntry)
                            strSubject = "Supplier Price transaction approval notification."
                            SAPUserMessage(StrMailMessage, strDocEntry, crUser, strSubject)
                    End Select
                End If
            ElseIf strStatus = "R" Then
                Select Case enDocType
                    Case HeaderDoctype.Fix
                        strQuery = "Update ""@Z_OFATA"" Set ""U_Z_AppStatus"" = 'R',""U_Z_DocStatus"" = 'R',""U_Z_AppRemarks""='" & oApplication.Utilities.getEdittextvalue(aForm, "10") & "' Where ""DocEntry"" = '" + strDocEntry + "'"
                        oRecordSet.DoQuery(strQuery)
                        StrMailMessage = "Fised Asset transaction has been rejected for the Transaction number :" & CInt(strDocEntry) & " and Reason :" & oApplication.Utilities.getEdittextvalue(aForm, "10") & ""
                        strSubject = "Fixed asset transaction approval notification."
                        SAPUserMessage(StrMailMessage, strDocEntry, crUser, strSubject)
                        SendMail_Approval(StrMailMessage, strSubject, getEmailid(crUser))
                    Case HeaderDoctype.Spl
                        strQuery = "Update ""@Z_OVPL"" Set ""U_Z_AppStatus"" = 'R',""U_Z_DocStatus"" = 'R',""U_Z_AppRemarks""='" & oApplication.Utilities.getEdittextvalue(aForm, "10") & "' Where ""DocEntry"" = '" + strDocEntry + "'"
                        oRecordSet.DoQuery(strQuery)
                        StrMailMessage = "Supplier Price  transaction has been rejected for the Transaction number :" & CInt(strDocEntry) & " and Reason :" & oApplication.Utilities.getEdittextvalue(aForm, "10") & ""
                        strSubject = "Supplier Price transaction approval notification."
                        SAPUserMessage(StrMailMessage, strDocEntry, crUser, strSubject)
                        SendMail_Approval(StrMailMessage, strSubject, getEmailid(crUser))
                End Select
            End If
            Select Case enDocType
                Case HeaderDoctype.Fix
                    If strStatus = "A" And oCombo.Selected.Value <> "-" Then
                        StrMailMessage = "Fixed asset transaction need approval for the transaction id is :" & strDocEntry
                        strSubject = "Fixed asset transaction approval notification"
                        SendMessage(StrMailMessage, strSubject, enDocType, oApplication.Company.UserName, strTemplateNo, "@Z_OFATA", "DocEntry", strDocEntry, getEmailid(oApplication.Company.UserName))
                    End If
                Case HeaderDoctype.Spl
                    If strStatus = "A" And oCombo.Selected.Value <> "-" Then
                        StrMailMessage = "Supplier Price transaction need approval for the transaction id is :" & strDocEntry
                        strSubject = "Supplier Price transaction approval notification"
                        SendMessage(StrMailMessage, strSubject, enDocType, oApplication.Company.UserName, strTemplateNo, "@Z_OVPL", "DocEntry", strDocEntry, getEmailid(oApplication.Company.UserName))
                        End
                    End If
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Private Function getEmailid(ByVal Username As String) As String
        Dim aMail As String = ""
        Try
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If blnIsHanaDB Then
                strQuery = "Select IFNULL(""E_Mail"",'') from OUSR where ""USER_CODE""='" & Username & "'"
            Else
                strQuery = "Select ISNULL(""E_Mail"",'') from OUSR where ""USER_CODE""='" & Username & "'"
            End If
            oRecordSet.DoQuery(strQuery)
            aMail = oRecordSet.Fields.Item(0).Value
            Return aMail
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Private Sub SAPUserMessage(ByVal strMessage As String, ByVal strDocEntry As String, ByVal SAPUser As String, ByVal strSubject As String)
        Dim oCmpSrv As SAPbobsCOM.CompanyService
        Dim oMessageService As SAPbobsCOM.MessagesService
        Dim oMessage As SAPbobsCOM.Message
        Dim pMessageDataColumns As SAPbobsCOM.MessageDataColumns
        Dim pMessageDataColumn As SAPbobsCOM.MessageDataColumn
        Dim oLines As SAPbobsCOM.MessageDataLines
        Dim oLine As SAPbobsCOM.MessageDataLine
        Dim oRecipientCollection As SAPbobsCOM.RecipientCollection
        oCmpSrv = oApplication.Company.GetCompanyService()
        oMessageService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService)
        oMessage = oMessageService.GetDataInterface(SAPbobsCOM.MessagesServiceDataInterfaces.msdiMessage)
        oMessage.Subject = strSubject
        oMessage.Text = strMessage
        oRecipientCollection = oMessage.RecipientCollection
        oRecipientCollection.Add()
        oRecipientCollection.Item(0).SendInternal = SAPbobsCOM.BoYesNoEnum.tYES
        oRecipientCollection.Item(0).UserCode = SAPUser
        pMessageDataColumns = oMessage.MessageDataColumns
        pMessageDataColumn = pMessageDataColumns.Add()
        pMessageDataColumn.ColumnName = "Document Number"
        oLines = pMessageDataColumn.MessageDataLines()
        oLine = oLines.Add()
        oLine.Value = strDocEntry
        oMessageService.SendMessage(oMessage)
    End Sub

    Public Sub SendMessage(ByVal aMessage As String, ByVal MailSubject As String, ByVal enDocType As modVariables.HeaderDoctype, ByVal strAuthorizer As String, ByVal strTemplateNo As String, _
                           ByVal strTable As String, ByVal sColumn As String, ByVal StrCode As String, ByVal EMailId As String)
        Try
            Dim strMessageUser As String
            Dim intLineID As Integer
            Dim oRecordSet, oTemp As SAPbobsCOM.Recordset
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select ""LineId"" From ""@Z_APPT2"" Where ""DocEntry"" = '" & strTemplateNo & "' And ""U_Z_AUser"" = '" & strAuthorizer & "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                intLineID = CInt(oRecordSet.Fields.Item(0).Value)
                If blnIsHanaDB Then
                    strQuery = "Select Top 1 ""U_Z_AUser"" From ""@Z_APPT2"" Where  ""DocEntry"" = '" & strTemplateNo & "' And ""LineId"" > '" & intLineID.ToString() & "' and IFNULL(""U_Z_AMan"",'')='Y'  Order By ""LineId"" Asc "
                Else
                    strQuery = "Select Top 1 ""U_Z_AUser"" From ""@Z_APPT2"" Where  ""DocEntry"" = '" & strTemplateNo & "' And ""LineId"" > '" & intLineID.ToString() & "' and ISNULL(""U_Z_AMan"",'')='Y'  Order By ""LineId"" Asc "
                End If
                oRecordSet.DoQuery(strQuery)
                If Not oRecordSet.EoF Then
                    strMessageUser = oRecordSet.Fields.Item(0).Value
                    Select Case enDocType
                        Case HeaderDoctype.Fix
                            strQuery = "Update """ & strTable & """ set ""U_Z_CurrApprover""='" & oApplication.Company.UserName & "',""U_Z_NextApprover""='" & strMessageUser & "' where """ & sColumn & """='" & StrCode & "'"
                            oTemp.DoQuery(strQuery)
                            SAPUserMessage(aMessage, StrCode, strMessageUser, MailSubject)
                            'SendMail_Approval(aMessage, MailSubject, getEmailid(strMessageUser))
                        Case HeaderDoctype.Spl
                            strQuery = "Update """ & strTable & """ set ""U_Z_CurrApprover""='" & oApplication.Company.UserName & "',""U_Z_NextApprover""='" & strMessageUser & "' where """ & sColumn & """='" & StrCode & "'"
                            oTemp.DoQuery(strQuery)
                            SAPUserMessage(aMessage, StrCode, strMessageUser, MailSubject)
                            'SendMail_Approval(aMessage, MailSubject, getEmailid(strMessageUser))
                    End Select
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub SendMail_Approval(ByVal aMessage As String, ByVal MailSubject As String, ByVal EMailId As String)
        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet.DoQuery("Select ""U_Z_SMTPSERV"",""U_Z_SMTPPORT"",""U_Z_SMTPUSER"",""U_Z_SMTPPWD"",""U_Z_SSL"" From ""@Z_OMAIL""")
        If Not oRecordSet.EoF Then
            mailServer = oRecordSet.Fields.Item("U_Z_SMTPSERV").Value
            mailPort = oRecordSet.Fields.Item("U_Z_SMTPPORT").Value
            mailId = oRecordSet.Fields.Item("U_Z_SMTPUSER").Value
            mailPwd = oRecordSet.Fields.Item("U_Z_SMTPPWD").Value
            mailSSL = oRecordSet.Fields.Item("U_Z_SSL").Value
            If mailServer <> "" And mailId <> "" And mailPwd <> "" Then
                If EMailId <> "" Then
                    SendMailforApproval(mailServer, mailPort, mailId, mailPwd, mailSSL, EMailId, EMailId, aMessage, MailSubject)
                End If
            Else
                oApplication.Utilities.Message("Mail Server Details Not Configured...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If
        End If
    End Sub

    Private Sub SendMailforApproval(ByVal mailServer As String, ByVal mailPort As String, ByVal mailId As String, ByVal mailpwd As String, _
                                    ByVal mailSSL As String, ByVal toId As String, ByVal ccId As String, ByVal Message As String, ByVal Subject As String)
        Try
            SmtpServer.Credentials = New Net.NetworkCredential(mailId, mailpwd)
            SmtpServer.Port = mailPort
            SmtpServer.EnableSsl = mailSSL
            SmtpServer.Host = mailServer
            mail = New Net.Mail.MailMessage()
            mail.From = New Net.Mail.MailAddress(mailId)
            mail.To.Add(toId)
            mail.IsBodyHtml = True
            mail.Priority = MailPriority.High
            mail.Subject = Subject
            mail.Body = Message
            SmtpServer.Send(mail)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
            mail.Dispose()
        End Try
    End Sub

    Public Sub LoadViewHistory(ByVal aForm As SAPbouiCOM.Form, ByVal enDocType As modVariables.HeaderDoctype, ByVal strDocEntry As String, ByVal GridId As String)
        Try
            oGrid = aForm.Items.Item(GridId).Specific
            strQuery = " Select ""DocEntry"",""U_Z_DocEntry"",""U_Z_DocType"",""U_Z_EmpId"",""U_Z_EmpName"",""U_Z_ApproveBy"",""CreateDate"",""CreateTime"",""UpdateDate"","
            strQuery += " ""UpdateTime"",""U_Z_AppStatus"",""U_Z_Remarks"" From ""@Z_APHIS"" "
            strQuery += " Where ""U_Z_DocType"" = '" & enDocType.ToString() & "'"
            strQuery += " And ""U_Z_DocEntry"" = '" & strDocEntry & "'"
            oGrid.DataTable.ExecuteQuery(strQuery)
            formatHistory(aForm, GridId)
            oApplication.Utilities.assignMatrixLineno(oGrid, aForm)
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Private Sub formatHistory(ByVal aForm As SAPbouiCOM.Form, ByVal GridId As String)
        Try
            aForm.Freeze(True)
            Dim oGrid As SAPbouiCOM.Grid
            Dim oGridCombo As SAPbouiCOM.ComboBoxColumn
            Dim oEditTextColumn As SAPbouiCOM.EditTextColumn
            oGrid = aForm.Items.Item(GridId).Specific
            oGrid.Columns.Item("DocEntry").Visible = False
            oGrid.Columns.Item("U_Z_DocEntry").TitleObject.Caption = "Reference No."
            oGrid.Columns.Item("U_Z_DocEntry").Visible = False
            oGrid.Columns.Item("U_Z_DocType").Visible = False
            oGrid.Columns.Item("U_Z_EmpId").TitleObject.Caption = "Employee ID"
            oEditTextColumn = oGrid.Columns.Item("U_Z_EmpId")
            oEditTextColumn.LinkedObjectType = "171"
            oGrid.Columns.Item("U_Z_EmpName").TitleObject.Caption = "Employee Name"
            oGrid.Columns.Item("U_Z_ApproveBy").TitleObject.Caption = "Approved By"
            oGrid.Columns.Item("U_Z_AppStatus").TitleObject.Caption = "Approved Status"
            oGrid.Columns.Item("U_Z_AppStatus").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oGridCombo = oGrid.Columns.Item("U_Z_AppStatus")
            oGridCombo.ValidValues.Add("P", "Pending")
            oGridCombo.ValidValues.Add("A", "Approved")
            oGridCombo.ValidValues.Add("R", "Rejected")
            oGridCombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
            oGrid.Columns.Item("U_Z_Remarks").TitleObject.Caption = "Remarks"
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oGrid.AutoResizeColumns()
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Public Sub addUpdateDocument(ByVal aForm As SAPbouiCOM.Form, ByVal enDocType As modVariables.HeaderDoctype)
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        oCompanyService = oApplication.Company.GetCompanyService()
        Dim otestRs As SAPbobsCOM.Recordset
        Dim blnRecordExists As Boolean = False
        Dim oRecordSet As SAPbobsCOM.Recordset
        Try
            If oApplication.SBO_Application.MessageBox("Documents once approved can not be changed. Do you want Continue?", , "Contine", "Cancel") = 2 Then
                Exit Sub
            End If
            oGeneralService = oCompanyService.GetGeneralService("Z_APHIS")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otestRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oGrid = aForm.Items.Item("1").Specific
            Dim strDocEntry As String = ""
            Dim strHeader As String = enDocType
            Dim strEmpID As String
            Dim HeadDocEntry, UserLineId As Integer
            oCombo = aForm.Items.Item("8").Specific
            oExEdit = aForm.Items.Item("10").Specific

            Select Case enDocType
                Case HeaderDoctype.Fix
                    For index As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                        If oGrid.Rows.IsSelected(index) Then
                            strDocEntry = oGrid.DataTable.GetValue("DocEntry", index)
                            strEmpID = oGrid.DataTable.GetValue("Creator", index)
                            Exit For
                        End If
                    Next
                Case HeaderDoctype.Spl
                    For index As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                        If oGrid.Rows.IsSelected(index) Then
                            strDocEntry = oGrid.DataTable.GetValue("DocEntry", index)
                            strEmpID = oGrid.DataTable.GetValue("U_Z_CardCode", index)
                            Exit For
                        End If
                    Next
            End Select
            Select Case enDocType
                Case HeaderDoctype.Fix
                    strQuery = "select T0.""DocEntry"",T1.""LineId"" from ""@Z_OAPPT"" T0 JOIN ""@Z_APPT2"" T1 on T0.""DocEntry""=T1.""DocEntry"""
                    strQuery += " JOIN ""@Z_APPT1"" T2 on T1.""DocEntry""=T2.""DocEntry"""
                    strQuery += " where T0.""U_Z_DocType""='" & enDocType.ToString() & "' AND T2.""U_Z_EmpId""='" & strEmpID & "' AND T1.""U_Z_AUser""='" & oApplication.Company.UserName & "'"
                Case HeaderDoctype.Spl
                    strQuery = "select T0.""DocEntry"",T1.""LineId"" from ""@Z_OAPPT"" T0 JOIN ""@Z_APPT2"" T1 on T0.""DocEntry""=T1.""DocEntry"""
                    strQuery += " JOIN ""@Z_APPT1"" T2 on T1.""DocEntry""=T2.""DocEntry"""
                    strQuery += " where T0.""U_Z_DocType""='" & enDocType.ToString() & "' AND T2.""U_Z_EmpId""='" & strEmpID & "' AND T1.""U_Z_AUser""='" & oApplication.Company.UserName & "'"
            End Select
            otestRs.DoQuery(strQuery)
            If otestRs.RecordCount > 0 Then
                HeadDocEntry = otestRs.Fields.Item(0).Value
                UserLineId = otestRs.Fields.Item(1).Value
            End If
            Dim strEmpName As String = ""
            strQuery = "Select * from ""@Z_APHIS"" where ""U_Z_DocEntry""='" & strDocEntry & "' and ""U_Z_DocType""='" & enDocType.ToString() & "' and ""U_Z_ApproveBy""='" & oApplication.Company.UserName & "'"
            oRecordSet.DoQuery(strQuery)
            If oRecordSet.RecordCount > 0 Then
                oGeneralParams.SetProperty("DocEntry", oRecordSet.Fields.Item("DocEntry").Value)
                oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                oGeneralData.SetProperty("U_Z_AppStatus", oCombo.Selected.Value)
                oGeneralData.SetProperty("U_Z_Remarks", oExEdit.Value)
                oGeneralData.SetProperty("U_Z_ADocEntry", HeadDocEntry)
                oGeneralData.SetProperty("U_Z_ALineId", UserLineId)
                Dim oTemp As SAPbobsCOM.Recordset
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If blnIsHanaDB Then
                    strQuery = "Select * ,IFNULL(""firstName"",'') ||  ' ' || IFNULL(""middleName"",'') ||  ' ' || IFNULL(""lastName"",'') AS ""EmpName"" from OHEM where ""userId""=" & oApplication.Company.UserSignature
                Else
                    strQuery = "Select * ,ISNULL(""firstName"",'') +  ' ' + ISNULL(""middleName"",'') +  ' ' + ISNULL(""lastName"",'') AS ""EmpName"" from OHEM where ""userId""=" & oApplication.Company.UserSignature
                End If
                oTemp.DoQuery(strQuery)
                If oTemp.RecordCount > 0 Then
                    oGeneralData.SetProperty("U_Z_EmpId", oTemp.Fields.Item("empID").Value.ToString())
                    oGeneralData.SetProperty("U_Z_EmpName", oTemp.Fields.Item("EmpName").Value)
                    strEmpName = oTemp.Fields.Item("EmpName").Value
                Else
                    oGeneralData.SetProperty("U_Z_EmpId", "")
                    oGeneralData.SetProperty("U_Z_EmpName", "")
                End If
                oGeneralService.Update(oGeneralData)
            ElseIf (strDocEntry <> "" And strDocEntry <> "0") Then
                Dim oTemp As SAPbobsCOM.Recordset
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim strQuery As String = String.Empty
                If blnIsHanaDB Then
                    strQuery = "Select * ,IFNULL(""firstName"",'') || ' ' || IFNULL(""middleName"",'') ||  ' ' || IFNULL(""lastName"",'') AS ""EmpName"" from OHEM where ""userId""=" & oApplication.Company.UserSignature
                Else
                    strQuery = "Select * ,ISNULL(""firstName"",'') + ' ' + ISNULL(""middleName"",'') +  ' ' + ISNULL(""lastName"",'') AS ""EmpName"" from OHEM where ""userId""=" & oApplication.Company.UserSignature
                End If
                oTemp.DoQuery(strQuery)
                If oTemp.RecordCount > 0 Then
                    oGeneralData.SetProperty("U_Z_EmpId", oTemp.Fields.Item("empID").Value.ToString())
                    oGeneralData.SetProperty("U_Z_EmpName", oTemp.Fields.Item("EmpName").Value)
                    strEmpName = oTemp.Fields.Item("EmpName").Value
                Else
                    oGeneralData.SetProperty("U_Z_EmpId", "")
                    oGeneralData.SetProperty("U_Z_EmpName", "")
                End If
                oGeneralData.SetProperty("U_Z_DocEntry", strDocEntry.ToString())
                oGeneralData.SetProperty("U_Z_DocType", enDocType.ToString())
                oGeneralData.SetProperty("U_Z_AppStatus", oCombo.Selected.Value)
                oGeneralData.SetProperty("U_Z_Remarks", oExEdit.Value)
                oGeneralData.SetProperty("U_Z_ApproveBy", oApplication.Company.UserName)
                oGeneralData.SetProperty("U_Z_Approvedt", System.DateTime.Now)
                oGeneralData.SetProperty("U_Z_ADocEntry", HeadDocEntry)
                oGeneralData.SetProperty("U_Z_ALineId", UserLineId)
                oGeneralService.Add(oGeneralData)
            End If

            updateFinalStatus(aForm, oApplication.Company.UserName, oCombo.Selected.Value, HeadDocEntry, enDocType, strEmpID, strDocEntry)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub Resize(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            aForm.Items.Item("1").Height = (aForm.Height / 2) - 50
            aForm.Items.Item("1").Width = aForm.Width - 10
            aForm.Items.Item("4").Top = aForm.Items.Item("1").Top + aForm.Items.Item("1").Height + 1
            aForm.Items.Item("5").Top = aForm.Items.Item("4").Top
            aForm.Items.Item("3").Top = aForm.Items.Item("4").Top + aForm.Items.Item("4").Height + 5
            aForm.Items.Item("3").Width = (aForm.Width / 2)
            aForm.Items.Item("3").Height = (aForm.Height / 2) - 50
            aForm.Items.Item("5").Left = aForm.Items.Item("3").Left + aForm.Items.Item("3").Width + 50
            aForm.Items.Item("7").Left = aForm.Items.Item("5").Left
            aForm.Items.Item("9").Left = aForm.Items.Item("5").Left
            aForm.Items.Item("8").Left = aForm.Items.Item("7").Left + aForm.Items.Item("7").Width + 1
            aForm.Items.Item("10").Left = aForm.Items.Item("9").Left + aForm.Items.Item("9").Width + 1
            aForm.Items.Item("8").Top = aForm.Items.Item("3").Top
            aForm.Items.Item("7").Top = aForm.Items.Item("8").Top
            aForm.Freeze(False)
        Catch ex As Exception

        End Try
    End Sub

    Public Function ApprovalValidation(ByVal aform As SAPbouiCOM.Form, ByVal enDocType As modVariables.HeaderDoctype) As Boolean
        Try
            oCombo = aform.Items.Item("8").Specific
            oExEdit = aform.Items.Item("10").Specific
            Select Case enDocType
                Case HeaderDoctype.Fix
                    If oCombo.Selected.Value = "R" Then
                        If oExEdit.Value = "" Then
                            oApplication.Utilities.Message("Remarks is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
            End Select
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function

End Class


