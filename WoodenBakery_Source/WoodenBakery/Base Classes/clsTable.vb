Public NotInheritable Class clsTable

#Region "Private Functions"
    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddTables
    'Parameter          : 
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Tables in DB. This function shall be called by 
    '                     public functions to create a table
    '**************************************************************************************************************
    Private Sub AddTables(ByVal strTab As String, ByVal strDesc As String, ByVal nType As SAPbobsCOM.BoUTBTableType)
        Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
        Try
            oUserTablesMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
            'Adding Table
            If Not oUserTablesMD.GetByKey(strTab) Then
                oUserTablesMD.TableName = strTab
                oUserTablesMD.TableDescription = strDesc
                oUserTablesMD.TableType = nType
                If oUserTablesMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
            oUserTablesMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddFields
    'Parameter          : SstrTab As String,strCol As String,
    '                     strDesc As String,nType As Integer,i,nEditSize,nSubType As Integer
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Fields in DB Tables. This function shall be called by 
    '                     public functions to create a Field
    '**************************************************************************************************************
    Private Sub AddFields(ByVal strTab As String, _
                            ByVal strCol As String, _
                                ByVal strDesc As String, _
                                    ByVal nType As SAPbobsCOM.BoFieldTypes, _
                                        Optional ByVal i As Integer = 0, _
                                            Optional ByVal nEditSize As Integer = 10, _
                                                Optional ByVal nSubType As SAPbobsCOM.BoFldSubTypes = 0, _
                                                    Optional ByVal Mandatory As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO)
        Dim oUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try

            If Not (strTab = "OINC" Or strTab = "OSCN" Or strTab = "OADM" Or strTab = "ORCT" Or strTab = "OVPM" Or strTab = "OPCH" Or strTab = "OITM" Or strTab = "OJDT" Or strTab = "INV1" Or strTab = "RDR1" Or strTab = "DRF1" Or strTab = "ODRF" Or strTab = "OINV" Or strTab = "OCRD" Or strTab = "SPP1" Or strTab = "OWHS" Or strTab = "OCRD") Then
                strTab = "@" + strTab
            End If

            If Not IsColumnExists(strTab, strCol) Then
                oUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

                oUserFieldMD.Description = strDesc
                oUserFieldMD.Name = strCol
                oUserFieldMD.Type = nType
                oUserFieldMD.SubType = nSubType
                oUserFieldMD.TableName = strTab
                oUserFieldMD.EditSize = nEditSize
                oUserFieldMD.Mandatory = Mandatory

                If oUserFieldMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD)

            End If

        Catch ex As Exception
            Throw ex
        Finally
            oUserFieldMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    Private Sub AddFields_Link(ByVal strTab As String, _
                            ByVal strCol As String, _
                                ByVal strDesc As String, _
                                    ByVal nType As SAPbobsCOM.BoFieldTypes, _
                                        Optional ByVal i As Integer = 0, _
                                            Optional ByVal nEditSize As Integer = 10, _
                                                Optional ByVal nSubType As SAPbobsCOM.BoFldSubTypes = 0, _
                                                    Optional ByVal Mandatory As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal strLink As String = "")
        Dim oUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try

            If Not (strTab = "OSCN" Or strTab = "OADM" Or strTab = "OPCH" Or strTab = "OITM" Or strTab = "OJDT" Or strTab = "INV1" Or strTab = "RDR1" Or strTab = "OINV" Or strTab = "OCRD" Or strTab = "SPP1" Or strTab = "OWHS") Then
                strTab = "@" + strTab
            End If

            If Not IsColumnExists(strTab, strCol) Then
                oUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

                oUserFieldMD.Description = strDesc
                oUserFieldMD.Name = strCol
                oUserFieldMD.Type = nType
                oUserFieldMD.SubType = nSubType
                oUserFieldMD.TableName = strTab
                oUserFieldMD.EditSize = nEditSize
                oUserFieldMD.Mandatory = Mandatory
                If strLink <> "" Then
                    oUserFieldMD.LinkedTable = strLink
                End If
                If oUserFieldMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD)

            End If

        Catch ex As Exception
            Throw ex
        Finally
            oUserFieldMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    Public Sub addField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
        Dim intLoop As Integer
        Dim strValue, strDesc As Array
        Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try

            strValue = ValidValues.Split(Convert.ToChar(","))
            strDesc = ValidDescriptions.Split(Convert.ToChar(","))
            If (strValue.GetLength(0) <> strDesc.GetLength(0)) Then
                Throw New Exception("Invalid Valid Values")
            End If

            objUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)

            If (Not IsColumnExists(TableName, ColumnName)) Then
                objUserFieldMD.TableName = TableName
                objUserFieldMD.Name = ColumnName
                objUserFieldMD.Description = ColDescription
                objUserFieldMD.Type = FieldType
                If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                    objUserFieldMD.Size = Size
                Else
                    objUserFieldMD.EditSize = Size
                End If
                objUserFieldMD.SubType = SubType
                objUserFieldMD.DefaultValue = SetValidValue
                For intLoop = 0 To strValue.GetLength(0) - 1
                    objUserFieldMD.ValidValues.Value = strValue(intLoop)
                    objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                    objUserFieldMD.ValidValues.Add()
                Next
                If (objUserFieldMD.Add() <> 0) Then
                    '  MsgBox(oApplication.Company.GetLastErrorDescription)
                End If
            Else


            End If

        Catch ex As Exception
            MsgBox(ex.Message)

        Finally

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            GC.Collect()

        End Try


    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : IsColumnExists
    'Parameter          : ByVal Table As String, ByVal Column As String
    'Return Value       : Boolean
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Function to check if the Column already exists in Table
    '**************************************************************************************************************
    Private Function IsColumnExists(ByVal Table As String, ByVal Column As String) As Boolean
        Dim oRecordSet As SAPbobsCOM.Recordset

        Try
            strSQL = "SELECT COUNT(*) FROM CUFD WHERE Upper(""TableID"") = '" & Table.ToString.ToUpper.Trim() & "' AND Upper(""AliasID"") = '" & Column.ToString.ToUpper.Trim() & "'"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strSQL)
            If oRecordSet.Fields.Item(0).Value = 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
            oRecordSet = Nothing
            GC.Collect()
        End Try
    End Function

    Private Sub AddKey(ByVal strTab As String, ByVal strColumn As String, ByVal strKey As String, ByVal i As Integer)
        Dim oUserKeysMD As SAPbobsCOM.UserKeysMD

        Try
            '// The meta-data object must be initialized with a
            '// regular UserKeys object
            oUserKeysMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys)

            If Not oUserKeysMD.GetByKey("@" & strTab, i) Then

                '// Set the table name and the key name
                oUserKeysMD.TableName = strTab
                oUserKeysMD.KeyName = strKey

                '// Set the column's alias
                oUserKeysMD.Elements.ColumnAlias = strColumn
                oUserKeysMD.Elements.Add()
                oUserKeysMD.Elements.ColumnAlias = "RentFac"

                '// Determine whether the key is unique or not
                oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES

                '// Add the key
                If oUserKeysMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If

            End If

        Catch ex As Exception
            Throw ex

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD)
            oUserKeysMD = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try

    End Sub

    '********************************************************************
    'Type		            :   Function    
    'Name               	:	AddUDO
    'Parameter          	:   
    'Return Value       	:	Boolean
    'Author             	:	
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To Add a UDO for Transaction Tables
    '********************************************************************
    Private Sub AddUDO(ByVal strUDO As String, ByVal strDesc As String, ByVal strTable As String, _
                                Optional ByVal sFind1 As String = "", Optional ByVal sFind2 As String = "", _
                                        Optional ByVal strChildTbl As String = "", Optional ByVal strChildTb2 As String = "", Optional ByVal strChildTb3 As String = "", Optional ByVal nObjectType As SAPbobsCOM.BoUDOObjType = SAPbobsCOM.BoUDOObjType.boud_Document, Optional ByVal strChildTb4 As String = "")

        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Try
            oUserObjectMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjectMD.GetByKey(strUDO) = 0 Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES

                If sFind1 <> "" And sFind2 <> "" Then
                    oUserObjectMD.FindColumns.ColumnAlias = sFind1
                    oUserObjectMD.FindColumns.Add()
                    oUserObjectMD.FindColumns.SetCurrentLine(1)
                    oUserObjectMD.FindColumns.ColumnAlias = sFind2
                    oUserObjectMD.FindColumns.Add()
                End If

                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.LogTableName = "A" & strTable
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.ExtensionName = ""
                Dim intTables As Integer = 0
                If strChildTbl <> "" Then
                    oUserObjectMD.ChildTables.TableName = strChildTbl
                End If
                If strChildTb2 <> "" Then
                    If strChildTbl <> "" Then
                        oUserObjectMD.ChildTables.Add()
                        intTables = intTables + 1
                    End If
                    oUserObjectMD.ChildTables.SetCurrentLine(intTables)
                    oUserObjectMD.ChildTables.TableName = strChildTb2
                End If
                If strChildTb3 <> "" Then
                    If strChildTb2 <> "" Then
                        oUserObjectMD.ChildTables.Add()
                        intTables = intTables + 1
                    End If
                    oUserObjectMD.ChildTables.SetCurrentLine(intTables)

                    oUserObjectMD.ChildTables.TableName = strChildTb3
                End If

                If strChildTb4 <> "" Then
                    If strChildTb3 <> "" Then
                        oUserObjectMD.ChildTables.Add()
                        intTables = intTables + 1
                    End If
                    oUserObjectMD.ChildTables.SetCurrentLine(intTables)

                    oUserObjectMD.ChildTables.TableName = strChildTb4
                End If

                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.Code = strUDO
                oUserObjectMD.Name = strDesc
                oUserObjectMD.ObjectType = nObjectType
                oUserObjectMD.TableName = strTable

                If oUserObjectMD.Add() <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If

        Catch ex As Exception
            Throw ex

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
            oUserObjectMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try

    End Sub

#End Region

#Region "Public Functions"
    '*************************************************************************************************************
    'Type               : Public Function
    'Name               : CreateTables
    'Parameter          : 
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Creating Tables by calling the AddTables & AddFields Functions
    '**************************************************************************************************************
    Public Sub CreateTables()
        Dim oProgressBar As SAPbouiCOM.ProgressBar
        Try

            oApplication.Utilities.Message("Initializing Database...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'Reason Code 
            AddTables("Z_RECO", "Reason Code", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_RECO", "Z_Code", "Reason Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_RECO", "Z_Name", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("@Z_RECO", "Z_Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")
            addField("@Z_RECO", "Z_Type", "Reason Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 2, SAPbobsCOM.BoFldSubTypes.st_Address, "SO,PO", "Sales Order,Purchases Order", "SO")

            'Document Type

            'Reason Code 
            AddTables("Z_ODOC", "User Security Document Type", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_ODOC", "Z_Code", "Document Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_ODOC", "Z_Name", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_ODOC", "Z_FormUID", "Form UID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_ODOC", "Z_FieldID", "Item UID", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            addField("@Z_ODOC", "Z_Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")


            'FA Transaction
            AddTables("Z_OFATA", "Fixed Asset Transaction ", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_OFATA", "Z_Code", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_OFATA", "Z_Name", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            ',CostCenter,C
            addField("@Z_OFATA", "Z_TransType", "Transaction Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "L,E", "Location,Employee Transfer", "L")
            AddFields("Z_OFATA", "Z_FromCode", "Transfer From ", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OFATA", "Z_FName", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OFATA", "Z_ToCode", "Transfer To", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OFATA", "Z_TName", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            AddFields("Z_OFATA", "Z_FromDate", "From Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_OFATA", "Z_ToDate", "To Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_OFATA", "Z_DisRule", "Distribution Rule", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)

            addField("@Z_OFATA", "Z_DocStatus", "Document Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "D,N,P,A,R,C,L", "Draft,Confirm,Pending for Approval,Approved,Rejected,Cancel,Close", "D")
            addField("@Z_OFATA", "Z_IsApp", "Approval Required", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            addField("@Z_OFATA", "Z_AppStatus", "Approval Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "P,A,R", "Pending,Approved,Rejected", "P")
            AddFields("Z_OFATA", "Z_CurrApprover", "Current Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OFATA", "Z_NextApprover", "Next Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OFATA", "Z_Attachment", "Attachments", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_OFATA", "Z_AppReqDate", "Required Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_OFATA", "Z_ApproveId", "Approve Template Id", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_OFATA", "Z_Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_OFATA", "Z_AppRemarks", "Approver Remarks", SAPbobsCOM.BoFieldTypes.db_Memo)

            'SO Closing Reason
            AddFields("RDR1", "Z_RECODE", "Closing Reason Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)

            'Delivery Date Updation

            AddFields("OITM", "Z_DelDays", "Delivery Days", SAPbobsCOM.BoFieldTypes.db_Numeric)
            'Week End Master
            AddTables("Z_OWEM", "Week End Master", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_OWEM", "Z_Code", "Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OWEM", "Z_Name", "Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OWEM", "Z_From", "Week End From", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_OWEM", "Z_End", "Week End End", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            addField("@Z_OWEM", "Z_Default", "Default", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            addField("@Z_OWEM", "Z_Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")


            'Map Week End to Customer
            AddFields("OCRD", "Z_WeekEnd", "WeekEnd Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)

            'Unlock Specific Date of posting
            AddTables("Z_OLUSR", "Unlock Specific Date", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddTables("Z_LUSR1", "Date Details", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_OLUSR", "Z_UserCode", "User Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OLUSR", "Z_UserName", "User Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            addField("@Z_OLUSR", "Z_Super", "Super User", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")


            AddFields("Z_LUSR1", "Z_DocType", "Document Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_LUSR1", "Z_Date", "Posting Date", SAPbobsCOM.BoFieldTypes.db_Date)
            addField("@Z_LUSR1", "Z_Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")

            ''Approval Tables
            AddTables("Z_OAPPT", "Approval Template", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddTables("Z_APPT1", "Approval Orginator", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddTables("Z_APPT2", "Approval Authorizer", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddTables("Z_APPT3", "Department Authorizer", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)

            AddFields("Z_OAPPT", "Z_Code", "Approval Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_OAPPT", "Z_Name", "Approval Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OAPPT", "Z_DocType", "Document Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_OAPPT", "Z_DocDesc", "Document Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OAPPT", "Z_Active", "Active Template", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)

            AddFields("Z_APPT1", "Z_EmpId", "Orginator Id", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_APPT1", "Z_OName", "Orginator Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            'AddFields("Z_APPT1", "Z_EmpID", "T&A Employee Id", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)

            AddFields("Z_APPT3", "Z_DeptCode", "Department Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_APPT3", "Z_DeptName", "Department Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)

            AddFields("Z_APPT2", "Z_AUser", "Authorizer Id", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_APPT2", "Z_AName", "Authorizer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_APPT2", "Z_AMan", "Mandatory", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_APPT2", "Z_AFinal", "Final Stage", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)

            ''Approval History Table
            AddTables("Z_APHIS", "Approval History", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_APHIS", "Z_DocEntry", "Document Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_APHIS", "Z_DocType", "Document Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_APHIS", "Z_EmpId", "Employee Id", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_APHIS", "Z_EmpName", "Employee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            addField("@Z_APHIS", "Z_AppStatus", "Approved Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "P,A,R", "Pending,Approved,Rejected", "P")
            AddFields("Z_APHIS", "Z_Remarks", "Comments", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_APHIS", "Z_ApproveBy", "Approved By", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_APHIS", "Z_Approvedt", "Approver Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_APHIS", "Z_ADocEntry", "Template DocEntry", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_APHIS", "Z_ALineId", "Template LineId", SAPbobsCOM.BoFieldTypes.db_Numeric)

            ''Email Details Table
            AddTables("Z_OMAIL", "Email SetUp Details", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_OMAIL", "Z_SMTPSERV", "SMTP SERVER", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OMAIL", "Z_SMTPPORT", "SMTP PORT", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_OMAIL", "Z_SMTPUSER", "SMTP USER", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OMAIL", "Z_SMTPPWD", "SMTP PASSWORD", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OMAIL", "Z_SSL", "SMTP SSL", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)


            'Item Category 
            AddTables("Z_OITC", "Item Category", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_OITC", "Z_Code", "Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OITC", "Z_Name", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("@Z_OITC", "Z_Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")


            'BP Category 
            AddTables("Z_OBPC", "Business Partner Category", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_OBPC", "Z_Code", " Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OBPC", "Z_Name", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("@Z_OBPC", "Z_Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")

            'Warehouse Category 
            AddTables("Z_OWHC", "Warehouse Category", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_OWHC", "Z_Code", "Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OWHC", "Z_Name", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("@Z_OWHC", "Z_Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")

            AddFields("OITM", "Z_ITCCODE", "Item Categrory", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("OCRD", "Z_BPCCODE", "Business Partner Categrory", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("OWHS", "Z_WHSCODE", "Warehouse Categrory", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)

            'User Mapping 
            AddTables("Z_LUSR2", "Item Categories", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_LUSR2", "Z_Code", "Item Category Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_LUSR2", "Z_Name", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            addField("@Z_LUSR2", "Z_Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")

            AddTables("Z_LUSR3", "BP Categories", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_LUSR3", "Z_Code", "BP Category Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_LUSR3", "Z_Name", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            addField("@Z_LUSR3", "Z_Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")

            AddTables("Z_LUSR4", "Warehouse Categories", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_LUSR4", "Z_Code", "Warehouse Category Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_LUSR4", "Z_Name", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            addField("@Z_LUSR4", "Z_Active", "Active", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")

            'Yearly Volume Posting
            AddTables("Z_OYVP", "Yearly Volume Posting", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_OYVP", "Z_Debit", "Debit Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OYVP", "Z_Credit", "Credit Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OYVP", "Z_TaxDebit", "Taxable Debit Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            AddFields("OCRD", "Z_OYVP", "Yearly Volume Rebate%", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            addField("OINV", "Z_OYVP", "Rebate Posting", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")

            AddFields("OITM", "Z_USERCODE", "Mapped User", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("OCRD", "Z_USERCODE", "Mapped User", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("OWHS", "Z_USERCODE", "Mapped User", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("OJDT", "Z_ARInvoice", "A/R Invoice Ref.No", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("OJDT", "Z_BaseEntry", "A/R Invoice DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("OINV", "Z_JournalRef", "Journal Entry Ref.No", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)


            'Web Application Related UDF's

            addField("ORDR", "Z_Source", "Document Source", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "W,R", "Web,Regular", "R")
            AddTables("Z_OWRE", "Sales Web Reference", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("OCRD", "Z_DefUoM", "Default Web UoM", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("OSCN", "UoM", "Credit Note UoM", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            addField("OSCN", "IsReturn", "Allow Return", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")


            'Supplier Price List Document
            AddTables("Z_OVPL", "Supplier Price List", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_OVPL", "Z_CardCode", "Supplier Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OVPL", "Z_CardName", "Supplier Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OVPL", "Z_ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OVPL", "Z_ItemName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("Z_OVPL", "Z_CPrice", "Current Price", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price, "", "", "")
            AddFields("Z_OVPL", "Z_CCurrency", "Current Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 3)
            addField("Z_OVPL", "Z_UPrice", "Updated Price", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price, "", "", "")
            AddFields("Z_OVPL", "Z_UCurrency", "Updated Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 3)
            AddFields("Z_OVPL", "Z_Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo)

            addField("@Z_OVPL", "Z_DocStatus", "Document Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "D,N,P,A,R,C,L", "Draft,Confirm,Pending for Approval,Approved,Rejected,Cancel,Close", "D")
            addField("@Z_OVPL", "Z_IsApp", "Approval Required", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            addField("@Z_OVPL", "Z_AppStatus", "Approval Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "P,A,R", "Pending,Approved,Rejected", "P")
            AddFields("Z_OVPL", "Z_CurrApprover", "Current Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OVPL", "Z_NextApprover", "Next Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OVPL", "Z_Attachment", "Attachments", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_OVPL", "Z_AppReqDate", "Required Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_OVPL", "Z_ApproveId", "Approve Template Id", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_OVPL", "Z_AppRemarks", "Approver Remarks", SAPbobsCOM.BoFieldTypes.db_Memo)

            addField("OSCN", "Z_CPrice", "Current Price", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price, "", "", "")
            AddFields("OSCN", "Z_CCurrency", "Current Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 3)
            addField("OSCN", "Z_UPrice", "Previous Price", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price, "", "", "")
            AddFields("OSCN", "Z_UCurrency", "Previous Currency", SAPbobsCOM.BoFieldTypes.db_Alpha, , 3)

            AddFields("OSCN", "Z_ReqBy", "Requested by", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("OSCN", "Z_ReqDate", "Requested Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("OSCN", "Z_AppBy", "Approved by", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("OSCN", "Z_AppDate", "Approved Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("OSCN", "Z_DocNum", "Reference Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)


            AddFields("OINV", "Z_DelDate", "Delivery Date", SAPbobsCOM.BoFieldTypes.db_Date)
            addField("OINV", "Z_IsDel", "Is Delivered", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("OINV", "Z_DelRef", "Delivery Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("OINV", "Z_ScnUser", "Scanned User", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)


            AddFields("ORCT", "Z_DelDate", "Delivery Date", SAPbobsCOM.BoFieldTypes.db_Date)
            addField("ORCT", "Z_IsDel", "Is Delivered", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("ORCT", "Z_DelRef", "Delivery Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("ORCT", "Z_ScnUser", "Scanned User", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("OVPM", "Z_DelDate", "Delivery Date", SAPbobsCOM.BoFieldTypes.db_Date)
            addField("OVPM", "Z_IsDel", "Is Delivered", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("OVPM", "Z_DelRef", "Delivery Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("OVPM", "Z_ScnUser", "Scanned User", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddTables("Z_ODEL", "Delivery Document Header", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_ODEL", "Z_DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_ODEL", "Z_DelDate", "Delivery Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_ODEL", "Z_ScanNo", "Scan Area No", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_ODEL", "Z_Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo)

            AddTables("Z_DEL1", "Delivery Document Details", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_DEL1", "Z_DocEntry", "Doc Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_DEL1", "Z_DocNo", "Document No", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_DEL1", "Z_CardCode", "Card Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_DEL1", "Z_CardName", "Supplier Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_DEL1", "Z_Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_DEL1", "Z_DocDate", "Posting Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_DEL1", "Z_DocType", "Document Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_DEL1", "Z_TarTable", "Target table", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)


            'Promotion Template
            AddTables("Z_OPRM", "Promotion Template", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_OPRM", "Z_PrCode", "Promotion Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_OPRM", "Z_PrName", "Promotion Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_OPRM", "Z_EffFrom", "Effective From", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_OPRM", "Z_EffTo", "Effective To", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_OPRM", "Z_Active", "Is Active", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_OPRM", "Z_Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo)

            'Promotion Items
            AddTables("Z_PRM1", "Promotion Items", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_PRM1", "Z_ItmCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_PRM1", "Z_ItmName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_PRM1", "Z_Qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("Z_PRM1", "Z_UOMGroup", "UOM Group", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            addField("Z_PRM1", "Z_DisType", "Discount Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "D,I", "Discount,Item", "D")
            AddFields("Z_PRM1", "Z_Dis", "Discount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            AddFields("Z_PRM1", "Z_OffCode", "Offer Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_PRM1", "Z_OffName", "Offer Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_PRM1", "Z_OQty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("Z_PRM1", "Z_ODis", "Discount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            AddFields("Z_PRM1", "Z_OUOMGroup", "UOM Group", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)

            'Commission Charges Reference Table
            AddTables("Z_OCPR", "Customer Promotion Mapping", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_OCPR", "Z_CustCode", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_OCPR", "Z_PrCode", "Promotion Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_OCPR", "Z_EffFrom", "Effective From", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_OCPR", "Z_EffTo", "Effective To", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_OCPR", "Z_Active", "Is Active", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)

            'AddFields("Z_OCPR", "Z_ItmCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            'AddFields("Z_OCPR", "Z_ItmName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            'AddFields("Z_OCPR", "Z_Qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            'AddFields("Z_OCPR", "Z_UOMGroup", "UOM Group", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)

            'AddFields("Z_OCPR", "Z_OffCode", "Offer Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            'AddFields("Z_OCPR", "Z_OffName", "Offer Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            'AddFields("Z_OCPR", "Z_OQty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            'AddFields("Z_OCPR", "Z_ODis", "Discount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            'AddFields("Z_OCPR", "Z_OUOMGroup", "Offer UOM Group", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)

            AddTables("Z_OPRF", "Promotion Reference", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddTables("Z_OPRE", "Promotion Reason", SAPbobsCOM.BoUTBTableType.bott_NoObject)

            AddFields("RDR1", "Z_PrCode", "Promotion Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            addField("RDR1", "Z_PrmApp", "Promotion Applied", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            AddFields("RDR1", "Z_PrRef", "Promotion Ref", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            addField("RDR1", "Z_IType", "Item Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "R,F", "Regular,Free", "R")
            AddFields_Link("RDR1", "Z_PrReason", "Promotion Reason", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, "Z_OPRE")

            'Inventory Count Header
            AddTables("Z_OICT", "Inventory Count Header", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_OICT", "Z_CntDate", "Count Date", SAPbobsCOM.BoFieldTypes.db_Date)
            addField("Z_OICT", "Z_Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "O,L,C", "Open,Cancelled,Close", "O")
            AddFields("Z_OICT", "Z_Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo)

            'Inventory Count Lines
            AddTables("Z_ICT1", "Inventory Count Items", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_ICT1", "Z_ItmCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_ICT1", "Z_ItmName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_ICT1", "Z_WareHouse", "Whs Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_ICT1", "Z_UOM", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_ICT1", "Z_Qty", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            addField("Z_ICT1", "Z_Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "O,L,C", "Open,Cancelled,Close", "O")
            AddFields("Z_ICT1", "Z_Remarks", "Line Remarks", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_ICT1", "Z_ICRef", "Inventory Count Ref", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_ICT1", "Z_IUOM", "Inventory UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_ICT1", "Z_IQty", "Inventory Quantity", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)

            AddFields("OINC", "Z_ICTREF", "Inventory Count Ref", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("OINC", "Z_ICTREFL", "Inventory Count Ref L", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)

            'ESS Fields
            addField("OCRD", "Z_AllReturn", "Allow Returns", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            addField("OCRD", "Z_AllStAcc", "Allow Statement of Accounts", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_None, "N,Y", "No,Yes", "N")
            AddFields("ORDR", "Z_Noofdays", "No.of days in Del/post date", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("RDR1", "Z_Noofdays", "No.of days in Del/post date", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("DRF1", "Z_Noofdays", "No.of days in Del/post date", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("ODRF", "Z_Noofdays", "No.of days in Del/post date", SAPbobsCOM.BoFieldTypes.db_Numeric)

            'Item Category -2016-01-05
            addField("OITM", "Z_Identifier", "Identifier", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "F,V", "Fixed,Variable", "F")

            AddFields("OCRD", "Driver", "Driver Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)

            CreateUDO()
        Catch ex As Exception
            Throw ex
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    Public Sub CreateUDO()
        Try

            AddUDO("Z_APHIS", "Approval History", "Z_APHIS", "DocEntry", "U_Z_DocEntry", , , , SAPbobsCOM.BoUDOObjType.boud_Document)
            AddUDO("Z_OAPPT", "Template", "Z_OAPPT", "DocEntry", "U_Z_Code", "Z_APPT1", "Z_APPT2", "Z_APPT3", SAPbobsCOM.BoUDOObjType.boud_Document)
            AddUDO("Z_OFATA", "Asset_Transaction", "Z_OFATA", "DocNum", "U_Z_Code", , , , SAPbobsCOM.BoUDOObjType.boud_Document)
            AddUDO("Z_RECO", "ReasonCode", "Z_RECO", "U_Z_Code", "DocNum", , , , SAPbobsCOM.BoUDOObjType.boud_Document)
            AddUDO("Z_OWEM", "WeekEnd_Master", "Z_OWEM", "U_Z_Code", "DocNum", , , , SAPbobsCOM.BoUDOObjType.boud_Document)
            AddUDO("Z_OUSR", "User_Mapping", "Z_OLUSR", "DocNum", "U_Z_UserCode", "Z_LUSR1", "Z_LUSR2", "Z_LUSR3", SAPbobsCOM.BoUDOObjType.boud_Document, "Z_LUSR4")
            AddUDO("Z_ODOC", "Document Type", "Z_ODOC", "DocNum", "U_Z_Code", , , , SAPbobsCOM.BoUDOObjType.boud_Document)

            AddUDO("Z_OITC", "ItemCategory", "Z_OITC", "U_Z_Code", "DocNum", , , , SAPbobsCOM.BoUDOObjType.boud_Document)
            AddUDO("Z_OBPC", "BPCategory", "Z_OBPC", "U_Z_Code", "DocNum", , , , SAPbobsCOM.BoUDOObjType.boud_Document)
            AddUDO("Z_OWHC", "WarehouseCategory", "Z_OWHC", "U_Z_Code", "DocNum", , , , SAPbobsCOM.BoUDOObjType.boud_Document)
            AddUDO("Z_OVPL", "Vendor_Price", "Z_OVPL", "U_Z_CardCode", "DocNum", , , , SAPbobsCOM.BoUDOObjType.boud_Document)
            AddUDO("Z_ODEL", "Delivery_Document", "Z_ODEL", "DocEntry", "U_Z_DelDate", "Z_DEL1", , , SAPbobsCOM.BoUDOObjType.boud_Document)
            AddUDO("Z_OPRM", "Promotion_Template", "Z_OPRM", "U_Z_PrCode", "U_Z_PrName", "Z_PRM1", , , SAPbobsCOM.BoUDOObjType.boud_Document)
            AddUDO("Z_OICT", "Inventory_Count", "Z_OICT", "DocNum", "U_Z_CntDate", "Z_ICT1", , , SAPbobsCOM.BoUDOObjType.boud_Document)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

End Class
