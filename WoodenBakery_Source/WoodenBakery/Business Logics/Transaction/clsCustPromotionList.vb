Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsCustPromotionList
    Inherits clsBase
    Private oGrid As SAPbouiCOM.Grid
    Private oDtPromotionList As SAPbouiCOM.DataTable
    Private strQuery As String
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub LoadForm(ByVal strCust As String)
        Try
            oForm = oApplication.Utilities.LoadForm(xml_Z_CPRL, frm_Z_CPRL)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            initialize(oForm, strCust)
            oForm.Freeze(False)
        Catch ex As Exception
            Throw ex
        Finally
            oForm.Freeze(False)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Z_CPRL Then
                Select Case pVal.BeforeAction
                    Case True

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
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

#Region "Function"
    Private Sub initialize(ByVal oForm As SAPbouiCOM.Form, ByVal strCust As String)
        Try
            oGrid = oForm.Items.Item("3").Specific
            strQuery = " Select Distinct T0.""U_Z_PrCode"",T0.""U_Z_PrName"",T0.""U_Z_EffFrom"",T0.""U_Z_EffTo"",""U_Z_ItmCode"",""U_Z_ItmName"" "
            strQuery += ",""U_Z_Qty"",""U_Z_UOMGroup"",""U_Z_OffCode"",""U_Z_OffName"",""U_Z_OQty"",""U_Z_OUOMGroup"",""U_Z_ODis"" From ""@Z_OPRM"" T0 "
            strQuery += " JOIN ""@Z_PRM1"" T1 On T0.""DocEntry"" = T1.""DocEntry"" "
            strQuery += " JOIN ""@Z_OCPR"" T2 On T0.""U_Z_PrCode"" = T2.""U_Z_PrCode"" "
            strQuery += " Where T2.""U_Z_CustCode"" = '" + strCust + "' "
            oForm.DataSources.DataTables.Add("dtPromotionList")
            oDtPromotionList = oForm.DataSources.DataTables.Item(0)
            oDtPromotionList.ExecuteQuery(strQuery)
            oGrid.DataTable = oDtPromotionList

            'Format
            oGrid.Columns.Item("U_Z_PrCode").TitleObject.Caption = "Promotion Code"
            oGrid.Columns.Item("U_Z_PrName").TitleObject.Caption = "Project Name"
            oGrid.Columns.Item("U_Z_EffFrom").TitleObject.Caption = "Effective From"
            oGrid.Columns.Item("U_Z_EffTo").TitleObject.Caption = "Effective To"
            oGrid.Columns.Item("U_Z_ItmCode").TitleObject.Caption = "Item Code"
            oEditTextColumn = oGrid.Columns.Item("U_Z_ItmCode")
            oEditTextColumn.LinkedObjectType = "4"
            oGrid.Columns.Item("U_Z_ItmName").TitleObject.Caption = "Item Name"
            oGrid.Columns.Item("U_Z_Qty").TitleObject.Caption = "Quantity"
            oGrid.Columns.Item("U_Z_Qty").RightJustified = True
            oGrid.Columns.Item("U_Z_OffCode").TitleObject.Caption = "Offer Item"
            oEditTextColumn = oGrid.Columns.Item("U_Z_OffCode")
            oEditTextColumn.LinkedObjectType = "4"
            oGrid.Columns.Item("U_Z_OffName").TitleObject.Caption = "Offer Name"
            oGrid.Columns.Item("U_Z_OQty").TitleObject.Caption = "Offer Qty"
            oGrid.Columns.Item("U_Z_OQty").RightJustified = True
            oGrid.Columns.Item("U_Z_ODis").TitleObject.Caption = "Offer Discount %"
            oGrid.Columns.Item("U_Z_ODis").RightJustified = True
            oGrid.Columns.Item("U_Z_UOMGroup").TitleObject.Caption = "UOM Group"
            oGrid.Columns.Item("U_Z_OUOMGroup").TitleObject.Caption = "UOM Group"

            'Collapse Level By Project
            oGrid.CollapseLevel = 1

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub loadValues(ByVal oForm As SAPbouiCOM.Form, ByVal blnStatus As Boolean)
        Try

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

End Class
