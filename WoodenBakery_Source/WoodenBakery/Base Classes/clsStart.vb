Public Class clsStart
    
    Shared Sub Main()
        Dim oRead As System.IO.StreamReader
        Dim LineIn, strUsr, strPwd As String
        Dim i As Integer
        Try
            Try
                oApplication = New clsListener
                oApplication.Utilities.Connect()
                ' oApplication.SetFilter()
                With oApplication.Company.GetCompanyService
                    CompanyDecimalSeprator = .GetAdminInfo.DecimalSeparator
                    CompanyThousandSeprator = .GetAdminInfo.ThousandsSeparator
                End With
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
                Exit Sub
            End Try
            oApplication.Utilities.CreateTables()
            oApplication.Utilities.AddRemoveMenus("Menu.xml")
            oApplication.Utilities.createHRMainAuthorization()
            oApplication.Utilities.AuthorizationCreation()

            Dim oMenuItem1 As SAPbouiCOM.MenuItem
            oMenuItem1 = oApplication.SBO_Application.Menus.Item("W_2000")
            oMenuItem1.Image = Application.StartupPath & "\Inv.bmp"

            If oApplication.Company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                blnIsHanaDB = True
            End If

            oApplication.Utilities.Message("Wooden Bakery Addon Connected successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            oApplication.Utilities.NotifyAlert()
            System.Windows.Forms.Application.Run()
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try

    End Sub

End Class
