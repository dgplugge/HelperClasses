Imports HelperClasses
Imports System.Drawing

Public Class frmUninstallUpgrade

    Public FormUnInstallString As String = ""
    Public FormInstallString As String = ""

    Private Sub ReflectUpgradeMessage()

        If FormUnInstallString = "" Then FormUnInstallString = SoftwareHandler.ActiveUninstallString.ToString
        Dim info = FormUnInstallString.Split(" ")

        lblCurrent.Font = New Font(System.Drawing.FontFamily.GenericSansSerif, 10)
        lblCurrent.ForeColor = Color.DarkRed
        lblCurrent.Text = String.Format("This application requires an uninstall in order to upgrade.")

        lblCurrentVal.Font = New Font(System.Drawing.FontFamily.GenericSansSerif, 10)
        lblCurrentVal.ForeColor = Color.DarkGreen
        lblCurrentVal.Text = String.Format(SoftwareHandler.AssemblyName)

        lblNew.Font = New Font(System.Drawing.FontFamily.GenericSansSerif, 10)
        lblNew.ForeColor = Color.DarkRed
        lblNew.Text = String.Format("Please select the Uninstall button.{0}", vbCrLf)
        lblNew.Text += String.Format("The application specified below will then be installed as a replacement.")

        lblNewVal.Font = New Font(System.Drawing.FontFamily.GenericSansSerif, 10)
        lblNewVal.ForeColor = Color.DarkGreen
        lblNewVal.Text = String.Format(FormInstallString)

        cmdUninstall.Enabled = True

    End Sub

    Private Sub frmUninstallUpgrade_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ReflectUpgradeMessage()

    End Sub

    Private Sub cmdUninstall_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUninstall.Click

        SoftwareUpgrade.UnInstallApp()
        If SoftwareUpgrade.UnInstallSuccessfull Then
            SoftwareUpgrade.InstallApp(FormInstallString)
        End If

        Me.Close()

    End Sub
End Class