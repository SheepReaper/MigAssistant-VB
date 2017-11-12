Imports System.ComponentModel

Public Class AdvancedSettingsForm

    Private Sub BtnAdvancedSettingsClose_Click(sender As Object, e As EventArgs) Handles btnAdvancedSettingsClose.Click

        ' Close Form
        Close()

    End Sub

    Private Sub Form_Closing(sender As Object, e As CancelEventArgs) Handles MyBase.Closing

        If rbnAdvancedSettingsQuestion4B.Checked Then
            bln_MigrationSettingsLocalAccounts = True
        Else
            bln_MigrationSettingsLocalAccounts = False
        End If

        ' Stop the form from actually closing, and hide instead
        e.Cancel = True
        Hide()

    End Sub

    Private Sub FormMigrationAdvancedSettings_Load(sender As Object, e As EventArgs) Handles MyBase.Load, MyBase.VisibleChanged
        ' If Encryption is disabled, hide on the settings page
        If bln_MigrationEncryptionDisabled Then
            lblAdvancedSettingsQuestion2.Visible = False
            rbnAdvancedSettingsQuestion2A.Visible = False
            rbnAdvancedSettingsQuestion2B.Visible = False
        End If

        lblAdvancedSettingsQuestion4.Visible = False
        rbnAdvancedSettingsQuestion4A.Visible = False
        rbnAdvancedSettingsQuestion4B.Visible = False

        ' If performing a backup...
        If str_MigrationType = "SCANSTATE" Then
            ' and migrating more than the current user, all local account migration too
            If bln_MigrationSettingsAllUsers Then
                lblAdvancedSettingsQuestion4.Visible = True
                rbnAdvancedSettingsQuestion4A.Visible = True
                rbnAdvancedSettingsQuestion4B.Visible = True
            End If
        End If
    End Sub

    Private Sub RbnAdvancedSettingsQuestion1_CheckedChanged(sender As Object, e As EventArgs) Handles rbnAdvancedSettingsQuestion1B.CheckedChanged, rbnAdvancedSettingsQuestion1A.CheckedChanged

        If rbnAdvancedSettingsQuestion1B.Checked Then
            If Not bln_MigrationLocationUseOther Then
                If fbdAdvancedSettingsDataStore.ShowDialog(Me) = DialogResult.OK Then
                    bln_MigrationLocationUseOther = True
                    str_MigrationLocationOther = fbdAdvancedSettingsDataStore.SelectedPath
                Else
                    rbnAdvancedSettingsQuestion1A.Checked = True
                End If
            End If
        Else
            bln_MigrationLocationUseOther = False
        End If

    End Sub

    Private Sub RbnAdvancedSettingsQuestion2_CheckedChanged(sender As Object, e As EventArgs) Handles rbnAdvancedSettingsQuestion2B.CheckedChanged, rbnAdvancedSettingsQuestion2A.CheckedChanged

        If rbnAdvancedSettingsQuestion2B.Checked Then
            If Not bln_MigrationEncryptionCustom Then
                bln_MigrationEncryptionCustom = True
                CustomEncryptionForm.ShowDialog()
            End If
        Else
            bln_MigrationEncryptionCustom = False
        End If

    End Sub

End Class