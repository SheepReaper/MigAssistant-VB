Imports System.ComponentModel

Public Class FormMigrationAdvancedSettings
    Private Sub btnAdvancedSettingsClose_Click(sender As Object, e As EventArgs) Handles btnAdvancedSettingsClose.Click

        ' Close Form
        Close()
    End Sub

    Private Sub form_Closing(sender As Object, e As CancelEventArgs) Handles MyBase.Closing

        If rbnAdvancedSettingsQuestion4B.Checked Then
            BlnMigrationSettingsLocalAccounts = True
        Else
            BlnMigrationSettingsLocalAccounts = False
        End If

        ' Stop the form from actually closing, and hide instead
        e.Cancel = True
        Hide()
    End Sub

    Private Sub formMigrationAdvancedSettings_Load(sender As Object, e As EventArgs) _
        Handles MyBase.Load, MyBase.VisibleChanged
        ' If Encryption is disabled, hide on the settings page
        If BlnMigrationEncryptionDisabled Then
            lblAdvancedSettingsQuestion2.Visible = False
            rbnAdvancedSettingsQuestion2A.Visible = False
            rbnAdvancedSettingsQuestion2B.Visible = False
        End If

        lblAdvancedSettingsQuestion4.Visible = False
        rbnAdvancedSettingsQuestion4A.Visible = False
        rbnAdvancedSettingsQuestion4B.Visible = False

        ' If performing a backup...
        If StrMigrationType = "SCANSTATE" Then
            ' and migrating more than the current user, all local account migration too
            If BlnMigrationSettingsAllUsers Then
                lblAdvancedSettingsQuestion4.Visible = True
                rbnAdvancedSettingsQuestion4A.Visible = True
                rbnAdvancedSettingsQuestion4B.Visible = True
            End If
        End If
    End Sub

    Private Sub rbnAdvancedSettingsQuestion1_CheckedChanged(sender As Object, e As EventArgs) _
        Handles rbnAdvancedSettingsQuestion1B.CheckedChanged, rbnAdvancedSettingsQuestion1A.CheckedChanged

        If rbnAdvancedSettingsQuestion1B.Checked Then
            If Not BlnMigrationLocationUseOther Then
                If fbdAdvancedSettingsDataStore.ShowDialog(Me) = DialogResult.OK Then
                    BlnMigrationLocationUseOther = True
                    StrMigrationLocationOther = fbdAdvancedSettingsDataStore.SelectedPath
                Else
                    rbnAdvancedSettingsQuestion1A.Checked = True
                End If
            End If
        Else
            BlnMigrationLocationUseOther = False
        End If
    End Sub

    Private Sub rbnAdvancedSettingsQuestion2_CheckedChanged(sender As Object, e As EventArgs) _
        Handles rbnAdvancedSettingsQuestion2B.CheckedChanged, rbnAdvancedSettingsQuestion2A.CheckedChanged

        If rbnAdvancedSettingsQuestion2B.Checked Then
            If Not BlnMigrationEncryptionCustom Then
                BlnMigrationEncryptionCustom = True
                FormCustomEncryption.ShowDialog()
            End If
        Else
            BlnMigrationEncryptionCustom = False
        End If
    End Sub
End Class