Imports System.ComponentModel

Public Class FormCustomEncryption
    Private Sub btnAdvancedSettingsClose_Click(sender As Object, e As EventArgs) Handles btnCustomEncryptionClose.Click

        Close()
    End Sub

    Public Sub form_Closing(sender As Object, e As CancelEventArgs) Handles MyBase.Closing

        Select Case StrMigrationType
            Case "SCANSTATE"
                If Not FormMigrationAdvancedSettings.rbnAdvancedSettingsQuestion2A.Checked Then
                    If tbxCustomEncryptionKey1.Text.Length = 0 And tbxCustomEncryptionKey2.Text.Length = 0 Then
                        MsgBox(My.Resources.encryptionNoKeySpecified, MsgBoxStyle.Information, My.Resources.appTitle)
                        FormMigrationAdvancedSettings.rbnAdvancedSettingsQuestion2A.Checked = True
                    ElseIf tbxCustomEncryptionKey1.Text <> tbxCustomEncryptionKey2.Text Then
                        MsgBox(My.Resources.encryptionKeysDontMatch, MsgBoxStyle.Exclamation, My.Resources.appTitle)
                        tbxCustomEncryptionKey1.Text = Nothing
                        tbxCustomEncryptionKey2.Text = Nothing
                        tbxCustomEncryptionKey1.Focus()
                        e.Cancel = True
                        Exit Sub
                    End If
                    StrCustomEncryptionKey = tbxCustomEncryptionKey1.Text

                    ' Stop the form from actually closing, and hide instead
                    e.Cancel = True
                    Hide()
                End If
            Case "LOADSTATE"
                If tbxCustomEncryptionKey1.Text.Length = 0 And tbxCustomEncryptionKey2.Text.Length = 0 Then
                    e.Cancel = True
                    Exit Sub
                ElseIf tbxCustomEncryptionKey1.Text <> tbxCustomEncryptionKey2.Text Then
                    MsgBox(My.Resources.encryptionKeysDontMatch, MsgBoxStyle.Exclamation, My.Resources.appTitle)
                    tbxCustomEncryptionKey1.Text = Nothing
                    tbxCustomEncryptionKey2.Text = Nothing
                    tbxCustomEncryptionKey1.Focus()
                    e.Cancel = True
                    Exit Sub
                End If
                StrCustomEncryptionKey = tbxCustomEncryptionKey1.Text

                ' Stop the form from actually closing, and hide instead
                e.Cancel = True
                Hide()

        End Select
    End Sub
End Class