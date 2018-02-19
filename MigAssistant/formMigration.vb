Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.IO
Imports System.Security.AccessControl
Imports System.Xml
Imports Microsoft.VisualBasic.FileIO
Imports MigAssistant.My.Resources
Imports SearchOption = Microsoft.VisualBasic.FileIO.SearchOption


Public Class FormMigration

#Region "Fields"

    Private WithEvents _classHealthCheck As ClassHealthCheck

    ' Set up the classes with events
    Private WithEvents _classMigration As ClassMigration

    ' Set up the USB Event watcher
    Private WithEvents _eventwatcherUsbStateChange As ManagementEventWatcher

#End Region

#Region "Methods"

    ' Data Capture - Advanced Settings
    Private Sub btnMigrationSettingsAdvanced_Click(sender As Object, e As EventArgs) _
        Handles button_AdvancedSettings.Click

        Sub_DebugMessage()
        Sub_DebugMessage("* Advanced Settings Button Events *")

        FormMigrationAdvancedSettings.Show()
        button_Start.Focus()
    End Sub

    ' Handle Start / Stop Buttons
    Private Sub button_StartStopClick(sender As Object, e As EventArgs) Handles button_Start.Click

        Sub_DebugMessage()
        Sub_DebugMessage("* Start / Stop Button Click Events *")

        If Not BlnHealthCheckInProgress And Not BlnMigrationInProgress Then
            Sub_DebugMessage("No actions currently in progress")
            sub_MigrationInitialise()
        Else
            If BlnHealthCheckInProgress Then
                Sub_DebugMessage("Health Check currently in progress")
                _classHealthCheck.SpinDown()
            End If
            If BlnMigrationInProgress Then
                Sub_DebugMessage("Migration currently in progress")
                _classMigration.SpinDown()
            End If
        End If
    End Sub

    ' Data Capture - Health Check Skip Checkbox
    Private Sub checkbox_HealthCheck_CheckedChanged(sender As Object, e As EventArgs) _
        Handles checkbox_HealthCheck.CheckedChanged

        Sub_DebugMessage()
        Sub_DebugMessage("* Health Check Checkbox Events *")

        If checkbox_HealthCheck.Checked Then
            Sub_DebugMessage("Health Check Enabled")
            BlnHealthCheck = True
        Else
            Sub_DebugMessage("Health Check Disabled")
            BlnHealthCheck = False
        End If
    End Sub

    ' Form Closing
    Private Sub form_MigrationClosing(sender As Object, e As CancelEventArgs) Handles MyBase.Closing

        Sub_DebugMessage()
        Sub_DebugMessage("* Form Closing Events *")

        'Exit the application
        AppShutdown(CInt(exitCodeOk))
    End Sub

    ' Form Loading
    Private Sub form_MigrationLoad(sender As Object, e As EventArgs) Handles MyBase.Load

        AppInitialise()

        Sub_DebugMessage()
        Sub_DebugMessage("* Form Startup Events *")

        ' *** Set up form options
        Sub_DebugMessage("Setting form defaults...")

        Text = $"{appTitle} {appBuild}"

        ' Set the default Migration option
        StrMigrationType = "SCANSTATE"

        ' Check the OS Architecture and set options accordingly
        Select Case StrOsArchitecture
            Case "x86"
                StrUsmtguid = StrUsmtguiDx86
            Case "x64"
                StrUsmtguid = StrUsmtguiDx64
        End Select

        ' Get OS Information and display on the form
        label_OSVersion.Text = StrOsFullName

        ' Set Multi User Mode if desired
        If BlnMigrationMultiUserMode And Not radiobox_WorkstationDetails2.Checked Then
            radiobox_WorkstationDetails2.Checked = True
        End If

        ' If Workstation Details is disabled through settings, disable the form controls
        If BlnSettingsWorkstationDetailsDisabled Then
            slabel_WorkstationDetails.Enabled = False
            radiobox_WorkstationDetails1.Enabled = False
            radiobox_WorkstationDetails2.Enabled = False
        End If

        ' If Health Check is enabled through settings, enable on the form controls
        If BlnSettingsHealthCheckDefaultEnabled Then
            checkbox_HealthCheck.Checked = True
            BlnHealthCheck = True
        End If

        ' If Advanced Settings is disabled through settings, disable the form controls
        If BlnSettingsAdvancedSettingsDisabled Then
            button_AdvancedSettings.Visible = False
        Else
            button_AdvancedSettings.Visible = True
        End If

        ' Reset status labels
        label_MigrationCurrentPhase.Text = Elipses
        label_MigrationEstSize.Text = Elipses
        label_MigrationEstTimeRemaining.Text = Elipses

        ' Check for USB devices and initialise USB Event watcher
        sub_USBInitialise()

        ' Check the command-line for actioning
        sub_InitParseCommandLine()
    End Sub

    ' Form Showing
    Private Sub form_MigrationShown(sender As Object, e As EventArgs) Handles MyBase.Shown

        Sub_DebugMessage()
        Sub_DebugMessage("* Form Shown Events *")

        ' Check if the Key is encrypted in the settings file
        'sub_InitCheckEncryption()

        ' Check if USMT is installed
        Sub_InitCheckForUSMT()

        ' Check if Config files exist
        sub_InitCheckForRuleSetFiles()

        ' Check if the WMA configuration is valid
        If Not func_InitCheckForValidConfiguration() Then
            Sub_DebugMessage(
                "WARNING: It looks like you haven't yet modified the WMA configuration file (MigAssistant.Exe.Config). Your migration may fail as a result!",
                True)
            ' If we're running in Progress Only mode, start the migration immediately
        ElseIf BlnAppProgressOnlyMode Then
            Sub_DebugMessage("Progress Only Mode. Starting Migration...")
            button_Start.PerformClick()
        End If
    End Sub

    Private Function func_InitCheckForValidConfiguration() As Boolean

        Sub_DebugMessage()
        Sub_DebugMessage("* Check for Valid WMA Configuration *")

        If _
            My.Settings.MigrationNetworkLocation = StrWmaConfigNetworkCheck And
            Not My.Settings.MigrationNetworkLocationDisabled Then
            Sub_DebugMessage("WARNING: Configuration does not seem to be valid")
            Return False
        Else
            Sub_DebugMessage("Configuration seems to be valid")
            Return True
        End If
    End Function

    Private Function func_MigrationSupportScripts(migrationTimeLine As String) As Boolean
        ArraylistScriptsCurrent.Clear()

        Select Case StrMigrationType
            Case "SCANSTATE"
                Select Case migrationTimeLine.ToUpper
                    Case "PRE"
                        For Each script As String In ArrayMigrationScriptsPreCapture
                            If Not script.Trim = Nothing Then
                                ArraylistScriptsCurrent.Add(script)
                            End If
                        Next
                    Case "POST"
                        For Each script As String In ArrayMigrationScriptsPostCapture
                            If Not script.Trim = Nothing Then
                                ArraylistScriptsCurrent.Add(script)
                            End If
                        Next
                End Select
            Case "LOADSTATE"
                Select Case migrationTimeLine.ToUpper
                    Case "PRE"
                        For Each script As String In ArrayMigrationScriptsPreRestore
                            If Not script.Trim = Nothing Then
                                ArraylistScriptsCurrent.Add(script)
                            End If
                        Next
                    Case "POST"
                        For Each script As String In ArrayMigrationScriptsPostRestore
                            If Not script.Trim = Nothing Then
                                ArraylistScriptsCurrent.Add(script)
                            End If
                        Next
                End Select
        End Select

        ' If no scripts to run, then exit the Function
        If ArraylistScriptsCurrent.Count = 0 Then
            Sub_DebugMessage("No scripts to run...")
            Return True
        End If

        ' Otherwise, process each script and return if one of them fails
        For Each script As String In ArraylistScriptsCurrent
            Sub_DebugMessage($"{migrationTimeLine}-Migration Script: {script.Remove(script.Length - 4)}...")
            label_MigrationCurrentPhase.Text =
                $"{migrationTimeLine}-Migration Script: {script.Remove(script.Length - 4)}{Elipses}"
            Application.DoEvents()
            Dim process As New Process
            Dim processInfo As New ProcessStartInfo()
            'Dim strTempFileCheck As String = Nothing
            Dim strMigrationScriptArguments As String =
                    "/USER " & StrEnvUserName & " " &
                    "/COMPUTER " & StrEnvComputerName & " " &
                    "/MIGFOLDER " & StrMigrationFolder & " "
            Dim strTempFileCheck = script.Substring(script.Length - 3)
            Select Case strTempFileCheck.ToUpper
                Case "VBS", "VBE"
                    processInfo.FileName = "CScript.Exe"
                    processInfo.Arguments = """" & script & """ " & strMigrationScriptArguments
                    processInfo.CreateNoWindow = BlnMigrationScriptsNoWindow
                Case "PS1"
                    processInfo.FileName = "PowerShell.Exe"
                    processInfo.Arguments = "-ExecutionPolicy Bypass -File """ & script & """ " &
                                            strMigrationScriptArguments
                    processInfo.CreateNoWindow = BlnMigrationScriptsNoWindow
                Case Else
                    processInfo.FileName = script
                    processInfo.Arguments = strMigrationScriptArguments
                    processInfo.CreateNoWindow = BlnMigrationScriptsNoWindow
            End Select
            ' Configure the Migration process
            processInfo.WorkingDirectory = StrWmaFolder
            processInfo.UseShellExecute = False

            ' Log debugging info
            Sub_DebugMessage("-- " & processInfo.FileName & " " & processInfo.Arguments)

            ' Start the Migration Process
            Try
                process.StartInfo = processInfo
                process.Start()
                process.WaitForExit()
            Catch ex As Exception
                Sub_DebugMessage("ERROR: " & ex.Message, True)
            End Try

            Select Case process.ExitCode
                Case 0
                    Return True
                Case Else
                    Return False
            End Select
        Next

        Return False
    End Function

    ' Use WMI to get the Drive letter based on the passed in WMI Name
    Private Shared Function Func_USBGetDriveLetter(name As String) As String

        Dim objqueryPartition, objqueryDisk As ObjectQuery
        Dim searcherPartition, searcherDisk As ManagementObjectSearcher
        Dim objPartition, objDisk As ManagementObject
        Dim strReturn = ""

        ' WMI queries use the "\" as an escape charcter
        name = Replace(name, "\", "\\")

        ' First we map the Win32_DiskDrive instance with the association called
        ' Win32_DiskDriveToDiskPartition.  Then we map the Win23_DiskPartion
        ' instance with the assocation called Win32_LogicalDiskToPartition

        Try

            objqueryPartition =
                New ObjectQuery(
                    "ASSOCIATORS OF {Win32_DiskDrive.DeviceID=""" & name &
                    """} WHERE AssocClass = Win32_DiskDriveToDiskPartition")
            searcherPartition = New ManagementObjectSearcher(objqueryPartition)
            For Each objPartition In searcherPartition.Get()

                objqueryDisk =
                    New ObjectQuery(
                        "ASSOCIATORS OF {Win32_DiskPartition.DeviceID=""" & CStr(objPartition("DeviceID")) &
                        """} WHERE AssocClass = Win32_LogicalDiskToPartition")
                searcherDisk = New ManagementObjectSearcher(objqueryDisk)
                For Each objDisk In searcherDisk.Get()
                    strReturn &= CStr(objDisk("Name")) & ","
                Next
            Next

        Catch ex As Exception
            Sub_DebugMessage("ERROR: Failed to determine USB drive letter. " & ex.Message)
            Return Nothing
        End Try

        Return strReturn.Trim(","c)
    End Function

    ' Data Capture - Workstation Details Radioboxes
    Private Sub radiobox_WorkstationDetails_CheckedChanged(sender As Object, e As EventArgs) _
        Handles radiobox_WorkstationDetails1.CheckedChanged, radiobox_WorkstationDetails2.CheckedChanged

        Sub_DebugMessage()
        Sub_DebugMessage("* Workstation Details Radiobox Change Events *")

        If radiobox_WorkstationDetails2.Checked Then
            Sub_DebugMessage("Switched to migrate All Users")
            BlnMigrationSettingsAllUsers = True
        Else
            Sub_DebugMessage("Switched to migrate Current User")
            BlnMigrationSettingsAllUsers = False
        End If
    End Sub

    Private Sub sub_HealthCheckFinish() Handles _classHealthCheck.HealthCheckFinished

        ' Handle thread invoking so that the UI can be updated cross-thread
        If InvokeRequired Then
            Invoke(New MethodInvoker(AddressOf sub_HealthCheckFinish))
        Else

            Sub_DebugMessage()
            Sub_DebugMessage("* Finish Health Check *")

            Try

                ' Max out the progress bar
                progressbar_Migration.Value = 100

                Sub_DebugMessage("Spinning down Health Check Class...")
                _classHealthCheck.SpinDown()

                ' Check the exit code. Let's not take chances, if it's not 0, then it's an error.
                Select Case _classHealthCheck.ExitCode
                    Case 0
                        Sub_DebugMessage("Success, no errors found")
                        label_MigrationCurrentPhase.Text = $"Health Check: {diskScanResultOk}"
                        BlnHealthCheckStatusOk = True
                    Case 1
                        Sub_DebugMessage("Success, CHKDSK has detected and fixed major errors")
                        label_MigrationCurrentPhase.Text = $"Health Check: {diskScanResultOk}"
                        BlnHealthCheckStatusOk = True
                    Case 2
                        Sub_DebugMessage("Success, CHKDSK has detected and fixed minor inconsistencies")
                        label_MigrationCurrentPhase.Text = $"Health Check: {diskScanResultOk}"
                        BlnHealthCheckStatusOk = True
                    Case 999
                        Sub_DebugMessage("Health Check Cancelled")
                        BlnMigrationCancelled = True
                    Case Else
                        Sub_DebugMessage("ERROR: Health Check Failed")
                        label_MigrationCurrentPhase.Text = $"Health Check: {diskScanResultErrors}"

                        ' If in Progress Only Mode, exit the app
                        If BlnAppProgressOnlyMode Then
                            AppShutdown(5)
                        End If

                        ' See if the user wants to fix the problem or continue
                        Sub_DebugMessage("Prompting User for action...")
                        Dim msgboxResult As DialogResult = MessageBox.Show(diskScanFixMessage,
                                                                           appTitle,
                                                                           MessageBoxButtons.YesNoCancel,
                                                                           MessageBoxIcon.Exclamation)
                        ' If yes, start CHKDSK
                        If msgboxResult = DialogResult.Yes Then
                            Sub_DebugMessage("INFO: User chose to fix disk and reboot")
                            Process.Start("CMD", "/C Echo Y | CHKDSK /F /R")
                            Process.Start("Shutdown",
                                          "-R -T 10 -C ""CHKDSK Initialised by " & appTitle &
                                          ". Restarting Workstation""")
                            AppShutdown(1)
                            ' If no, continue with the migration
                        ElseIf msgboxResult = DialogResult.No Then
                            Sub_DebugMessage("WARNING: " & diskScanFixCancelledMessage, True)
                            ' If cancel, close the application
                        Else
                            Sub_DebugMessage("INFO: User chose to exit application")
                            AppShutdown(8)
                        End If
                End Select

            Catch ex As Exception
                BlnMigrationStatusOk = False
                label_MigrationCurrentPhase.Text = ex.Message
                Sub_DebugMessage(ex.Message, True)
            End Try

            BlnHealthCheckInProgress = False

            ' Start Migration
            If Not BlnMigrationCancelled Then
                sub_MigrationSetup()
            End If

            If Not BlnMigrationCancelled Then
                sub_MigrationStart()
            End If

            If BlnMigrationCancelled Then
                sub_SupportFormControlsReset()
            End If

        End If
    End Sub

    Private Sub Sub_HealthCheckProgressMonitor() Handles _classHealthCheck.ProgressUpdate

        ' Handle thread invoking so that the UI can be updated cross-thread
        If InvokeRequired Then
            Invoke(New MethodInvoker(AddressOf Sub_HealthCheckProgressMonitor))
        Else

            ' Get the current status
            If Not _classHealthCheck.Progress = Nothing Then
                If _
                    ((Not (_classHealthCheck.PercentComplete = Nothing)) Or
                     (CInt((_classHealthCheck.PercentComplete = Nothing)) = 0)) _
                    Then
                    StrStatusMessage = _classHealthCheck.Progress & " (" & _classHealthCheck.PercentComplete &
                                       diskScanPercent & ")"
                Else
                    StrStatusMessage = _classHealthCheck.Progress
                End If
            End If

            ' If the status has changed, output
            If Not StrStatusMessage = StrPreviousStatusMessage Then
                Sub_DebugMessage(StrStatusMessage)
                label_MigrationCurrentPhase.Text = StrStatusMessage
                progressbar_Migration.Value = _classHealthCheck.PercentComplete
            End If

            StrPreviousStatusMessage = StrStatusMessage

        End If
    End Sub

    Private Sub Sub_HealthCheckStart()

        Sub_DebugMessage()
        Sub_DebugMessage("* Start Health Check *")

        ' Initialise the Health Check Class
        _classHealthCheck = New ClassHealthCheck

        ' Reset Status items
        'Dim strPreviousStatusMessage As String = Nothing
        'Dim strStatusMessage As String = Nothing
        progressbar_Migration.Value = 0

        ' Set the Health Check to being in progress
        BlnHealthCheckInProgress = True

        ' Start the Health Check
        Try
            Sub_DebugMessage("Spinning up Health Check class...")
            _classHealthCheck.Spinup()

        Catch ex As Exception
            Sub_DebugMessage("ERROR: " & ex.Message, True)
            label_MigrationCurrentPhase.Text = ex.Message
            _classHealthCheck.SpinDown()
            Exit Sub
        End Try
    End Sub

    Private Sub sub_InitCheckForRuleSetFiles()

        Sub_DebugMessage()
        Sub_DebugMessage("* Check for RuleSet Files *")

        Dim blnRuleSetMissing = False
        Dim blnRuleSetCopyFailed = False

        ' *** Set up the rules for migration
        For Each ruleset As String In ArrayMigrationRuleSet
            If Not My.Computer.FileSystem.FileExists(Path.Combine(StrWmaFolder, ruleset.Trim)) Then
                Sub_DebugMessage(
                    "WARNING: RuleSet file " & ruleset.Trim &
                    " does not exist. Copying default file from USMT folder...")
                Try
                    My.Computer.FileSystem.CopyFile(Path.Combine(StrUsmtFolder, ruleset.Trim),
                                                    Path.Combine(StrWmaFolder, ruleset.Trim), True)
                Catch ex As Exception
                    Sub_DebugMessage("ERROR: Unable to copy file - " & Path.Combine(StrUsmtFolder, ruleset.Trim))
                    blnRuleSetCopyFailed = True
                End Try
                blnRuleSetMissing = True
            End If
        Next

        If blnRuleSetMissing Then
            If blnRuleSetCopyFailed Then
                Sub_DebugMessage(
                    "ERROR: RuleSet files were not found in the WMA folder. An attempt to copy the default RuleSets from USMT failed. The application will now close.",
                    True)
                AppShutdown(21)
            End If
            Sub_DebugMessage(
                "INFO: RuleSet files were not found in the WMA folder. This may be your first time running WMA, so the default RuleSet files have been copied from USMT",
                True)
        Else
            Sub_DebugMessage("All RuleSet files exist")
        End If
    End Sub

    Private Sub Sub_InitCheckForUSMT()

        Sub_DebugMessage()
        Sub_DebugMessage("* Check for USMT *")

        'Dim blnIsUsmtInstalled = True
        'Dim strBddManifestFileFullPath As String = Path.Combine(StrTempFolder, StrBddManifestFile)
        'Dim strBddComponentFileFullPath As String = Path.Combine(StrTempFolder, My.Resources.bddComponentFile)
        Const strUsmtSearchFile = "ScanState.Exe"
        'Dim xmlGuid As String = Nothing
        'Dim xmlArchitecture As String = Nothing
        'Dim xmlInstallCondition As String = Nothing
        'Dim xmlDownloadUrl As String = Nothing
        'Dim xmlDownloadFile As String = Nothing
        'Dim xmlDescription As String = Nothing

        ' Check for USMT in the relevant folders
        Sub_DebugMessage("Checking for USMT file: " & strUsmtSearchFile & "...")
        Try
            If Not My.Computer.FileSystem.FileExists(Path.Combine(StrWmaFolder, strUsmtSearchFile)) Then
                ' If not found, search subdirectories by architecture for any matches
                For Each foundDirectory As String In
                    (My.Computer.FileSystem.GetFiles(StrWmaFolder, SearchOption.SearchAllSubDirectories,
                                                     strUsmtSearchFile))
                    ' Use the same architecture if required
                    If (Path.GetDirectoryName(foundDirectory).ToUpper.Contains(StrOsArchitecture.ToString.ToUpper)) _
                        Then
                        StrUsmtFolder = Path.GetDirectoryName(foundDirectory)
                        'blnIsUsmtInstalled = True
                        Exit For
                    End If
                Next
                ' If not found, search the Program Files folder for any matches
                For Each foundDirectory As String In
                    My.Computer.FileSystem.GetDirectories(My.Computer.FileSystem.SpecialDirectories.ProgramFiles,
                                                          SearchOption.SearchTopLevelOnly, "USMT*")
                    StrUsmtFolder = foundDirectory
                    'blnIsUsmtInstalled = True
                Next
                ' If still no USMT found, exit the application with an error
                If Not My.Computer.FileSystem.FileExists(Path.Combine(StrUsmtFolder, strUsmtSearchFile)) Then
                    Throw New Exception(Path.Combine(StrUsmtFolder, strUsmtSearchFile) & " does not exist")
                End If
            Else
                StrUsmtFolder = StrWmaFolder
                'blnIsUsmtInstalled = True
            End If
            Sub_DebugMessage("Success. Found at: " & Path.Combine(StrUsmtFolder, strUsmtSearchFile))
        Catch ex As Exception
            'blnIsUsmtInstalled = False
            Sub_DebugMessage("WARNING: Unable to find USMT installation: " & ex.Message & "")
        End Try

        label_MigrationCurrentPhase.Text = Elipses
        button_Start.Enabled = True
    End Sub

    'End Sub
    Private Sub sub_InitParseCommandLine()

        Sub_DebugMessage()
        Sub_DebugMessage("* Parse Command Line *")

        Try
            Dim colCommandLineArguments As ReadOnlyCollection(Of String) = My.Application.CommandLineArgs
            Dim intCount As Integer
            If Not colCommandLineArguments.Count > 0 Then
                Exit Sub
            End If
            For intCount = 0 To colCommandLineArguments.Count - 1
                Sub_DebugMessage("Checking: " & colCommandLineArguments(intCount).ToUpper)
                Select Case colCommandLineArguments(intCount).ToUpper
                    Case "/PROGRESSONLY"
                        Sub_DebugMessage("Match: Progress Only Mode")
                        BlnAppProgressOnlyMode = True
                        If (intCount + 1) < colCommandLineArguments.Count Then
                            Select Case colCommandLineArguments(intCount + 1).ToUpper
                                Case "CAPTURE"
                                    Sub_DebugMessage("Type: Capture")
                                    tabcontrol_MigrationType.SelectedTab = tabpage_Capture
                                Case "RESTORE"
                                    Sub_DebugMessage("Type: Restore")
                                    tabcontrol_MigrationType.SelectedTab = tabpage_Restore
                                Case Else
                                    Throw New Exception("/PROGRESSONLY was specified without CAPTURE or RESTORE")
                            End Select
                        Else
                            Throw New Exception("/PROGRESSONLY was specified without CAPTURE or RESTORE")
                        End If
                    Case "/MIGOVERWRITEEXISTING"
                        Sub_DebugMessage("Match: Overwrite Existing Migration Folders")
                        BlnMigrationOverwriteExistingFolders = True
                    Case "/MIGMAXOVERRIDE"
                        Sub_DebugMessage("Match: Override Maximum Migration Size Limit")
                        BlnMigrationMaxOverride = True
                    Case "/MIGFOLDER"
                        Sub_DebugMessage("Match: Alternate Migration Folder")
                        If (intCount + 1) < colCommandLineArguments.Count Then
                            StrMigrationLocationOther = colCommandLineArguments(intCount + 1).ToUpper
                            BlnMigrationLocationUseOther = True
                            Sub_DebugMessage("Folder: " & StrMigrationFolder)
                        Else
                            Throw New Exception("/MIGFOLDER was specified without a Migration Folder")
                        End If
                    Case "/CHANGEDOMAIN"
                        Sub_DebugMessage("Match: Change Domain Mode")
                        If (intCount + 1) < colCommandLineArguments.Count Then
                            StrMigrationDomainChange = colCommandLineArguments(intCount + 1).ToUpper
                            Sub_DebugMessage("New Domain Name: " & StrMigrationDomainChange)
                        Else
                            Throw New Exception("/CHANGEDOMAIN was specified without a New Domain Name")
                        End If
                    Case "/PRIMARYDATADRIVE"
                        Sub_DebugMessage("Match: Change Primary Data Drive Mode")
                        If (intCount + 1) < colCommandLineArguments.Count Then
                            StrPrimaryDataDrive = colCommandLineArguments(intCount + 1).ToUpper
                            Sub_DebugMessage("New Primary Data Drive: " & StrPrimaryDataDrive)
                        Else
                            Throw New Exception("/PRIMARYDATADRIVE was specified without a Primary Data Drive")
                        End If
                    Case "/MULTIUSER"
                        Sub_DebugMessage("Match: Multi-User Mode")
                        ' Enable Multi-User Mode
                        radiobox_WorkstationDetails2.Checked = True
                End Select
            Next
        Catch ex As Exception
            Sub_DebugMessage(
                "ERROR: A problem occurred while processing command-line parameters: " & vbNewLine & ex.Message, True)
            AppShutdown(CInt(exitCodeCMDLineParamError))
        End Try
    End Sub

    Private Shared Sub Sub_MigrationCreateFolderStructure(folderStructure As String)
        Dim arrayFolder() As String
        Dim strFolderBuild As String

        If Not My.Computer.FileSystem.DirectoryExists(folderStructure) Then
            ' Replace UNC path so the array can be built correctly
            folderStructure = Replace(folderStructure, "\\", "//")
            arrayFolder = Split(folderStructure, "\")
            'Dim i As Integer
            For i = 0 To UBound(arrayFolder)
                ' This if/else is commented because it's unreachable, see below td note
                'If strFolderBuild = Nothing Then
                '    strFolderBuild = arrayFolder(i)
                'Else
                '    strFolderBuild = Path.Combine(strFolderBuild, arrayFolder(i))
                'End If
                'TODO: Is this a bug? The if/else above doesn't matter because it's overriden in the following statement
                ' Reinstate UNC path
                strFolderBuild = Replace(folderStructure, "//", "\\")
                Try
                    If Not My.Computer.FileSystem.DirectoryExists(strFolderBuild) Then
                        My.Computer.FileSystem.CreateDirectory(strFolderBuild)
                    End If
                Catch exIo As IOException
                    Sub_DebugMessage($"ERROR: {exIo.Message}", True)
                Catch ex As Exception
                    Sub_DebugMessage($"ERROR: {ex.Message}", True)
                End Try
            Next
        End If
    End Sub

    Private Sub Sub_MigrationFindDataStore()

        Dim strTempMigrationFolder As String
        'Dim arrayTempFolder() As String = Nothing

        Sub_DebugMessage("* Find Migration Data Store *")

        ' Check that WMA is configured correctly
        If Not func_InitCheckForValidConfiguration() Then
            Sub_DebugMessage(
                "ERROR: It looks like you haven't yet modified the WMA configuration file (MigAssistant.Exe.Config). As a result, it is not possible to perform a restore!",
                True)
            tabcontrol_MigrationType.SelectedTab = tabpage_Capture
            Exit Sub
        End If

        ' *** Check where the Restore is being performed from...
        Try

            If BlnMigrationLocationUseOther Then
                Sub_DebugMessage("Migrating from Alternate location: " & StrMigrationLocationOther)
                StrMigrationFolder = StrMigrationLocationOther
            ElseIf BlnMigrationLocationUseUsb Then
                Sub_DebugMessage("Migrating from USB drive: " & StrMigrationLocationUsb)
                StrMigrationFolder = StrMigrationLocationUsb
            ElseIf StrMigrationLocationNetwork.Contains("\\") Then
                Sub_DebugMessage("Migrating from Network Location: " & StrMigrationLocationNetwork)
                Dim strMigrationServer As String = StrMigrationLocationNetwork.Replace("\\", "")
                strMigrationServer = strMigrationServer.Remove(strMigrationServer.IndexOf("\", StringComparison.Ordinal))
                ' If no network connection available, or unable to ping server
                If Not My.Computer.Network.IsAvailable Or Not My.Computer.Network.Ping(strMigrationServer) Then
                    Sub_DebugMessage("WARNING: " & migrationFindServerFailText, True)
                Else
                    StrMigrationFolder = StrMigrationLocationNetwork
                End If
            Else
                StrMigrationFolder = StrMigrationLocationNetwork
            End If

        Catch ex As Exception

        End Try

        Try
            ' Check for username in all migration folders
            Sub_DebugMessage("Checking for username in all migration folders...")
            Dim colMigrationFolders As ReadOnlyCollection(Of String) =
                    My.Computer.FileSystem.GetDirectories(StrMigrationFolder)
            Dim arraylistFolders As New ArrayList
            arraylistFolders.Clear()
            For Each strFolderSearch As String In colMigrationFolders
                'folder = Replace(folder, str_MigrationFolder & "\", "")
                strFolderSearch = Replace(strFolderSearch, StrMigrationFolder & "\", "", 1, -1, CompareMethod.Text)
                If CBool(InStr(strFolderSearch, StrEnvUserName, CompareMethod.Text)) Then
                    arraylistFolders.Add(strFolderSearch)
                    Sub_DebugMessage("Match Found: " & strFolderSearch)
                End If
            Next
            Select Case arraylistFolders.Count
                ' If no match
                Case 0
                    Throw New Exception(datastoreNotFound)
                    ' If one match
                Case 1
                    strTempMigrationFolder = arraylistFolders(0).ToString()
                    ' If multiple matches
                Case Else
                    ' Present option to select the correct datastore
                    FormRestoreMultiDatastore.cbxRestoreMultiDatastoreList.Items.Clear()
                    For Each folderTmp As String In arraylistFolders
                        FormRestoreMultiDatastore.cbxRestoreMultiDatastoreList.Items.Add(folderTmp)
                    Next
                    FormRestoreMultiDatastore.cbxRestoreMultiDatastoreList.SelectedItem =
                        FormRestoreMultiDatastore.cbxRestoreMultiDatastoreList.Items(0)
                    FormRestoreMultiDatastore.ShowDialog()
                    strTempMigrationFolder =
                        FormRestoreMultiDatastore.cbxRestoreMultiDatastoreList.SelectedItem.ToString()
                    If strTempMigrationFolder = Nothing Then
                        Throw New Exception(datastoreMultipleFoundNoneSelected)
                    End If
            End Select
        Catch ex As Exception
            Sub_DebugMessage("ERROR: " & datastoreDetectionError & ": " & ex.Message, True)
            Exit Sub
        End Try

        StrMigrationFolder = Path.Combine(StrMigrationFolder, strTempMigrationFolder).Trim

        Sub_DebugMessage("Migration DataStore Full Path: " & StrMigrationFolder)

        label_DatastoreLocation.Text = StrMigrationFolder

        Try
            Dim xmlDocument = New XmlDocument
            xmlDocument.Load(Path.Combine(StrMigrationFolder, "Logging\WMA_ScanState.XML"))
            Dim xmlNodes As XmlNodeList = xmlDocument.GetElementsByTagName("MigAssistant")
            'xmlDocument = Nothing
            For Each xmlNode As XmlNode In xmlNodes
                BlnMigrationSettingsAllUsers = CBool(xmlNode.Item("Options").Attributes.GetNamedItem("AllUsers").Value)
                BlnMigrationCompressionDisabled =
                    CBool(xmlNode.Item("Options").Attributes.GetNamedItem("CompressionDisabled").Value)
                BlnMigrationEncryptionDisabled =
                    CBool(xmlNode.Item("Options").Attributes.GetNamedItem("EncryptionDisabled").Value)
                If Not BlnMigrationEncryptionDisabled Then
                    BlnMigrationEncryptionCustom =
                        CBool(xmlNode.Item("Options").Attributes.GetNamedItem("EncryptionCustom").Value)
                    If BlnMigrationEncryptionCustom Then
                        FormCustomEncryption.ShowDialog()
                    End If
                End If
                Dim intMigrationDataSize =
                        CInt(xmlNode.Item("SCANSTATE").Attributes.GetNamedItem("DataSize").Value)
                If StrPrimaryDataDrive = Nothing Then StrPrimaryDataDrive = "C:"
                If intMigrationDataSize > Func_GetFreeSpace(StrPrimaryDataDrive) Then
                    Throw New Exception("There is not enough free space on this drive to perform the migration")
                End If
                Exit For
            Next
        Catch ex As Exception
            Sub_DebugMessage("ERROR: While parsing WMA_SCANSTATE.XML: " & ex.Message, True)
            AppShutdown(30)
        End Try
    End Sub

    Private Sub Sub_MigrationFinish() Handles _classMigration.MigrationFinished

        ' Handle thread invoking so that the UI can be updated cross-thread
        If InvokeRequired Then
            Invoke(New MethodInvoker(AddressOf Sub_MigrationFinish))
        Else

            Sub_DebugMessage()
            Sub_DebugMessage("* Finish Migration *")

            Try

                Sub_DebugMessage("Spinning down Migration Class...")
                _classMigration.SpinDown()

                Sub_DebugMessage("Checking Exit Code: " & _classMigration.ExitCode)
                ' Check the exit code. 
                Select Case _classMigration.ExitCode
                    Case 0, 1073741819, -1073741819
                        BlnMigrationStatusOk = True
                        ' Check if an abort has occurred and end the migration
                    Case 12
                        Throw _
                            New Exception(
                                "You must be an administrator to migrate one or more of the files or settings that are in the store. Log on as an administrator and try again.")
                    Case 999
                        BlnMigrationCancelled = True
                        Throw New Exception("Migration Cancelled")
                    Case Else
                        Throw New Exception("An unknown error occurred. Check the USMT log files for more details")
                End Select

            Catch ex As Exception
                BlnMigrationStatusOk = False
                label_MigrationCurrentPhase.Text = ex.Message
                Sub_DebugMessage(ex.Message, True)
            End Try

            ' *** Run Post-Migration Scripts
            If BlnMigrationStatusOk Then
                Sub_DebugMessage("Running Post-Migration Scripts...")
                If Not func_MigrationSupportScripts("Post") Then
                    label_MigrationCurrentPhase.ForeColor = Color.Red
                    label_MigrationCurrentPhase.Text = migrationPostMigrationScriptFail
                    If BlnAppProgressOnlyMode Then
                        AppShutdown(CInt(exitCodePostMigrationScriptFail))
                    End If
                    Exit Sub
                End If
            End If

            ' Write WMA XML Log File
            Try
                Sub_DebugMessage("Building XML Log File...")
                Dim xmlLogFile =
                        New XmlTextWriter(
                            Path.Combine(StrMigrationFolder, StrMigrationLoggingFolder) & "\WMA_" & StrMigrationType &
                            ".XML", Nothing)
                xmlLogFile.WriteStartDocument()
                xmlLogFile.WriteComment(appTitle & " " & appBuild)
                xmlLogFile.WriteStartElement("MigAssistant")
                xmlLogFile.WriteStartElement("Computer")
                xmlLogFile.WriteAttributeString("ComputerName", StrEnvComputerName)
                xmlLogFile.WriteAttributeString("User", StrEnvUserName)
                xmlLogFile.WriteAttributeString("OSName", StrOsFullName)
                xmlLogFile.WriteEndElement()
                xmlLogFile.WriteStartElement("Options")
                Select Case StrMigrationType.ToUpper
                    Case "SCANSTATE"
                        xmlLogFile.WriteAttributeString("HealthCheck", CStr(BlnHealthCheck))
                        If BlnHealthCheck Then
                            xmlLogFile.WriteAttributeString("HealthCheckStatusOk", CStr(BlnHealthCheckStatusOk))
                        End If
                        xmlLogFile.WriteAttributeString("AllUsers", CStr(BlnMigrationSettingsAllUsers))
                        xmlLogFile.WriteAttributeString("CompressionDisabled", CStr(BlnMigrationCompressionDisabled))
                End Select
                xmlLogFile.WriteAttributeString("EncryptionDisabled", CStr(BlnMigrationEncryptionDisabled))
                If Not BlnMigrationEncryptionDisabled Then
                    xmlLogFile.WriteAttributeString("EncryptionCustom", CStr(BlnMigrationEncryptionCustom))
                End If
                xmlLogFile.WriteEndElement()
                xmlLogFile.WriteStartElement(StrMigrationType)
                Select Case StrMigrationType.ToUpper
                    Case "SCANSTATE"
                        xmlLogFile.WriteAttributeString("DataSize", CStr(_classMigration.EstDataSize))
                End Select
                xmlLogFile.WriteAttributeString("TimeStart", DtmStartTime.ToString())
                xmlLogFile.WriteAttributeString("TimeEnd", DateTime.Now.ToString())
                xmlLogFile.WriteAttributeString("StatusOk", BlnMigrationStatusOk.ToString())
                xmlLogFile.WriteAttributeString("ExitCode", _classMigration.ExitCode.ToString())
                xmlLogFile.WriteEndElement()
                xmlLogFile.Close()
            Catch ex As Exception
                Sub_DebugMessage("ERROR: Failed to build XML log file: " & ex.Message, True)
            End Try

            ' Send Email
            If BlnMailSend And Not BlnMigrationCancelled Then
                Sub_DebugMessage("Emailing results...")
                ' Check for Network connectivity
                Sub_DebugMessage("Checking for network connection...")
                If Not My.Computer.Network.IsAvailable Then
                    Sub_DebugMessage("No internet connection available. Skipping...")
                Else

                    Try
                        Dim email As New ClassEmail
                        email.Server = StrMailServer
                        email.Recipients = StrMailRecipients
                        email.From = StrMailFrom

                        If BlnMigrationStatusOk Then
                            email.Subject =
                                $"MigAssistant Success - {StrMigrationType}: {StrEnvUserName} ({StrEnvComputerName})"
                        Else
                            email.Subject =
                                $"MigAsssitant Failure - {StrMigrationType}: {StrEnvUserName} ({StrEnvComputerName})"
                        End If

                        email.Message = mailMessage

                        ' Add attachments
                        Dim attachmentArray As New ArrayList
                        ' Add WMA XML Logfiles if found
                        If _
                            My.Computer.FileSystem.FileExists(
                                Path.Combine(StrMigrationFolder, StrMigrationLoggingFolder) & "\WMA_" &
                                StrMigrationType & ".XML") Then
                            attachmentArray.Add(Path.Combine(StrMigrationFolder,
                                                             StrMigrationLoggingFolder & "\WMA_" & StrMigrationType &
                                                             ".XML"))
                        End If
                        ' Add Migration Logfiles if found
                        If _
                            My.Computer.FileSystem.FileExists(
                                Path.Combine(StrMigrationFolder, StrMigrationLoggingFolder) & "\" & StrMigrationType &
                                ".Log") Then
                            attachmentArray.Add(Path.Combine(StrMigrationFolder,
                                                             StrMigrationLoggingFolder & "\" & StrMigrationType &
                                                             ".Log"))
                        End If
                        ' Add Migration Progress Logfiles if found
                        If _
                            My.Computer.FileSystem.FileExists(
                                Path.Combine(StrMigrationFolder, StrMigrationLoggingFolder) & "_Progress.Log") Then
                            attachmentArray.Add(Path.Combine(StrMigrationFolder,
                                                             StrMigrationLoggingFolder & "_Progress.Log"))
                        End If
                        ' Add the debug Logfile if found (by generating a copy)
                        If My.Computer.FileSystem.FileExists(StrLogFile) Then
                            My.Computer.FileSystem.CopyFile(StrLogFile, $"{StrLogFile}_.Log", True)
                            attachmentArray.Add($"{StrLogFile.Trim}_.Log")
                        End If
                        If Not attachmentArray.Count = 0 Then
                            email.Attachments = Join(attachmentArray.ToArray, ",")
                        End If

                        ' Update Migration status text
                        label_MigrationCurrentPhase.Text = StatusLabelEmailingResults
                        Application.DoEvents()

                        ' Send Email
                        email.Send()
                        Sub_DebugMessage("Email Sent")
                    Catch ex As Exception
                        Sub_DebugMessage("ERROR: Email Send failed: " & ex.Message, True)
                    End Try

                End If

            End If

            ' Update Migration status text
            If BlnMigrationStatusOk Then
                label_MigrationCurrentPhase.ForeColor = Color.Green
                label_MigrationCurrentPhase.Text = migrationSuccessStatus
                If BlnAppProgressOnlyMode Then
                    AppShutdown(CInt(exitCodeOk))
                End If
            Else
                label_MigrationCurrentPhase.ForeColor = Color.Red
                label_MigrationCurrentPhase.Text = migrationFailedStatus
                Sub_DebugMessage(migrationFailedMessage)
                If BlnAppProgressOnlyMode Then
                    AppShutdown(CInt(exitCodeMigrationFailed))
                End If
            End If

            BlnMigrationInProgress = False

            sub_SupportFormControlsReset()

        End If
    End Sub

    Private Sub sub_MigrationInitialise()

        Sub_DebugMessage()
        Sub_DebugMessage("* Migration Initialise *")

        BlnMigrationCancelled = False

        ' Initialise the Migration Class
        _classMigration = New ClassMigration

        ' Disable / Update form controls
        button_AdvancedSettings.Enabled = False
        tabcontrol_MigrationType.Enabled = False
        button_Start.Text = LabelStopButton

        ' Enable / Update status items
        group_Status.Enabled = True
        progressbar_Migration.Value = 0
        progressbar_Migration.Visible = True
        label_MigrationCurrentPhase.ForeColor = Color.Black
        label_MigrationCurrentPhase.Text = StatusLabelInitializing

        ' Start the Health Check
        If BlnHealthCheck Then
            ' Skip if LoadState - Only needed for capture, not restore.
            If Not StrMigrationType = "LOADSTATE" Then
                Sub_HealthCheckStart()
            Else
                ' Start Migration
                If Not BlnMigrationCancelled Then
                    sub_MigrationSetup()
                End If

                If Not BlnMigrationCancelled Then
                    sub_MigrationStart()
                End If
            End If
        Else
            ' Start Migration
            If Not BlnMigrationCancelled Then
                sub_MigrationSetup()
            End If

            If Not BlnMigrationCancelled Then
                sub_MigrationStart()
            End If
        End If
    End Sub

    Private Sub sub_MigrationProgressMonitor() Handles _classMigration.ProgressUpdate

        Try

            ' Handle thread invoking so that the UI can be updated cross-thread
            If InvokeRequired Then
                Invoke(New MethodInvoker(AddressOf sub_MigrationProgressMonitor))
            Else

                ' Update form from class information
                If Not _classMigration.EstTimeRemaining = 0 Then
                    label_MigrationEstTimeRemaining.Text =
                        $"{_classMigration.EstTimeRemaining} {migrationTotalMinutesRemaining}"
                End If
                If Not _classMigration.EstDataSize = 0 Then
                    label_MigrationEstSize.Text = _classMigration.EstDataSize &
                                                  migrationTotalSizeInMBToTransfer
                    ' Check if a ScanState is being performed
                    If StrMigrationType = "SCANSTATE" Then
                        If Not BlnSizeChecksDone Then
                            ' If not overridden via commandline, check the migration size is not exceeded
                            If (IntMigrationMaxSize > 0 And Not BlnMigrationMaxOverride) Then
                                If _classMigration.EstDataSize > IntMigrationMaxSize Then
                                    Throw _
                                        New Exception(
                                            "ERROR: The amount of data to be migrated exceeds the maximum allowed size. Please remove any unnecessary data and try again." &
                                            vbNewLine & vbNewLine & "Current: " & _classMigration.EstDataSize & "MB" &
                                            vbNewLine & "Maximum: " & IntMigrationMaxSize & "MB")
                                End If
                            End If
                            If _classMigration.EstDataSize > Func_GetFreeSpace(StrMigrationFolder) Then
                                Throw _
                                    New Exception(
                                        "ERROR: There is not enough space available to perform the migration. Estimated Migration Size: " &
                                        _classMigration.EstDataSize & "MB. Available Space: " &
                                        Func_GetFreeSpace(StrMigrationFolder) & "MB")
                            End If
                            BlnSizeChecksDone = True
                        End If
                    End If
                End If
                ' Get current status
                If Not _classMigration.PercentComplete = 0 Then
                    StrStatusMessage = _classMigration.Progress & " (" & _classMigration.PercentComplete &
                                       migrationTotalPercentageCompleted & ")"
                Else
                    StrStatusMessage = _classMigration.Progress
                End If

                ' If status has changed, Output
                If Not StrStatusMessage = StrPreviousStatusMessage Then
                    Sub_DebugMessage(StrStatusMessage)
                    label_MigrationCurrentPhase.Text = StrStatusMessage
                    progressbar_Migration.Value = _classMigration.PercentComplete
                End If

                StrPreviousStatusMessage = StrStatusMessage

            End If

        Catch ex As Exception
            Sub_DebugMessage(ex.Message, True)
            label_MigrationCurrentPhase.Text = ex.Message
            _classMigration.SpinDown()
            Exit Sub
        End Try
    End Sub

    Private Sub sub_MigrationSetup()

        Sub_DebugMessage()
        Sub_DebugMessage("* Migration Setup *")

        ' Reset all migration settings in the array
        ArraylistMigrationArguments.Clear()

        ' *** Set up the rules for migration
        For Each ruleset As String In ArrayMigrationRuleSet
            ArraylistMigrationArguments.Add("/I:""" & Path.Combine(StrWmaFolder, ruleset.Trim) & """")
        Next
        If My.Computer.FileSystem.FileExists(Path.Combine(StrWmaFolder, StrMigrationConfigFile)) Then
            Sub_DebugMessage(
                "Config file exists and will be used: " & Path.Combine(StrWmaFolder, StrMigrationConfigFile))
            ArraylistMigrationArguments.Add("/Config:""" & Path.Combine(StrWmaFolder, StrMigrationConfigFile) & """")
        Else
            Sub_DebugMessage("Config file does not exist. Standard migration will be performed")
        End If

        Sub_DebugMessage("Migration Type: " & StrMigrationType)
        Select Case StrMigrationType
            Case "SCANSTATE"

                ' *** Check where the Backup is being performed to...
                If BlnMigrationLocationUseOther Then
                    Sub_DebugMessage("Migrating from Alternate location: " & StrMigrationLocationOther)
                    StrMigrationFolder = StrMigrationLocationOther
                ElseIf BlnMigrationLocationUseUsb Then
                    Sub_DebugMessage("Migrating from USB drive: " & StrMigrationLocationUsb)
                    StrMigrationFolder = StrMigrationLocationUsb
                ElseIf BlnMigrationLocationNetworkDisabled Then
                    Sub_DebugMessage(
                        "ERROR: Network-based migrations have been disabled by your IT Administrator. A USB drive is required to proceed with the migration.",
                        True)
                    BlnMigrationCancelled = True
                    Exit Sub
                Else
                    Sub_DebugMessage("Migrating from Network Location: " & StrMigrationLocationNetwork)
                    ' If this is a network migration, verify that pass is accessible
                    If StrMigrationLocationNetwork.EndsWith(Path.VolumeSeparatorChar) Then
                        StrMigrationLocationNetwork = StrMigrationLocationNetwork & "\"
                    End If
                    StrMigrationFolder = StrMigrationLocationNetwork
                    If Not My.Computer.FileSystem.DirectoryExists(StrMigrationFolder) Then
                        Sub_DebugMessage(
                            "ERROR: The Network-based migration location is not available. Please verify you have network connectivity. If the problem persists, contact your IT Administrator.",
                            True)
                        BlnMigrationCancelled = True
                        Exit Sub
                    End If
                End If

                ' *** Generate the folder name for the migration...
                ' *** And verify it doesn't already exist...
                If _
                    My.Computer.FileSystem.DirectoryExists(Path.Combine(StrMigrationFolder,
                                                                        $"{StrEnvComputerName}_{StrEnvUserName}")) _
                    Then
                    Try
                        If Not BlnMigrationOverwriteExistingFolders Then
                            ' Present option to remove existing migration information
                            Dim msgboxResult As DialogResult = MessageBox.Show(Me,
                                                                               migrationOverwriteExistingFolder,
                                                                               $"{StrEnvComputerName}_{StrEnvUserName}",
                                                                               MessageBoxButtons.YesNo,
                                                                               MessageBoxIcon.Question)
                            Select Case msgboxResult
                                Case DialogResult.Yes
                                    My.Computer.FileSystem.DeleteDirectory(
                                        Path.Combine(StrMigrationFolder, $"{StrEnvComputerName}_{StrEnvUserName}"),
                                        DeleteDirectoryOption.DeleteAllContents)
                            End Select
                        Else
                            My.Computer.FileSystem.DeleteDirectory(
                                Path.Combine(StrMigrationFolder, StrEnvComputerName & "_" & StrEnvUserName),
                                DeleteDirectoryOption.DeleteAllContents)
                        End If
                    Catch exPrivilege As PrivilegeNotHeldException
                        Sub_DebugMessage("ERROR: " & migrationDeleteExistingError & vbNewLine & vbNewLine &
                                         exPrivilege.Message, True)
                    Catch ex As Exception
                        Sub_DebugMessage("ERROR: " & migrationDeleteExistingError & vbNewLine & vbNewLine &
                                         ex.Message, True)
                    End Try
                End If

                StrMigrationFolder = Path.Combine(StrMigrationFolder, StrEnvComputerName & "_" & StrEnvUserName)

                ' Create the folder structure
                Sub_MigrationCreateFolderStructure(Path.Combine(StrMigrationFolder, StrMigrationDataStoreFolder))
                Sub_MigrationCreateFolderStructure(Path.Combine(StrMigrationFolder, StrMigrationLoggingFolder))

                ' Add the standard arguments to Argument List
                ArraylistMigrationArguments.Add("/LocalOnly")

                ' *** If not migrating all accounts...
                If Not BlnMigrationSettingsAllUsers Then
                    ' Include current user (either domain or local) and exclude all others
                    ArraylistMigrationArguments.Add("/UI:""" & My.User.CurrentPrincipal.Identity.Name & """")
                    ArraylistMigrationArguments.Add("/UE:*\*")
                    ' *** Otherwise...
                Else
                    ' Exclude accounts older than the specified number of days
                    If IntMigrationExclusionsOlderThanDays > 0 Then
                        ArraylistMigrationArguments.Add("/UEL:" & IntMigrationExclusionsOlderThanDays & "")
                    End If
                    ' Exclude domain accounts specified in settings file
                    For Each exclusion As String In ArrayMigrationExclusionsDomain
                        ArraylistMigrationArguments.Add("/UE:""" & StrEnvDomain & "\" & exclusion.Trim & """")
                    Next
                    ' If not migrating local accounts...
                    If Not BlnMigrationSettingsLocalAccounts Then
                        ' Exclude local accounts
                        ArraylistMigrationArguments.Add("/UE:" & StrEnvComputerName & "\*")
                        ' *** Otherwise...
                    Else
                        ' Exclude local accounts specified in settings file
                        For Each exclusion As String In ArrayMigrationExclusionsLocal
                            ArraylistMigrationArguments.Add(
                                "/UE:""" & StrEnvComputerName & "\" & exclusion.Trim & """")
                        Next
                    End If
                End If

            Case "LOADSTATE"

                ' If migrating all accounts
                If BlnMigrationSettingsAllUsers Then

                    If StrMigrationRestoreAccountsPassword = "" Then
                        ' Create local accounts with blank password
                        ArraylistMigrationArguments.Add("/LAC")
                    Else
                        ' Create local accounts with pre-specified password
                        ArraylistMigrationArguments.Add("/LAC:""" & StrMigrationRestoreAccountsPassword & """")
                    End If

                    If BlnMigrationRestoreAccountsEnabled Then
                        ' Set new locally created accounts to Enabled
                        ArraylistMigrationArguments.Add("/LAE")
                    End If

                End If

                ' If a Domain Change is specified
                If Not StrMigrationDomainChange = "" Then
                    ArraylistMigrationArguments.Add("/MD:" & StrMigrationDomainChange & ":" & StrEnvDomain)
                End If

        End Select

        ' *** Set the migration location to the dataStore
        ArraylistMigrationArguments.Add("""" & Path.Combine(StrMigrationFolder, StrMigrationDataStoreFolder) & """")

        ' ***  Test Mode (No Compression)
        If BlnMigrationCompressionDisabled Then
            ArraylistMigrationArguments.Add("/NoCompress")
        End If

        ' *** Get Encryption Settings
        If Not BlnMigrationEncryptionDisabled Then
            If Not BlnMigrationCompressionDisabled Then
                Select Case StrMigrationType
                    Case "SCANSTATE"
                        ArraylistMigrationArguments.Add("/Encrypt")
                    Case "LOADSTATE"
                        ArraylistMigrationArguments.Add("/Decrypt")
                End Select
                If Not BlnMigrationEncryptionCustom Then
                    ' Set scanstate arguments to use standard encryption key
                    ArraylistMigrationArguments.Add("/Key:""" & StrMigrationEncryptionDefaultKey & """")
                Else
                    FormCustomEncryption.tbxCustomEncryptionKey1.Text = StrCustomEncryptionKey
                    ' Set scanstate arguments to use custom encryption key (built from standard key, and user specified)
                    ArraylistMigrationArguments.Add(
                        "/Key:""" & StrMigrationEncryptionDefaultKey & StrCustomEncryptionKey & """")
                End If
            End If
        End If

        ArraylistMigrationArguments.Add("/V:" & IntMigrationUsmtLoggingType)
    End Sub

    Private Sub sub_MigrationStart()

        Sub_DebugMessage()
        Sub_DebugMessage("* Start Migration *")

        StrPreviousStatusMessage = Nothing
        StrStatusMessage = Nothing
        DtmStartTime = Now
        ' Set the migration type
        _classMigration.Type = StrMigrationType

        ' Add the standard arguments to Argument List
        ArraylistMigrationArguments.Add("/R:5")
        ArraylistMigrationArguments.Add("/W:3")
        ArraylistMigrationArguments.Add("/C")

        ' Transfer the Arguments list to the Migration Class
        _classMigration.Arguments = ArraylistMigrationArguments


        ' *** Run Pre-Migration Scripts
        Sub_DebugMessage("Running Pre-Migration Scripts...")
        If Not func_MigrationSupportScripts("Pre") Then
            label_MigrationCurrentPhase.ForeColor = Color.Red
            label_MigrationCurrentPhase.Text = migrationPreMigrationScriptFail
            If BlnAppProgressOnlyMode Then
                AppShutdown(CInt(exitCodePreMigrationScriptFail))
            End If
            Exit Sub
        End If

        ' Make sure the Data Size Checks only run once
        BlnSizeChecksDone = False

        BlnMigrationInProgress = True

        Try
            Sub_DebugMessage("Spinning up Migration class...")
            _classMigration.Spinup()

        Catch ex As Exception
            Sub_DebugMessage(ex.Message, True)
            label_MigrationCurrentPhase.Text = ex.Message
            _classMigration.SpinDown()
            Exit Sub
        End Try
    End Sub

    Private Sub sub_SupportFormControlsReset()

        Sub_DebugMessage()
        Sub_DebugMessage("* Form Controls Reset *")

        ' Focus the application and reset the in progress items
        TopMost = True
        TopMost = False

        ' If cancelled, update Status
        If BlnMigrationCancelled Then
            label_MigrationCurrentPhase.ForeColor = Color.Red
            label_MigrationCurrentPhase.Text = LabelMigrationCancelled
        End If

        ' Disable / Update status items
        progressbar_Migration.Visible = False
        label_MigrationEstTimeRemaining.Text = Elipses
        label_MigrationEstSize.Text = Elipses

        ' Enable / Update form controls
        button_AdvancedSettings.Enabled = True
        tabcontrol_MigrationType.Enabled = True
        button_Start.Text = LabelStartButton

        ' Remove unwanted text
        label_MigrationEstSize.Text = Nothing
        label_MigrationEstTimeRemaining.Text = Nothing
    End Sub

    '    End If
    ' Initialise USB Support
    Private Sub sub_USBInitialise()

        Sub_DebugMessage()
        Sub_DebugMessage("* USB Initialisation *")

        ' *** Check if a USB disk drive is connected
        Try
            Sub_DebugMessage("Checking if USB drive is connected...")
            Dim wmiQuery = "SELECT * FROM Win32_DiskDrive"
            Dim searcher As New ManagementObjectSearcher(wmiQuery)
            For Each objQuery As ManagementObject In searcher.Get()
                ' If this is a USB disk, and the size exceeds what is specified in the settings file
                If _
                    CStr(objQuery("InterfaceType")) = "USB" And
                    ((CDbl(objQuery("Size")) / 1048576) >= IntMigrationMinUsbDiskSize) Then
                    Sub_DebugMessage(
                        "Drive Found: " & Func_USBGetDriveLetter(CStr(objQuery("Name"))) & " - " &
                        CStr(objQuery("Caption")))
                    ' Check SMART tolerences
                    Sub_DebugMessage("Drive SMART Status: " & CStr(objQuery("Status")))
                    Select Case CStr(objQuery("Status"))
                        Case "OK"
                            ' Automatically use the USB drive if the setting is present
                            If BlnMigrationUsbAutoUseIfAvailable Then
                                Sub_DebugMessage(
                                    $"Auto-Use of USB drive: {Func_USBGetDriveLetter(CStr(objQuery("Name")))}")
                                BlnMigrationLocationUseUsb = True
                                StrMigrationLocationUsb = Path.GetFullPath(Func_USBGetDriveLetter(CStr(objQuery("Name"))))
                            Else

                                ' Present option to use new drive
                                Sub_DebugMessage("User Dialog...")
                                Dim msgboxResult As DialogResult =
                                        MessageBox.Show(Me,
                                                        $"{usbDeviceDescription} {usbDeviceConnectedStartup} { _
                                                           usbDeviceUseDrive}",
                                                        $"{Func_USBGetDriveLetter(CStr(objQuery("Name")))} - { _
                                                           CStr(objQuery("Caption"))}", MessageBoxButtons.YesNo,
                                                        MessageBoxIcon.Question)
                                Select Case msgboxResult
                                    ' Set useUSBMigrationLocation to true and set new USB drive location
                                    Case DialogResult.Yes
                                        BlnMigrationLocationUseUsb = True
                                        StrMigrationLocationUsb =
                                            Path.GetFullPath(Func_USBGetDriveLetter(CStr(objQuery("Name"))))
                                        Sub_DebugMessage(
                                            "User chose to use drive: " &
                                            Func_USBGetDriveLetter(CStr(objQuery("Name"))))
                                        Exit Sub
                                    Case Else
                                        Sub_DebugMessage(
                                            "User chose not to use drive: " &
                                            Func_USBGetDriveLetter(CStr(objQuery("Name"))))
                                End Select
                            End If
                        Case Else
                            ' Display error about drive
                            Sub_DebugMessage("ERROR: " & usbDeviceDescription & " " &
                                             usbDeviceConnectedStartup & " " &
                                             usbDeviceSMARTFail1 &
                                             vbNewLine & vbNewLine & usbDeviceSMARTFail2 & vbNewLine &
                                             Func_USBGetDriveLetter(CStr(objQuery("Name"))) & " - " &
                                             CStr(objQuery("Caption")),
                                             True)
                    End Select
                End If
            Next

        Catch ex As Exception
            Sub_DebugMessage("ERROR: Failed to determine if USB drive was connected. " & ex.Message)
        End Try

        ' *** Start watching USB device State Change Events
        Dim wmiEventQuery =
                "SELECT * FROM __InstanceOperationEvent WITHIN 10 WHERE TargetInstance ISA ""Win32_DiskDrive"""
        _eventwatcherUsbStateChange = New ManagementEventWatcher(wmiEventQuery)
        Sub_DebugMessage("Starting USB event watcher...")
        Try
            _eventwatcherUsbStateChange.Start()
            Sub_DebugMessage("USB event watcher started")
        Catch ex As Exception
            Sub_DebugMessage("ERROR: Failed to start USB event watcher. " & ex.Message)
        End Try
    End Sub

    '        MsgBox("Decrypted: " & encryption_DecryptedData.Text)
    ' Monitor USB State Change Events
    Private Sub sub_USBStateChangeEvent(sender As Object, e As EventArrivedEventArgs) _
        Handles _eventwatcherUsbStateChange.EventArrived

        Sub_DebugMessage()
        Sub_DebugMessage("* USB State Change Event *")

        Try
            Dim objBase, objQuery As ManagementBaseObject
            objBase = e.NewEvent
            objQuery = CType(objBase("TargetInstance"), ManagementBaseObject)

            Sub_DebugMessage("Checking if USB drive has been connected / disconnected...")
            Select Case objBase.ClassPath.ClassName
                ' If creation event...
                Case "__InstanceCreationEvent"
                    ' If this is a USB disk, and the size exceeds what is specified in the settings file
                    If _
                        CStr(objQuery("InterfaceType")) = "USB" And
                        (CDbl(objQuery("Size")) / 1048576) >= IntMigrationMinUsbDiskSize _
                        Then
                        Sub_DebugMessage(
                            "Drive Found: " & Func_USBGetDriveLetter(CStr(objQuery("Name"))) & " - " &
                            CStr(objQuery("Caption")))
                        ' Check SMART tolerences
                        Sub_DebugMessage("Drive SMART Status: " & CStr(objQuery("Status")))
                        Select Case CStr(objQuery("Status"))
                            Case "OK"
                                ' Check if USB migration is already selected (ie, prior USB device selected)
                                Select Case BlnMigrationLocationUseUsb
                                    Case True
                                        ' Present option to use new drive instead
                                        Sub_DebugMessage("User Dialog...")
                                        Dim msgboxResult As DialogResult = MessageBox.Show(Me,
                                                                                           $"{ _
                                                                                              usbDeviceDescriptionAdditional _
                                                                                              } {usbDeviceConnectedEvent _
                                                                                              } {usbDeviceUseDrive}",
                                                                                           $"{Func_USBGetDriveLetter(
                                                                                               CStr(objQuery("Name"))) _
                                                                                              } - { _
                                                                                              CStr(objQuery("Caption")) _
                                                                                              }",
                                                                                           MessageBoxButtons.YesNo,
                                                                                           MessageBoxIcon.Question)
                                        Select Case msgboxResult
                                            ' Set new USB drive location
                                            Case DialogResult.Yes
                                                StrMigrationLocationUsb =
                                                    Path.GetFullPath(Func_USBGetDriveLetter(CStr(objQuery("Name"))))
                                                Sub_DebugMessage(
                                                    "User chose to use drive: " &
                                                    Func_USBGetDriveLetter(CStr(objQuery("Name"))))
                                                Exit Sub
                                            Case Else
                                                Sub_DebugMessage(
                                                    $"User chose not to use drive: { _
                                                                    Func_USBGetDriveLetter(CStr(objQuery("Name")))}")
                                        End Select
                                        ' Present option to use drive
                                    Case False
                                        Dim msgboxResult As DialogResult =
                                                MessageBox.Show(Me,
                                                                $"{usbDeviceDescription} {usbDeviceConnectedEvent} { _
                                                                   usbDeviceUseDrive}",
                                                                $"{Func_USBGetDriveLetter(CStr(objQuery("Name")))} - { _
                                                                   CStr(objQuery("Caption"))}",
                                                                MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                                        Select Case msgboxResult
                                            ' Set useUSBMigrationLocation to true and set new USB drive location
                                            Case DialogResult.Yes
                                                BlnMigrationLocationUseUsb = True
                                                StrMigrationLocationUsb =
                                                    Path.GetFullPath(Func_USBGetDriveLetter(CStr(objQuery("Name"))))
                                                Exit Sub
                                        End Select
                                End Select
                            Case Else
                                ' Display error about drive
                                Sub_DebugMessage("ERROR: " & usbDeviceDescription & " " &
                                                 usbDeviceConnectedEvent & " " &
                                                 usbDeviceSMARTFail1 &
                                                 vbNewLine & vbNewLine & usbDeviceSMARTFail2 & vbNewLine &
                                                 Func_USBGetDriveLetter(CStr(objQuery("Name"))) & " - " &
                                                 CStr(objQuery("Caption")), True)
                        End Select

                    End If
                    ' If deletion event...
                Case "__InstanceDeletionEvent"
                    If BlnMigrationLocationUseUsb = True Then
                        Sub_DebugMessage("USB drive has been removed. USB mode disabled")
                        ' Set usbUSBDrive to false and display message
                        BlnMigrationLocationUseUsb = False
                        Sub_DebugMessage("WARNING: " & usbDeviceDescriptionCurrent & " " &
                                         usbDeviceDisconnectedEvent & " " &
                                         vbNewLine & vbNewLine & usbDeviceSwitchToStandard & vbNewLine &
                                         CStr(objQuery("Caption")), True)
                    End If
            End Select

        Catch ex As Exception
            Sub_DebugMessage("ERROR: Failed to determine if USB drive was connected / disconnected. " & ex.Message)
        End Try
    End Sub

    ' Tab Control
    Private Sub tabcontrol_MigrationType_SelectedIndexChanged(sender As Object, e As EventArgs) _
        Handles tabcontrol_MigrationType.SelectedIndexChanged, tabcontrol_MigrationType.DrawItem

        Sub_DebugMessage()
        Sub_DebugMessage("* Tab Change / Drawn Events * ")

        ' Advanced / Back / Forward button setup
        Select Case tabcontrol_MigrationType.SelectedIndex

            Case tabcontrol_MigrationType.TabPages.IndexOf(tabpage_Capture)

                Sub_DebugMessage("Capture Tab Selected")
                StrMigrationType = "SCANSTATE"

            Case tabcontrol_MigrationType.TabPages.IndexOf(tabpage_Restore)

                Sub_DebugMessage("Restore Tab Selected")
                StrMigrationType = "LOADSTATE"

                label_DatastoreLocation.Text = StatusLabelSearching
                Application.DoEvents()

                Sub_MigrationFindDataStore()

        End Select

        button_Start.Focus()
    End Sub

#End Region

    'Private Sub sub_InitCheckEncryption()

    '    sub_DebugMessage()
    '    sub_DebugMessage("* Check Encryption *")

    '    ' If encryption is disabled, skip
    '    If bln_MigrationEncryptionDisabled Then

    '        sub_DebugMessage("Encryption is disabled. Skipping...")
    '        Exit Sub
    '    End If

    '    ' If the key is not already encrypted
    '    If Not bln_MigrationEncryptionDefaultKeyEncrypted Then

    '        sub_DebugMessage("Key is not previously encrypted. Encrypting...")

    '        ' Encrypt Key
    '        Dim encryption_EncryptedData As Encryption.Data = _
    '            encryption_SymmetricEncryption.Encrypt(New Encryption.Data(str_MigrationEncryptionDefaultKey), encryption_DataHash)

    '        My.Settings.MigrationEncryptionDefaultKey = encryption_EncryptedData.Text
    '        My.Settings.MigrationEncryptionDefaultKeyEncrypted = True
    '        My.Settings.Save()

    '        MsgBox("Encrypted: " & encryption_EncryptedData.Text)

    '        ' Decrypt Key
    '        Dim encryption_DecryptedData As Encryption.Data = _
    '            encryption_SymmetricEncryption.Decrypt(encryption_EncryptedData, encryption_DataHash)
End Class
