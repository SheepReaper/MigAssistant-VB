Imports System.IO
Imports Microsoft.VisualBasic.FileIO
Imports MigAssistant.Encryption

Module ModuleGlobalVariables
    ' P/Invokes
    Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA"(lpDirectoryName As String,
                                                                                           ByRef _
                                                                                              lpFreeBytesAvailableToCaller _
                                                                                              As Long,
                                                                                           ByRef lpTotalNumberOfBytes As _
                                                                                              Long,
                                                                                           ByRef _
                                                                                              lpTotalNumberOfFreeBytes _
                                                                                              As Long) As Long

    ' Constants
    Public Const StrMigrationDataStoreFolder As String = "Datastore"
    Public Const StrMigrationLoggingFolder As String = "Logging"
    Public Const StrMigrationXmlConfigName As String = "Migration.XML"

    ' Locale Settings
    Public StrLocaleDecimal As String = Mid(CStr(11/10), 2, 1)
    Public StrLocaleComma As String = Chr(90 - Asc(strLocaleDecimal))

    ' Set up encryption type
    Public EncryptionSymmetricEncryption As New Symmetric(Symmetric.Provider.TripleDES)
    ' Set encryption key
    Public EncryptionDataHash As New Data("1kb3n33nb3st")

    ' Get OS Information
    Public DblOsVersion As Double = CDbl(Left(My.Computer.Info.OSVersion, 3).Replace(".", strLocaleDecimal))
    Public StrOsFullName As String = My.Computer.Info.OSFullName
    Public StrOsArchitecture As String = Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE")

    ' Current Workstation Information
    Public StrEnvDomain As String = Environment.UserDomainName

    Public _
        StrEnvUserName As String = Replace(My.User.CurrentPrincipal.Identity.Name, Environment.UserDomainName & "\", "",
                                            , , CompareMethod.Text)

    Public StrEnvComputerName As String = My.Computer.Name

    ' Set Up Variables
    Public ArraylistMigrationArguments As New ArrayList
    Public ArraylistScriptsCurrent As New ArrayList
    Public BlnHealthCheck As Boolean = False
    Public BlnHealthCheckStatusOk As Boolean = False
    Public BlnMigrationStatusOk As Boolean = False
    Public BlnAppProgressOnlyMode As Boolean = False
    Public StrMigrationType As String = Nothing
    Public StrMigrationFolder As String = Nothing
    Public BlnMigrationCancelled As Boolean = False
    Public StrUsmtguid As String = Nothing
    Public ObjLogFile As TextWriter
    Public StrPreviousStatusMessage As String = Nothing
    Public StrStatusMessage As String = Nothing
    Public DtmStartTime As DateTime = Nothing
    Public BlnSizeChecksDone As Boolean = False
    Public BlnHealthCheckInProgress As Boolean = False
    Public BlnMigrationInProgress As Boolean = False
    Public BlnDownloadComplete As Boolean = False
    Public IntDownloadProgress As Integer = 0
    Public StrPrimaryDataDrive As String = Nothing

    ' Get Application Information
    Public StrUsmtFolder As String = My.Computer.FileSystem.SpecialDirectories.ProgramFiles & "\USMT301"
    Public StrWmaFolder As String = My.Application.Info.DirectoryPath
    Public StrTempFolder As String = Path.GetTempPath.TrimEnd("\"(0))
    Public StrLogFile As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\WMA.Log"
    Public StrWmaConfigNetworkCheck As String = "\\ServerName\MigrationShare"

    ' Get Settings from .Exe.Settings file
    Public StrMigrationConfigFile As String = My.Settings.MigrationConfig
    Public ArrayMigrationExclusionsDomain() As String = Split(My.Settings.MigrationExclusionsDomain, ",")
    Public ArrayMigrationExclusionsLocal() As String = Split(My.Settings.MigrationExclusionsLocal, ",")
    Public BlnMigrationMultiUserMode As Boolean = My.Settings.MigrationMultiUserMode
    Public IntMigrationExclusionsOlderThanDays As Integer = My.Settings.MigrationExclusionsOlderThanDays
    Public IntMigrationUsmtLoggingType As Integer = My.Settings.USMTLoggingValue
    Public StrMigrationLocationNetwork As String = My.Settings.MigrationNetworkLocation
    Public BlnMigrationLocationNetworkDisabled As Boolean = My.Settings.MigrationNetworkLocationDisabled
    Public IntMigrationMaxSize As Integer = My.Settings.MigrationMaxSize
    Public BlnMigrationEncryptionDisabled As Boolean = My.Settings.MigrationEncryptionDisabled
    Public StrMigrationEncryptionDefaultKey As String = My.Settings.MigrationEncryptionDefaultKey
    Public StrMigrationRestoreAccountsPassword As String = My.Settings.MigrationRestoreAccountsPassword
    Public BlnMigrationRestoreAccountsEnabled As Boolean = My.Settings.MigrationRestoreAccountsEnabled
    Public BlnSettingsAdvancedSettingsDisabled As Boolean = My.Settings.SettingsAdvancedSettingsDisabled
    Public BlnSettingsHealthCheckDefaultEnabled As Boolean = My.Settings.SettingsHealthCheckDefaultEnabled
    Public BlnSettingsWorkstationDetailsDisabled As Boolean = My.Settings.SettingsWorkstationDetailsDisabled
    Public BlnSettingsDebugMode As Boolean = My.Settings.SettingsDebugMode
    Public ArrayMigrationRuleSet() As String = Split(My.Settings.MigrationRuleSet, ",")
    Public IntMigrationMinUsbDiskSize As Integer = My.Settings.MigrationUSBMinSize
    Public BlnMigrationUsbAutoUseIfAvailable As Boolean = My.Settings.MigrationUSBAutoUseIfAvailable
    Public BlnMigrationCompressionDisabled As Boolean = My.Settings.MigrationCompressionDisabled
    Public StrMigrationDomainChange As String = My.Settings.MigrationDomainChange
    Public ArrayMigrationScriptsPreCapture() As String = Split(My.Settings.MigrationScriptsPreCapture, ",")
    Public ArrayMigrationScriptsPostCapture() As String = Split(My.Settings.MigrationScriptsPostCapture, ",")
    Public ArrayMigrationScriptsPreRestore() As String = Split(My.Settings.MigrationScriptsPreRestore, ",")
    Public ArrayMigrationScriptsPostRestore() As String = Split(My.Settings.MigrationScriptsPostRestore, ",")
    Public BlnMigrationScriptsNoWindow As Boolean = My.Settings.MigrationScriptsNoWindow
    Public BlnMigrationOverwriteExistingFolders As Boolean = My.Settings.MigrationOverWriteExistingFolders
    Public BlnMailSend As Boolean = My.Settings.MailSend
    Public StrMailServer As String = My.Settings.MailServer
    Public StrMailRecipients As String = My.Settings.MailRecipients
    Public StrMailFrom As String = My.Settings.MailFrom

    ' Resources
    Public StrBddManifestUrl As String = My.Resources.bddManifestURL
    Public StrBddManifestFile As String = My.Resources.bddManifestFile
    Public StrUsmtguiDx86 As String = My.Resources.usmtGUIDx86
    Public StrUsmtguiDx64 As String = My.Resources.usmtGUIDx64

    ' *** Migration Settings
    Public BlnMigrationSettingsAllUsers As Boolean = False
    Public BlnMigrationSettingsLocalAccounts As Boolean = False
    Public BlnMigrationLocationUseOther As Boolean = False
    Public StrMigrationLocationOther As String = Nothing
    Public BlnMigrationLocationUseUsb As Boolean = False
    Public StrMigrationLocationUsb As String = Nothing
    Public BlnMigrationEncryptionCustom As Boolean = False
    Public StrCustomEncryptionKey As String = "0"
    Public BlnMigrationMaxOverride As Boolean = False
    Public BlnMigrationFolderOverride As Boolean = False

    Public Sub Sub_DebugMessage(Optional ByVal strDebugMessage As String = Nothing,
                                Optional ByVal blnDisplayError As Boolean = False,
                                Optional ByVal blnWriteEventLogEntry As Boolean = False)

        If blnDisplayError = True Then
            If strDebugMessage.Contains("INFO:") Then
                MessageBox.Show(strDebugMessage,
                                My.Resources.appTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
            ElseIf strDebugMessage.Contains("WARNING:") Then
                MessageBox.Show(strDebugMessage, My.Resources.appTitle, MessageBoxButtons.OK,
                                MessageBoxIcon.Exclamation)
            ElseIf strDebugMessage.Contains("ERROR:") Then
                MessageBox.Show(strDebugMessage, My.Resources.appTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If

        If blnWriteEventLogEntry Then
            My.Application.Log.WriteEntry(strDebugMessage)
        End If

        strDebugMessage = "[" & DateTime.Now & "] " & strDebugMessage

        If BlnSettingsDebugMode And Not ObjLogFile Is Nothing Then
            ObjLogFile.WriteLine(strDebugMessage)
            ObjLogFile.Flush()
        Else
            Exit Sub
        End If

        Console.WriteLine(strDebugMessage)
    End Sub

    Public Sub AppInitialise()

        Sub_DebugMessage()
        Sub_DebugMessage("* Application Initialisation *")

        ' Create the logfile if in Debug mode, otherwise, just output to the console...
        If BlnSettingsDebugMode Then
            Sub_DebugMessage("Running in Debug Mode")
            Try
                ' Check if the logfile already exists. If yes, delete it
                If My.Computer.FileSystem.FileExists(StrLogFile) Then
                    Sub_DebugMessage("Logfile already exists. Attempting to delete...")
                    Try
                        My.Computer.FileSystem.DeleteFile(StrLogFile, UIOption.OnlyErrorDialogs,
                                                          RecycleOption.DeletePermanently)
                        Sub_DebugMessage("Logfile deleted")
                    Catch ex As Exception
                        Throw New Exception(ex.Message)
                    End Try
                End If

                ' Connect to the logfile
                Try
                    Sub_DebugMessage("Initialising logfile...")
                    ObjLogFile = My.Computer.FileSystem.OpenTextFileWriter(StrLogFile, True)
                Catch ex As Exception
                    Throw New Exception(ex.Message)
                End Try

            Catch ex As Exception
                MessageBox.Show($"ERROR: Unable to create Debug log file: {ex.Message}. Debugging switched off",
                                My.Resources.appTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                BlnSettingsDebugMode = False
            End Try
        End If

        Sub_DebugMessage(My.Resources.appTitle & " " & My.Resources.appBuild & " - " & My.Resources.appCompany)
    End Sub

    Public Sub AppShutdown(intExitCode As Integer)

        Sub_DebugMessage()
        Sub_DebugMessage("* Application Shutdown *")

        Sub_DebugMessage("Exiting Application with Exit Code: " & intExitCode, False, True)

        ' Close the logfile
        If BlnSettingsDebugMode Then
            Sub_DebugMessage("Closing Logfile...")
            ObjLogFile.Close()
            ObjLogFile = Nothing
        End If

        Environment.Exit(intExitCode)
    End Sub

    Public Function Func_GetFreeSpace(strLocation As String) As Long

        Dim lngBytesTotal, lngFreeBytes, lngFreeBytesAvailable, lngResult As Long
        lngResult = GetDiskFreeSpaceEx(strLocation, lngFreeBytesAvailable, lngBytesTotal, lngFreeBytes)
        If lngResult > 0 Then
            Return Func_BytesToMB(lngFreeBytes)
        Else
            Throw New Exception("ERROR: Invalid or unreadable location")
        End If
    End Function

    Private Function Func_BytesToMB(lngBytes As Long) As Long

        Dim dblResult As Double
        dblResult = (lngBytes/1024)/1024
        Func_BytesToMB = CLng(Format(dblResult, "###,###,##0.00"))
    End Function
End Module
