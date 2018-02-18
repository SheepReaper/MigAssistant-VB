Imports System.IO
Imports System.Threading
Imports Microsoft.VisualBasic.FileIO

Public Class ClassMigration

#Region "Declarations"

    ' Event Declarations
    Private WithEvents _fswMigrationProgressFileWatcher As New FileSystemWatcher

    ' Class Declarations
    Private _strMigrationArguments As String = Nothing
    Private _arraylistMigrationArguments As New ArrayList
    Private _strMigrationProgress As String = Nothing
    Private _strMigrationDebugInfo As String = Nothing
    Private _intMigrationEstDataSize As Integer = 0
    Private _intMigrationExitCode As Integer = 0
    Private _blnMigrationInProgress As Boolean = False
    Private _strMigrationLogFile As String = Nothing
    Private _intMigrationEstTimeRemaining As Integer = 0
    Private _intMigrationPercentComplete As Integer = 0
    Private _strMigrationProgressFile As String = Nothing
    Private _ioMigrationProgressFileParser As TextFieldParser
    Private _intMigrationProgressFileLastLineNumber As Integer = 0
    Private _blnMigrationProgressCheckInProgress As Boolean = False
    Private _strMigrationType As String = Nothing
    Private _thread As Thread
    Private _process As Process

#End Region

#Region "Properties"
    ' Property: Return the argument list for USMT
    Public Property Arguments As ArrayList
        Get
            Return _arraylistMigrationArguments
        End Get
        Set
            _arraylistMigrationArguments = value
        End Set
    End Property
    ' Property: Return the migration type
    Public Property Type As String
        Get
            Return _strMigrationType
        End Get
        Set
            _strMigrationType = value
        End Set
    End Property
    ' Property: Return whether the migration is in progress
    Public ReadOnly Property InProgress As Boolean
        Get
            Return _blnMigrationInProgress
        End Get
    End Property
    ' Property: Return the migration current progress
    Public ReadOnly Property Progress As String
        Get
            Return _strMigrationProgress
        End Get
    End Property
    ' Property: Return the migration exit code
    Public ReadOnly Property ExitCode As Integer
        Get
            Return _intMigrationExitCode
        End Get
    End Property
    ' Property: Return the migration current debug info
    Public ReadOnly Property DebugInfo As String
        Get
            Return _strMigrationDebugInfo
        End Get
    End Property
    ' Property: Return the migration data size
    Public ReadOnly Property EstDataSize As Integer
        Get
            Return _intMigrationEstDataSize
        End Get
    End Property
    ' Property: Return the migration minutes remaining
    Public ReadOnly Property EstTimeRemaining As Integer
        Get
            Return _intMigrationEstTimeRemaining
        End Get
    End Property
    ' Property: Return the migration percent complete
    Public ReadOnly Property PercentComplete As Integer
        Get
            Return _intMigrationPercentComplete
        End Get
    End Property
    ' Property: Return the path to the migration log file
    Public ReadOnly Property LogFile As String
        Get
            Return StrMigrationFolder & "\" & StrMigrationLoggingFolder & "\" & _strMigrationLogFile
        End Get
    End Property

#End Region

#Region "Events"

    Public Event ProgressUpdate()
    Public Event MigrationFinished()

#End Region

#Region "Subroutines"

    Private Sub Start()
        Try
            ' If the Migration Progress File exists, delete it
            If My.Computer.FileSystem.FileExists(StrMigrationFolder & "\" &
                                                 StrMigrationLoggingFolder & "\" & _strMigrationProgressFile) Then
                My.Computer.FileSystem.DeleteFile(StrMigrationFolder & "\" &
                                                  StrMigrationLoggingFolder & "\" & _strMigrationProgressFile)
            End If

            ' Add Progress and Logfile details to the argument list
            _arraylistMigrationArguments.Add(
                "/Progress:""" & StrMigrationFolder & "\" & StrMigrationLoggingFolder & "\" &
                _strMigrationProgressFile & """")
            _arraylistMigrationArguments.Add(
                "/L:""" & StrMigrationFolder & "\" & StrMigrationLoggingFolder & "\" & _strMigrationLogFile & """")
            ' Sort and Dump the arguments from the array into a string containing spaces
            _arraylistMigrationArguments.Sort()
            _strMigrationArguments = Join(_arraylistMigrationArguments.ToArray, " ")

            ' Setup and configure the new process
            _process = New Process
            Dim processInfo As New ProcessStartInfo($"{StrUsmtFolder}\{_strMigrationType}.Exe", _strMigrationArguments) With {
                .WorkingDirectory = StrUsmtFolder,
                .UseShellExecute = True,
                .WindowStyle = ProcessWindowStyle.Hidden
                                                     }
            _process.StartInfo = processInfo

            ' Start the Migration Process
            _process.Start()
            _process.WaitForExit()
            _intMigrationExitCode = _process.ExitCode

            ' Prevent error for appearing when process is aborted and STDOut gets redirected
        Catch exInvalidOperation As InvalidOperationException
            Sub_DebugMessage("WARNING: Migration has been cancelled.")
            ' Return Cancelled Error Code
            _intMigrationExitCode = 999
        Catch ex As Exception
            MsgBox("An error occurred during the Migration:" & vbNewLine & vbNewLine & "Error: " & ex.Message,
                   MsgBoxStyle.Critical, My.Resources.appTitle)
        Finally
            Sub_DebugMessage("RaiseEvent: MigrationFinished")
            RaiseEvent MigrationFinished()
        End Try
    End Sub

    Public Sub Spinup()
        ' Reset Variables
        _strMigrationProgress = Nothing
        _strMigrationDebugInfo = Nothing
        _intMigrationEstDataSize = 0
        _intMigrationEstTimeRemaining = 0
        _intMigrationExitCode = 0
        _intMigrationPercentComplete = 0
        _intMigrationProgressFileLastLineNumber = 0
        _strMigrationLogFile = _strMigrationType & ".Log"
        _strMigrationProgressFile = StrMigrationType & "_Progress.Log"

        ' Set up the Progress File watcher
        With _fswMigrationProgressFileWatcher
            .Path = StrMigrationFolder & "\" & StrMigrationLoggingFolder
            .Filter = _strMigrationProgressFile
            .IncludeSubdirectories = False
            .EnableRaisingEvents = True
        End With

        ' Start the migration in a new thread
        _blnMigrationInProgress = True

        Dim threadStart As New ThreadStart(AddressOf Me.Start)
        _thread = New Thread(threadStart)
        _thread.Start()
    End Sub

    Public Sub SpinDown()

        ' Cleanup
        Sub_DebugMessage("Thread Cleanup...")
        _fswMigrationProgressFileWatcher.EnableRaisingEvents = False
        _fswMigrationProgressFileWatcher.Dispose()
        _blnMigrationInProgress = False

        Try

            ' Ensure the process has terminated
            If Not _process.HasExited Then
                Sub_DebugMessage("Terminating process...")
                _process.Kill()
            End If
            _process.Close()
            ' _thread = Nothing

        Catch ex As Exception

        End Try
    End Sub

    Private Sub _sub_MigrationProgressReturn(sender As Object, e As FileSystemEventArgs) _
        Handles _fswMigrationProgressFileWatcher.Changed

        ' If we're currently checking the progress, don't monitor it again
        ' sub_DebugMessage("DEBUG: Checking if progress-monitoring is already running...")
        If _blnMigrationProgressCheckInProgress Then Exit Sub

        ' Sleep to ensure file writing is complete
        Thread.Sleep(500)

        ' Set up initial values
        Dim arrayMigrationProgressFileCurrentRow As String() = Nothing
        Dim intMigrationProgressFileCurrentLine = 0

        ' Read Progress File into the parser
        Try
            ' sub_DebugMessage("DEBUG: Reading Progress File: " & str_MigrationFolder & "\" & str_MigrationLoggingFolder & "\" & _str_MigrationProgressFile)
            _ioMigrationProgressFileParser = My.Computer.FileSystem.OpenTextFieldParser(StrMigrationFolder & "\" &
                                                                                         StrMigrationLoggingFolder &
                                                                                         "\" &
                                                                                         _strMigrationProgressFile)
            _ioMigrationProgressFileParser.TextFieldType = FieldType.Delimited
            _ioMigrationProgressFileParser.Delimiters = New String() {","}
            _ioMigrationProgressFileParser.HasFieldsEnclosedInQuotes = True
            _ioMigrationProgressFileParser.TrimWhiteSpace = True
        Catch ex As Exception
            ' sub_DebugMessage("DEBUG: EXCEPTION CAUGHT: " & ex.Message)
            ' Exit the sub if failed (ie, our file doesn't exist yet)
            Exit Sub
        End Try

        ' Go through each line in the log file
        While Not _ioMigrationProgressFileParser.EndOfData
            ' Make sure we exit the loop if the migration has been cancelled
            If Not _fswMigrationProgressFileWatcher.EnableRaisingEvents Then
                Sub_DebugMessage("Stopping progress monitoring...")
                Exit While
            End If
            ' sub_DebugMessage("DEBUG: File not finished")
            _blnMigrationProgressCheckInProgress = True
            Try
                ' Read the current line into an array
                ' sub_DebugMessage("DEBUG: Reading current line...")
                arrayMigrationProgressFileCurrentRow = _ioMigrationProgressFileParser.ReadFields()
                ' sub_DebugMessage("DEBUG: " & Join(_array_MigrationProgressFileCurrentRow, ","))
                ' If this line number is higher than the last line, process it

                If _
                    UBound(arrayMigrationProgressFileCurrentRow) > 3 And
                    intMigrationProgressFileCurrentLine > _intMigrationProgressFileLastLineNumber Then
                    Select Case arrayMigrationProgressFileCurrentRow(3)
                        Case "PHASE"
                            Select Case arrayMigrationProgressFileCurrentRow(4)
                                Case "Initializing"
                                    _strMigrationProgress = My.Resources.migrationPhaseInitializing
                                Case "Scanning"
                                    _strMigrationProgress = My.Resources.migrationPhaseScanning
                                Case "Collecting"
                                    _strMigrationProgress = My.Resources.migrationPhaseCollecting
                                Case "Saving"
                                    _strMigrationProgress = My.Resources.migrationPhaseSaving
                                    _strMigrationDebugInfo = Nothing
                                Case "Estimating"
                                    _strMigrationProgress = My.Resources.migrationPhaseEstimating
                                Case "Applying"
                                    _strMigrationProgress = My.Resources.migrationPhaseApplying
                            End Select
                        Case "totalSizeInMBToTransfer"
                            If Not arrayMigrationProgressFileCurrentRow(4) = "" Then _
                                _intMigrationEstDataSize = CInt(arrayMigrationProgressFileCurrentRow(4).Replace(".",
                                                                                                                   strLocaleDecimal))
                        Case "totalPercentageCompleted"
                            If Not arrayMigrationProgressFileCurrentRow(4) = "" Then _
                                _intMigrationPercentComplete =
                                    CInt(arrayMigrationProgressFileCurrentRow(4).Replace(".",
                                                                                           strLocaleDecimal))
                        Case "totalMinutesRemaining"
                            If Not arrayMigrationProgressFileCurrentRow(4) = "" Then _
                                _intMigrationEstTimeRemaining =
                                    CInt(arrayMigrationProgressFileCurrentRow(4).Replace(".",
                                                                                           strLocaleDecimal))
                        Case "detectedUser"
                            If arrayMigrationProgressFileCurrentRow(6) = "Yes" Then
                                _strMigrationDebugInfo = My.Resources.migrationDetectedUser & " " &
                                                          arrayMigrationProgressFileCurrentRow(4)
                            End If
                        Case "forUser"
                            If arrayMigrationProgressFileCurrentRow(8) = "Yes" Then
                                _strMigrationDebugInfo = My.Resources.migrationForUser & " " &
                                                          arrayMigrationProgressFileCurrentRow(4) & " - " &
                                                          arrayMigrationProgressFileCurrentRow(6)
                            End If
                        Case "collectingUser"
                            _strMigrationDebugInfo = My.Resources.migrationCollectingUser & " " &
                                                      arrayMigrationProgressFileCurrentRow(4)
                        Case "errorCode"
                            If CInt(arrayMigrationProgressFileCurrentRow(4)) = 0 Then
                                _strMigrationDebugInfo = My.Resources.migrationCompleteSuccess & " " &
                                                          My.Resources.migrationNonFatalErrors & " " &
                                                          arrayMigrationProgressFileCurrentRow(6)
                            Else
                                _strMigrationDebugInfo = My.Resources.migrationCompleteError & " " &
                                                          arrayMigrationProgressFileCurrentRow(4)
                            End If
                    End Select
                    ' Update the recorded line number
                    _intMigrationProgressFileLastLineNumber = intMigrationProgressFileCurrentLine
                End If

            Catch ex As MalformedLineException
                Sub_DebugMessage(
                    "Line " & intMigrationProgressFileCurrentLine & " is malformed and will be skipped: " &
                    Join(arrayMigrationProgressFileCurrentRow, ", ") & " - " & ex.Message)
            Catch ex As Exception
                Sub_DebugMessage(
                    "Error occurred while reading line " & intMigrationProgressFileCurrentLine & ": " &
                    Join(arrayMigrationProgressFileCurrentRow, ", ") & " - " & ex.Message)
            Finally
                intMigrationProgressFileCurrentLine = intMigrationProgressFileCurrentLine + 1
                ' Raise event that the progress has changed
                If _blnMigrationInProgress Then RaiseEvent ProgressUpdate()
            End Try
        End While
        _ioMigrationProgressFileParser.Close()
        _blnMigrationProgressCheckInProgress = False
    End Sub

#End Region
End Class
