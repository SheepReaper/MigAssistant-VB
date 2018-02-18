Imports System.Threading

Public Class ClassHealthCheck
    Implements IDisposable

#Region "Fields"

    Private _blnHealthCheckInProgress As Boolean = False
    Private _intHealthCheckExitCode As Integer = 0
    Private _intHealthCheckPercentComplete As Integer = 0
    Private _process As Process
    Private _strHealthCheckOutput As String = Nothing
    Private _strHealthCheckProgress As String = Nothing
    Private _thread As Thread

#End Region

#Region "Events"

    Public Event HealthCheckFinished(sender As Object, e As EventArgs)

    Public Event ProgressUpdate(sender As Object, e As EventArgs)

#End Region

#Region "Properties"

    Public ReadOnly Property ExitCode As Integer
        Get
            Return _intHealthCheckExitCode
        End Get
    End Property

    ' Property: Return Yes / No as to whether the scan is in progress
    Public ReadOnly Property InProgress As Boolean
        Get
            Return _blnHealthCheckInProgress
        End Get
    End Property

    Public ReadOnly Property PercentComplete As Integer
        Get
            Return _intHealthCheckPercentComplete
        End Get
    End Property

    Public ReadOnly Property Progress As String
        Get
            Return _strHealthCheckProgress
        End Get
    End Property

#End Region

#Region "Methods"

    Friend Sub SpinDown()

        ' Cleanup
        Sub_DebugMessage("Thread Cleanup...")
        _blnHealthCheckInProgress = False

        ' Ensure the process has terminated
        Try

            If Not _process.HasExited Then
                Sub_DebugMessage("Terminating process...")
                _process.Kill()
            End If
            _process.Close()
            ' _thread = Nothing
        Catch ex As Exception

        End Try
    End Sub

    Public Sub Spinup()
        'Reset variables
        _strHealthCheckOutput = Nothing
        _strHealthCheckProgress = Nothing
        _intHealthCheckExitCode = 0
        _intHealthCheckPercentComplete = 0

        ' Start the scan in a new thread
        _blnHealthCheckInProgress = True

        Dim threadStart As New ThreadStart(AddressOf Start)
        _thread = New Thread(threadStart)
        _thread.Start()
    End Sub

    Private Sub _sub_HealthCheckProgressReturn()

        If Not _strHealthCheckOutput = Nothing Then
            If _strHealthCheckOutput.Contains("file records") Or _strHealthCheckOutput.Contains("parameter specified") _
                Then
                _intHealthCheckPercentComplete = 0
                _strHealthCheckProgress = My.Resources.diskScanStatus1
            ElseIf _
                _strHealthCheckOutput.Contains("index entries") Or _strHealthCheckOutput.Contains("verifying indexes") _
                Then
                _intHealthCheckPercentComplete = 0
                _strHealthCheckProgress = My.Resources.diskScanStatus2
            ElseIf _
                _strHealthCheckOutput.Contains("descriptors") Or
                _strHealthCheckOutput.Contains("verification completed") Then
                _intHealthCheckPercentComplete = 0
                _strHealthCheckProgress = My.Resources.diskScanStatus3
            End If
            If _strHealthCheckOutput.Contains("percent complete") Then
                _intHealthCheckPercentComplete = CInt(_strHealthCheckOutput.Substring(0, 2).Trim.Replace(".",
                                                                                                         StrLocaleDecimal))
            End If
        End If

        ' Raise event that the progress has changed
        If _blnHealthCheckInProgress Then RaiseEvent ProgressUpdate(Me,Nothing)
    End Sub

    Private Sub Start()
        Try
            ' Set up and start the CHKDSK process
            _process = New Process
            Dim processInfo As New ProcessStartInfo("CHKDSK", "C: /I /C") With {
                    .RedirectStandardError = True,
                    .RedirectStandardOutput = True,
                    .RedirectStandardInput = True,
                    .CreateNoWindow = True,
                    .UseShellExecute = False
                    }
            _process.StartInfo = processInfo
            _process.Start()

            ' Continously check the output from CHKDSK and dump to a global variable
            Do While Not _process.StandardOutput.EndOfStream
                If _process.StandardOutput.ReadLine.Length <> 0 Then
                    _strHealthCheckOutput = _process.StandardOutput.ReadLine
                    _sub_HealthCheckProgressReturn()
                End If
            Loop

            ' Wait until the process exits and check the error code
            _process.WaitForExit()

            ' Get the exit code
            _intHealthCheckExitCode = _process.ExitCode

            ' Prevent error for appearing when process is aborted and STDOut gets redirected
        Catch exInvalidOperation As InvalidOperationException
            Sub_DebugMessage("WARNING: Health Check has been cancelled.")
            ' Return Cancelled Error Code
            _intHealthCheckExitCode = 999
        Catch ex As Exception
            Sub_DebugMessage($"ERROR: Heath Check Failed: {ex.Message}", True)
        Finally
            Sub_DebugMessage("RaiseEvent: HealthCheckFinished")
            RaiseEvent HealthCheckFinished(Me,Nothing)
        End Try
    End Sub
    

#End Region

    Public Overridable Sub Dispose() Implements IDisposable.Dispose
        _process?.Dispose()
    End Sub
End Class