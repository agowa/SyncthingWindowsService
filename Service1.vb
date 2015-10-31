Imports System.Runtime.InteropServices
Imports System.Threading

Public Class Service1 'extends System.ServiceProcess.ServiceBase
    Private p(5) As Process
    Private pStartInfo(5) As ProcessStartInfo
    Private stopping As Boolean = False
    Private stoppedEvent As ManualResetEvent = New ManualResetEvent(False)

    Public Sub New()
        InitializeComponent()
        Debugging()
    End Sub

    Private Sub Debugging()
        If My.Settings.WaitForDebugger Then
            If My.Settings.DebuggerIsRemote Then
                'Debugging Remotely
                While Not Debugger.IsAttached
                    Thread.Sleep(100)
                End While
            Else
                'Debugging Localy
                Debugger.Launch()
            End If
        End If
    End Sub


    Private _serviceStatus As ServiceStatus
    Private Sub TESTSUB()
        Dim handle As IntPtr = Me.ServiceHandle
        _serviceStatus.dwCurrentState = Fix(ServiceState.SERVICE_START_PENDING)
        SetServiceStatus(handle, _serviceStatus)
        Dim logMessage As String = String.Empty
        
        _serviceStatus.dwCurrentState = Fix(ServiceState.SERVICE_RUNNING)
        SetServiceStatus(handle, _serviceStatus)
        Me.Stop()
    End Sub

    Private Enum Processes
        pStartup = 0
        pMain = 1
        pCheckIfRunning = 2
        pPause = 3
        pContinue = 4
        pStop = 5
    End Enum
    Private Function NewProcess(ByVal process As Processes, ByVal FileName As String, ByVal Arguments As String) As Process
        ' TODO: Cleanup
        Me.p(process) = New Process()
        Me.pStartInfo(process) = New ProcessStartInfo()
        pStartInfo(process).WorkingDirectory = System.IO.Path.GetDirectoryName(FileName)
        pStartInfo(process).FileName = System.IO.Path.GetFullPath(FileName)
        pStartInfo(process).WindowStyle = ProcessWindowStyle.Hidden

        'TODO: Get CMD Windows Output
        'pStartInfo(process).OutputDataReceived += Function(origin, args) WriteTextAsync(args.Data, output)
        'pStartInfo(process).ErrorDataReceived += Function(origin, args) WriteTextAsync(args.Data, output)

        pStartInfo(process).Arguments = Arguments
        pStartInfo(process).RedirectStandardError = True
        pStartInfo(process).RedirectStandardOutput = True
        pStartInfo(process).RedirectStandardInput = True
        pStartInfo(process).UseShellExecute = False
        pStartInfo(process).CreateNoWindow = True
        p(process).EnableRaisingEvents = True
        p(process).StartInfo = pStartInfo(process)

        Return p(process)
    End Function
    Private Function GetProcess(ByVal process As Processes) As Process
        Return Me.p(process)
    End Function

    Private Sub SetServiceStatus(ByVal serviceState As ServiceState)
        Dim serviceStatus As ServiceStatus = New ServiceStatus()
        serviceStatus.dwCurrentState = serviceState

        ' Estimated time in milliseconds
        ' So windows knows if something is failing
        ' Default 100000 milliseconds (100 seconds)
        Select Case (serviceState)
            Case Service1.ServiceState.SERVICE_CONTINUE_PENDING
                serviceStatus.dwWaitHint = My.Settings.WaitContinue
            Case Service1.ServiceState.SERVICE_PAUSE_PENDING
                serviceStatus.dwWaitHint = My.Settings.WaitPause
            Case Service1.ServiceState.SERVICE_START_PENDING
                serviceStatus.dwWaitHint = My.Settings.WaitStart
            Case Service1.ServiceState.SERVICE_STOP_PENDING
                serviceStatus.dwWaitHint = My.Settings.WaitStop
            Case Else
        End Select

        SetServiceStatus(Me.ServiceHandle, serviceStatus)
    End Sub

    Public Shadows Property CanPauseAndContinue = My.Settings.canPause
    Protected Overrides ReadOnly Property CanRaiseEvents As Boolean
        Get
            Return My.Settings.canRaiseEvents
        End Get
    End Property

    Public Enum ServiceState
        SERVICE_STOPPED = 1
        SERVICE_START_PENDING = 2
        SERVICE_STOP_PENDING = 3
        SERVICE_RUNNING = 4
        SERVICE_CONTINUE_PENDING = 5
        SERVICE_PAUSE_PENDING = 6
        SERVICE_PAUSED = 7
    End Enum
    <StructLayout(LayoutKind.Sequential)>
    Public Structure ServiceStatus
        Public dwServiceType As Long
        Public dwCurrentState As ServiceState
        Public dwControlsAccepted As Long
        Public dwWin32ExitCode As Long
        Public dwServiceSpecificExitCode As Long
        Public dwCheckPoint As Long
        Public dwWaitHint As Long
    End Structure
    Private Declare Auto Function SetServiceStatus Lib "advapi32.dll" (ByVal handle As IntPtr, ByRef serviceStatus As ServiceStatus) As Boolean

    Protected Overrides Sub OnStart(ByVal args() As String)
        SetServiceStatus(ServiceState.SERVICE_START_PENDING)

        If My.Settings.StartProgram Then
            NewProcess(Processes.pStartup, My.Settings.StartProgram, My.Settings.StartArguments).Start()
            GetProcess(Processes.pStartup).WaitForExit()
        End If

        NewProcess(Processes.pMain, My.Settings.MainProgram, My.Settings.MainArguments).Start()

        ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf CheckRunningThread))

        SetServiceStatus(ServiceState.SERVICE_RUNNING)
    End Sub
    ' If OnStart() or OnStop() is taking some time, increment dwCheckPoint, to ask for more time
    ' https://msdn.microsoft.com/en-us/library/zt39148a(v=vs.110).aspx?f=255&MSPPError=-2147217396&cs-save-lang=1&cs-lang=vb#code-snippet-8
    '
    Protected Overrides Sub OnStop()
        SetServiceStatus(ServiceState.SERVICE_STOP_PENDING)

        Me.stopping = True
        Me.stoppedEvent.WaitOne()
        If My.Settings.hasCustomStop Then
            NewProcess(Processes.pStop, My.Settings.StopProgram, My.Settings.StopArguments).Start()
            GetProcess(Processes.pStop).WaitForExit()
        Else
            If Not GetProcess(Processes.pMain).HasExited Then
                If My.Settings.MainIsCmdApplication Then
                    KillNicelyCmdProg.StopProgramByAttachingToItsConsoleAndIssuingCtrlCEvent(GetProcess(Processes.pMain))
                Else
                    If Not GetProcess(Processes.pMain).HasExited Then
                        GetProcess(Processes.pMain).Kill()
                        Thread.Sleep(100)
                    End If
                End If
                GetProcess(Processes.pMain).Close()
                GetProcess(Processes.pMain).Dispose()
            End If
        End If
        SetServiceStatus(ServiceState.SERVICE_STOPPED)
    End Sub

    Private Sub CheckRunningThread()
        Dim isExiting As Boolean = False
        Do While Not isExiting
            Thread.Sleep(My.Settings.checkisRunning)
            If Not My.Settings.canSelfCheckRunning Then
                ' Monitor that the main process is still running, otherwise exit service to signalize this to Windows
                Try
                    If GetProcess(Processes.pMain).HasExited Then
                        If GetProcess(Processes.pMain).ExitCode > 0 Then
                            isExiting = True
                            Me.ExitCode = GetProcess(Processes.pMain).ExitCode
                        End If
                    End If
                Catch ex As InvalidOperationException
                    isExiting = True
                End Try
            Else
                ' Program must return an ExitCode greater than zero to indicate, that the process not running to stop the service
                ' Program is called every 2000 milliseconds, change in config if needed
                ' Don't set it to high, or windows will not know when your programm has crashes
                NewProcess(Processes.pCheckIfRunning, My.Settings.CheckRunningProgram, My.Settings.CheckRunningArguments).Start()
                GetProcess(Processes.pCheckIfRunning).WaitForExit()
                If GetProcess(Processes.pCheckIfRunning).ExitCode > 0 Then
                    Me.ExitCode = GetProcess(Processes.pCheckIfRunning).ExitCode
                    isExiting = True
                End If
            End If

            ' If the process was shut down unexpectedly, we're going to call the onStop() Funktion to allow cleanups, pleas consider
            ' this when implementing your onStop script.
            ' If we are requested to shut down checking by the onStop Method we only set the stoppedEvent and exit
            If Me.stopping Then
                Me.stoppedEvent.Set()
                Exit Sub
            ElseIf isExiting Then
                Me.stoppedEvent.Set()
                Me.Stop()
                Exit Sub
            End If
        Loop
    End Sub

    Protected Overrides Sub OnPause()
        SetServiceStatus(ServiceState.SERVICE_PAUSE_PENDING)
        NewProcess(Processes.pPause, My.Settings.PauseProgram, My.Settings.PauseArguments).Start()
        SetServiceStatus(ServiceState.SERVICE_PAUSED)
    End Sub

    Protected Overrides Sub OnContinue()
        SetServiceStatus(ServiceState.SERVICE_CONTINUE_PENDING)
        NewProcess(Processes.pContinue, My.Settings.ContinueProgram, My.Settings.ContinueArguments).Start()
        SetServiceStatus(ServiceState.SERVICE_RUNNING)
    End Sub

    ''TODO: REDIRECT OUTPUT
    'Private Shared Sub WriteTextAsync(ByVal text As String, ByVal output As System.IO.FileStream)
    '    If String.IsNullOrWhiteSpace(text) Then
    '        Dim filename As String = My.Settings.LogPath
    '        Dim output As System.IO.FileStream = System.IO.File.Open(filename, System.IO.FileMode.Append)

    '        Dim unicodeEncoding As Text.UnicodeEncoding = New Text.UnicodeEncoding()
    '        Dim result() As Byte = unicodeEncoding.GetBytes(text + Environment.NewLine)
    '        output.Seek(0, IO.SeekOrigin.End)

    '        output.Write(result, 0, result.Length)
    '    End If
    'End Sub

End Class
