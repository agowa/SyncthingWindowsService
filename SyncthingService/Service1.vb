Imports System.Net
Imports System.Text
Imports System.Threading

Public Class Service1 'extends System.ServiceProcess.ServiceBase
    Private udpServer As New Sockets.UdpClient(My.Settings.udpListeningPort)
    Private p As Process
    Private pStartInfo As ProcessStartInfo
    Private stopping As Boolean = False
    Private stoppedEvent As ManualResetEvent = New ManualResetEvent(False)
    Private isUdpClientConnected As Boolean = False
    Private serverThread As Thread

    Public Sub New()
        InitializeComponent()
        Debugging()
        serverThread = New Threading.Thread(New Threading.ThreadStart(AddressOf Server))
        serverThread.Start()
    End Sub

    Public Sub Server()
        Dim RemoteIpEndPoint As New IPEndPoint(My.Settings.udpListeningIP, 0)
        Dim receiveBytes As Byte()
        receiveBytes = udpServer.Receive(RemoteIpEndPoint)
        udpServer.Connect(RemoteIpEndPoint)
        Me.isUdpClientConnected = True
    End Sub

    Private Sub Debugging()
        If My.Settings.WaitForDebugger Then
            If My.Settings.DebuggerIsRemote Then
                'Debugging Remotely
                While Not Debugger.IsAttached
                    Thread.Sleep(20)
                End While
            Else
                'Debugging Localy
                Debugger.Launch()
            End If
        End If
    End Sub

    Public Shadows Property CanPauseAndContinue As Boolean = True
    Protected Overrides ReadOnly Property CanRaiseEvents As Boolean
        Get
            Return My.Settings.canRaiseEvents
        End Get
    End Property

    Protected Overrides Sub OnStart(ByVal args() As String)
        p = New Process()
        pStartInfo = New ProcessStartInfo()
        Dim path As String = Reflection.Assembly.GetExecutingAssembly().Location
        pStartInfo.WorkingDirectory = IO.Path.GetDirectoryName(path)
        pStartInfo.FileName = IO.Path.Combine(IO.Path.GetDirectoryName(path), My.Settings.MainProgram)
        pStartInfo.WindowStyle = ProcessWindowStyle.Hidden
        pStartInfo.Arguments = My.Settings.MainArguments
        pStartInfo.RedirectStandardError = True
        pStartInfo.RedirectStandardOutput = True
        pStartInfo.RedirectStandardInput = False
        pStartInfo.UseShellExecute = False
        pStartInfo.CreateNoWindow = True
        p.EnableRaisingEvents = True
        p.StartInfo = pStartInfo

        'Redirect CMD Output
        AddHandler p.ErrorDataReceived, AddressOf proc_OutputDataReceived
        AddHandler p.OutputDataReceived, AddressOf proc_OutputDataReceived

        p.Start()
        p.BeginErrorReadLine()
        p.BeginOutputReadLine()

        ThreadPool.QueueUserWorkItem(New WaitCallback(AddressOf CheckRunningThread))
    End Sub


    Delegate Sub UpdateTextBoxDelg(text As String)
    Public myDelegate As UpdateTextBoxDelg = New UpdateTextBoxDelg(AddressOf udpSendOutputData)
    Public Sub udpSendOutputData(text As String)
        Dim senddata As Byte()
        senddata = System.Text.Encoding.ASCII.GetBytes(text & Environment.NewLine)
        Try
            If isUdpClientConnected Then
                udpServer.Send(senddata, senddata.Length)
            End If
        Catch ex As System.InvalidOperationException

        End Try

    End Sub
    Public Sub proc_OutputDataReceived(ByVal sender As Object, ByVal e As DataReceivedEventArgs)
        udpSendOutputData(e.Data)
    End Sub


    ' If OnStart() or OnStop() is taking some time, increment dwCheckPoint, to ask for more time
    ' https://msdn.microsoft.com/en-us/library/zt39148a(v=vs.110).aspx?f=255&MSPPError=-2147217396&cs-save-lang=1&cs-lang=vb#code-snippet-8
    '
    Protected Overrides Sub OnStop()
        Me.stopping = True
        Me.stoppedEvent.WaitOne()
        If Not p.HasExited Then
            KillNicelyCmdProg.StopProgramByAttachingToItsConsoleAndIssuingCtrlCEvent(p)
            p.Close()
            p.Dispose()
        End If
        If serverThread.ThreadState = Threading.ThreadState.Running Then
            serverThread.Abort()
        End If
        udpServer.Close()
    End Sub

    <Flags()> Public Enum ThreadAccess As Integer
        TERMINATE = (&H1)
        SUSPEND_RESUME = (&H2)
        GET_CONTEXT = (&H8)
        SET_CONTEXT = (&H10)
        SET_INFORMATION = (&H20)
        QUERY_INFORMATION = (&H40)
        SET_THREAD_TOKEN = (&H80)
        IMPERSONATE = (&H100)
        DIRECT_IMPERSONATION = (&H200)
    End Enum
    Private Declare Function OpenThread Lib "kernel32.dll" (ByVal dwDesiredAccess As ThreadAccess, ByVal bInheritHandle As Boolean, ByVal dwThreadId As UInteger) As IntPtr
    Private Declare Function SuspendThread Lib "kernel32.dll" (ByVal hThread As IntPtr) As UInteger
    Private Declare Function ResumeThread Lib "kernel32.dll" (ByVal hThread As IntPtr) As UInteger
    Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hThread As IntPtr) As Boolean
    Protected Overrides Sub OnPause()
        For Each thread As ProcessThread In p.Threads
            Dim th As IntPtr
            th = OpenThread(ThreadAccess.SUSPEND_RESUME, False, thread.Id)
            If th <> IntPtr.Zero Then
                SuspendThread(th)
                CloseHandle(th)
            End If
        Next
    End Sub
    Protected Overrides Sub OnContinue()
        For Each thread As ProcessThread In p.Threads
            Dim th As IntPtr
            th = OpenThread(ThreadAccess.SUSPEND_RESUME, False, thread.Id)
            If th <> IntPtr.Zero Then
                ResumeThread(th)
                CloseHandle(th)
            End If
        Next
    End Sub

    Private Sub CheckRunningThread()
        Dim isExiting As Boolean = False
        Do While Not isExiting
            Thread.Sleep(My.Settings.checkisRunning)
            ' Monitor that the main process is still running, otherwise exit service to signalize this to Windows
            Try
                If p.HasExited Then
                    isExiting = True
                    Me.ExitCode = p.ExitCode
                End If
            Catch ex As InvalidOperationException
                isExiting = True
            End Try
            '' TODO: Check if service is working
            'NewProcess(Processes.pCheckIfRunning, My.Settings.CheckRunningProgram, My.Settings.CheckRunningArguments).Start()
            '    GetProcess(Processes.pCheckIfRunning).WaitForExit()
            'If GetProcess(Processes.pCheckIfRunning).ExitCode > 0 Then
            '    Me.ExitCode = GetProcess(Processes.pCheckIfRunning).ExitCode
            '    isExiting = True
            'End If

            ' If the process was shut down unexpectedly, we're going to call the onStop() Funktion to allow cleanups
            ' If we are already shutting down we only have to set the stoppedEvent (onStop is waiting for us) and exit
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

End Class
