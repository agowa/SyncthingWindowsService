Imports System.ServiceProcess
Imports System.IO
Imports System.IO.Compression
Imports System.Threading

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Service1
    Inherits System.ServiceProcess.ServiceBase

    'UserService überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    ' Der Haupteinstiegspunkt für den Prozess
    '<MTAThread()>
    '<System.Diagnostics.DebuggerNonUserCode()>
    Shared Sub Main()
        If Not isServiceInstalled() Then
            installService()
        End If

        Dim ServicesToRun() As System.ServiceProcess.ServiceBase
        ServicesToRun = New System.ServiceProcess.ServiceBase() {New Service1}
        System.ServiceProcess.ServiceBase.Run(ServicesToRun)
    End Sub
    Private Shared Function isServiceInstalled() As Boolean
        Dim myService As String = My.Settings.ServiceName
        For Each sc As ServiceController In ServiceController.GetServices()
            If (sc.ServiceName = myService) Then
                Return True
            End If
        Next
        Return False
    End Function
    Private Shared Sub installService()
        Dim path As String = System.Reflection.Assembly.GetExecutingAssembly().Location

        If isElevated() Then
            ' Is in Programs Directory? If not, move there.
            ' TODO: Respect Install dir if a install wizard was used.
            If path.StartsWith(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)) Then
                ' Is Syncthing present?
                Dim SyncthingPath As String = IO.Path.Combine(IO.Path.GetDirectoryName(path), My.Settings.MainProgram)
                If Not FileIO.FileSystem.FileExists(SyncthingPath) Then
                    getNewSyncthing(SyncthingPath + ".zip")
                    getExeFromZip(SyncthingPath + ".zip", SyncthingPath)
                End If
                ' Install Service
                System.Configuration.Install.ManagedInstallerClass.InstallHelper(New String() {path})
                ' Start Service
                Dim service As ServiceController = New ServiceController(My.Settings.ServiceName)
                service.Start()
                Thread.Sleep(10000)
                ' Open Browser with Web-UI
                Process.Start("http://127.0.0.1:8384/")
            Else
                Dim programDir As String = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), My.Settings.ServiceDisplayName)
                Directory.CreateDirectory(programDir)
                Dim destinationpath As String = IO.Path.Combine(programDir, IO.Path.GetFileName(path))
                File.Copy(path, destinationpath)
                Try
                    File.Copy(path + ".config", destinationpath + ".config")
                Catch ex As FileNotFoundException ' No Config File found, so nothing to copy.
                End Try
                Dim proc As New ProcessStartInfo
                proc.UseShellExecute = True
                proc.WorkingDirectory = programDir
                proc.FileName = destinationpath
                proc.Verb = "runas"
                Process.Start(proc)
            End If
        Else
            ' Relaunch as administrator
            Dim proc As New ProcessStartInfo
            proc.UseShellExecute = True
            proc.WorkingDirectory = Environment.CurrentDirectory
            proc.FileName = path
            proc.Verb = "runas"
            Try
                Process.Start(proc)
            Catch
                ' The user refused to allow privileges elevation.
                ' Do nothing and return directly ...
            End Try
        End If
        System.Diagnostics.Process.GetCurrentProcess.Kill()
    End Sub
    Private Shared Sub getNewSyncthing(ByVal path)
        ' Download from Github latest Release and replace current one
        ' Get the File URL first
        Dim jsonURL As String = "https://api.github.com/repos/syncthing/syncthing/releases/latest"
        Dim is64OS As Boolean = Environment.Is64BitOperatingSystem
        Dim fileURL As String = ""
        Dim jsonWebClient As Net.WebClient = New Net.WebClient()
        jsonWebClient.Headers.Add("user-agent", "Dummy User-Agent")
        Dim tempFile As String = jsonWebClient.DownloadString(jsonURL)

        For Each line In tempFile.Split(",")
            If line.Contains("browser_download_url") And line.Contains("syncthing-windows-386") And Not is64OS Then
                Dim tempArray As String() = line.Split("""") ' Split by Double Quote ' " '
                For Each e As String In tempArray
                    If e.StartsWith("http") Or e.StartsWith("ftp") Then
                        fileURL = e
                    End If
                Next
            ElseIf line.Contains("browser_download_url") And line.Contains("syncthing-windows-amd64") And is64OS Then
                Dim tempArray As String() = line.Split("""") ' Split by Double Quote ' " '
                For Each e As String In tempArray
                    If e.StartsWith("http") Or e.StartsWith("ftp") Then
                        fileURL = e
                    End If
                Next
            End If
        Next
        'My.Computer.Network.DownloadFile(fileURL, path)
        Dim syncthingWebClient As Net.WebClient = New Net.WebClient
        syncthingWebClient.DownloadFile(fileURL, path)
    End Sub
    Private Shared Function isElevated() As Boolean
        Dim principal As New Security.Principal.WindowsPrincipal(Security.Principal.WindowsIdentity.GetCurrent)
        Return principal.IsInRole(Security.Principal.WindowsBuiltInRole.Administrator)
    End Function
    Private Shared Sub getExeFromZip(ByVal zipPath As String, ByVal exePath As String)
        Dim extractPath As String = Path.Combine(Path.GetTempPath, Guid.NewGuid.ToString)
        Do While Directory.Exists(extractPath) Or File.Exists(extractPath)
            extractPath = Path.Combine(Path.GetTempPath, Guid.NewGuid.ToString)
        Loop
        ZipFile.ExtractToDirectory(zipPath, extractPath)
        Dim unzipedExePath As String() = My.Computer.FileSystem.GetFiles(extractPath, FileIO.SearchOption.SearchAllSubDirectories, "syncthing.exe").ToArray
        If (unzipedExePath.Length = 1) Then
            File.Move(unzipedExePath(0), exePath)
            Dim fs = File.GetAccessControl(exePath)
            fs.SetAccessRuleProtection(False, False) ' Inherit new parent folder permissions
            File.SetAccessControl(exePath, fs)
        Else
            Throw New FileNotFoundException("Could not find ""syncthing.exe"", may the zip was not downloaded successfully or its structure changed.")
        End If
        Try
            Directory.Delete(extractPath, True)
            File.Delete(zipPath)
        Catch ex As Exception
            ' Cleaning up Tempfiles faild, just ignore it, nobody will notice, ...
        End Try
    End Sub

    ' Wird vom Komponenten-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    ' Hinweis: Die folgende Prozedur ist für den Komponenten-Designer erforderlich.
    ' Das Bearbeiten ist mit dem Komponenten-Designer möglich.  
    ' Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        components = New System.ComponentModel.Container()
        Me.ServiceName = My.Settings.ServiceDisplayName
    End Sub

End Class
