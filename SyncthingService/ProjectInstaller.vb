Imports System.ComponentModel
Imports System.Configuration.Install
Imports System.Threading

Public Class ProjectInstaller
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

    Public Sub New()
        InitializeComponent()
        Debugging()
    End Sub

    Private Sub ProjectInstaller_BeforeInstall(sender As Object, e As InstallEventArgs) Handles Me.BeforeInstall
        Try
            Dim dependsOn(My.Settings.ServiceDependsOn.Count - 1) As String
            My.Settings.ServiceDependsOn.CopyTo(dependsOn, 0)
            ServiceInstaller1.ServicesDependedOn = dependsOn
        Catch ex As NullReferenceException

        End Try


        ServiceInstaller1.DisplayName = My.Settings.ServiceDisplayName
        ServiceInstaller1.ServiceName = My.Settings.ServiceName
        ServiceInstaller1.Description = My.Settings.ServiceDescription

        'Username can be LocalService, NetworkService, LocalSystem or a valid username with password
        Select Case My.Settings.ServiceUserAccountName.ToLower
            Case "LocalService".ToLower
                ServiceProcessInstaller1.Account = ServiceProcess.ServiceAccount.LocalService
            Case "NetworkService".ToLower
                ServiceProcessInstaller1.Account = ServiceProcess.ServiceAccount.NetworkService
            Case "LocalSystem".ToLower
                ServiceProcessInstaller1.Account = ServiceProcess.ServiceAccount.LocalSystem
            Case Else
                ServiceProcessInstaller1.Account = ServiceProcess.ServiceAccount.User
                ' TODO: Check for Valid User
                ServiceProcessInstaller1.Username = My.Settings.ServiceUserAccountName
                If Not String.IsNullOrWhiteSpace(My.Settings.ServiceUserAccountPassword) Then
                    ServiceProcessInstaller1.Password = My.Settings.ServiceUserAccountPassword
                End If
        End Select

        'StartType can be Automatic, Manual, Disabled
        Select Case My.Settings.ServiceStartType.ToLower
            Case "Automatic".ToLower
                ServiceInstaller1.StartType = ServiceProcess.ServiceStartMode.Automatic
            Case "Manual".ToLower
                ServiceInstaller1.StartType = ServiceProcess.ServiceStartMode.Manual
            Case Else
                ServiceInstaller1.StartType = ServiceProcess.ServiceStartMode.Disabled
        End Select

    End Sub

End Class
