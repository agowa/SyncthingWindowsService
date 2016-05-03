Imports System.Configuration.Install
Imports System.DirectoryServices.AccountManagement
Imports System.Security.Principal
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
        Try ' TODO: Check if this is correct.
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
                ' Check for Valid User and Password
                If checkUser.validity(My.Settings.ServiceUserAccountName, My.Settings.ServiceUserAccountPassword) Then
                    ServiceProcessInstaller1.Account = ServiceProcess.ServiceAccount.User
                    ServiceProcessInstaller1.Username = My.Settings.ServiceUserAccountName
                    If Not String.IsNullOrWhiteSpace(My.Settings.ServiceUserAccountPassword) Then
                        ServiceProcessInstaller1.Password = My.Settings.ServiceUserAccountPassword
                    End If
                Else
                    ' Falling back to NetworkService because specified credentials are invalid.
                    ServiceProcessInstaller1.Account = ServiceProcess.ServiceAccount.NetworkService
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

    Class checkUser
        Public Shared Function validity(username As String, Optional ByVal password As String = "") As Boolean
            ' Validate Password, if one of these returns true they are ok ;-)
            Dim valid As Boolean = False
            Using context As PrincipalContext = New PrincipalContext(ContextType.Domain)
                valid += context.ValidateCredentials(username, password)
            End Using
            Using context As PrincipalContext = New PrincipalContext(ContextType.ApplicationDirectory)
                valid += context.ValidateCredentials(username, password)
            End Using
            Using context As PrincipalContext = New PrincipalContext(ContextType.Machine)
                valid += context.ValidateCredentials(username, password)
            End Using
            'Return
            If valid Then
                Return True
            Else
                Return False
            End If
        End Function
    End Class

End Class
