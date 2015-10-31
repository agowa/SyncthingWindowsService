Imports System.Runtime.InteropServices
Imports System.Threading

Public Class KillNicelyCmdProg
    Public Shared Sub StopProgramByAttachingToItsConsoleAndIssuingCtrlCEvent(ByVal process As Process, Optional ByVal waitForExitTimeout As Integer = 2000)
        If (Not AttachConsole(process.Id)) Then
            Exit Sub
        End If

        ' Disable Ctrl-C handling for our program
        SetConsoleCtrlHandler(Nothing, True)

        ' Sent Ctrl-C to the attached console
        GenerateConsoleCtrlEvent(CtrlTypes.ctrl_c_event, 0)

        ' Wait for the graceful end of the process.
        ' If the process will not exit in time specified by 'waitForExitTimeout', the process will be killed
        Using New Timer((Sub(dummy)
                             If Not process.HasExited Then
                                 process.Kill()
                             End If
                         End Sub), Nothing, waitForExitTimeout, Timeout.Infinite)
            ' Must wait here. If we don't wait and re-enable Ctrl-C handling below too fast, we might terminate ourselves.
            process.WaitForExit()
        End Using

        FreeConsole()

        ' Re-enable Ctrl-C handling or any subsequently started programs will inherit the disabled state.
        SetConsoleCtrlHandler(Nothing, False)
    End Sub

#Region "DllImports"

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function AttachConsole(ByVal dwProcessId As UInt32) As Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True, ExactSpelling:=True)>
    Private Shared Function FreeConsole() As Boolean
    End Function

    Private Declare Function SetConsoleCtrlHandler Lib "kernel32" (HandlerRoutine As ConsoleCtrlDelegate, Add As Boolean) As Boolean

    Private Delegate Function ConsoleCtrlDelegate(CtrlType As CtrlTypes) As Boolean

    Private Enum CtrlTypes : uint
        ctrl_c_event = 0
        CTRL_BREAK_EVENT
        CTRL_CLOSE_EVENT
        CTRL_LOGOFF_EVENT = 5
        CTRL_SHUTDOWN_EVENT
    End Enum

    <DllImport("kernel32.dll")>
    Private Shared Function GenerateConsoleCtrlEvent(dwCtrlEvent As CtrlTypes, dwProcessGroupId As UInt32) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

#End Region

End Class