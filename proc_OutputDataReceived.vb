Imports System.Threading

Class proc_OutputDataReceived
    Private ReadOnly processNr As Service1.Processes
    Private ReadOnly filenrfmt As String = "000" 'Used to tell ToString (from Integer) to prefix numbers with zeros to a (minimum) total length of 4 digits
    Private ReadOnly fileext As String = ".log"
    Private ReadOnly maxLogSize As Integer = 10 * 1024 * 1024 'in Byte = 10 MB
    Private filename As String
    Private filenr As Integer
    Private outputFS As System.IO.FileStream

    Public Sub New(processNr As Service1.Processes)
        Me.processNr = processNr
        Me.filename = getCurrentFilename()
        Me.outputFS = System.IO.File.Open(filename, System.IO.FileMode.Append)
    End Sub

    Private Function getCurrentFilename() As String
        Return (My.Settings.LogPath + "\" + Me.processNr.ToString + "-UTC-" + DateTime.UtcNow.Year.ToString + "-" + _
                DateTime.UtcNow.Month.ToString + "-" + DateTime.UtcNow.Day.ToString + "-" + DateTime.UtcNow.Hour.ToString + _
                "-" + DateTime.UtcNow.Minute.ToString + DateTime.UtcNow.Second.ToString + DateTime.UtcNow.Millisecond.ToString + _
                "-" + Me.filenr.ToString(filenrfmt) + Me.fileext)
    End Function

    Public Sub proc_OutputDataReceived(ByVal sender As Object, ByVal e As DataReceivedEventArgs)
        If Not String.IsNullOrWhiteSpace(e.Data) Then

            Dim whileContinue As Integer = 0
            While Not whileContinue = -1
                Try
                    'Change File if it is to big
                    If My.Computer.FileSystem.GetFileInfo(filename).Length > maxLogSize Then
                        Dim test As Integer = My.Computer.FileSystem.GetFileInfo(filename).Length
                        Me.filenr += 1
                        Me.outputFS.Close()
                        Me.filename = getCurrentFilename()
                        Me.outputFS = System.IO.File.Create(filename)
                    End If
                    whileContinue = -1
                Catch ex As System.IO.IOException
                    If System.IO.File.Exists(filename) Then
                        If whileContinue > 5 Then
                            'No luck writing to file, so skip it
                            Exit Sub
                        End If
                        Thread.Sleep(300)
                    Else
                        Throw New System.IO.FileNotFoundException("File not found", filename)
                    End If
                    whileContinue += 1
                End Try
            End While

            Dim unicodeEncoding As Text.UnicodeEncoding = New Text.UnicodeEncoding()
            Dim result() As Byte = unicodeEncoding.GetBytes(e.Data + Environment.NewLine)
            outputFS.Seek(0, IO.SeekOrigin.End)
            outputFS.Write(result, 0, result.Length)
        End If
    End Sub
End Class
