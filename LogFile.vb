
Public Class LogFile
    Private Shared logPath As String = ".\"
    'Private Shared logPathClient = "DownloadLog"
    Private Shared logWriter As IO.TextWriter
    Public Shared Sub WriteLog(ByVal msg As String)
        Try
            If My.Settings.IsLogEnable = True Then
                Dim sFilename As String
                sFilename = logPath & Date.Now.ToString("yyyyMMdd") & "_logFile.log"
                logWriter = New IO.StreamWriter(sFilename, True)
                logWriter.WriteLine("[" & Now.ToString("HH:mm:ss") & "]" & ControlChars.Tab & " - " & msg & vbCrLf)
                logWriter.Close()
            End If
        Catch ex As Exception
        End Try
    End Sub



End Class
