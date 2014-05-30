Imports Microsoft.Win32
Imports System.Text
Imports System.IO
Public Class EmbeddedInstall
#Region "Internal variables"

    'Variables for setup.exe command line
    Private m_instanceName As String = "SQLEXPRESS"
    Private installSqlDir As String = ""
    Private installSqlSharedDir As String = ""
    Private installSqlDataDir As String = ""
    Private addLocal As String = "All"
    Private sqlAutoStart As Boolean = True
    Private sqlBrowserAutoStart As Boolean = True
    Private sqlBrowserAccount As String = ""
    Private m_sqlBrowserPassword As String = ""
    Private sqlAccount As String = ""
    Private sqlPassword As String = ""
    Private sqlSecurityMode As Boolean = False
    Private saPassword As String = "abc@123"
    Private sqlCollation As String = ""
    Private m_disableNetworkProtocols As Boolean = False
    Private errorReporting As Boolean = True
    Private sqlExpressSetupFileLocation As String = System.Environment.GetEnvironmentVariable("TEMP") + "\sqlexpr.exe"
    Private extractorSetupFileLocation As String = ".\"
    Private m_b64bitOsDetected As Boolean = False
    Private _zipExtractFilePath As String
#End Region
#Region "Properties"
    Public Property InstanceName() As String
        Get
            Return m_instanceName
        End Get
        Set(ByVal value As String)
            m_instanceName = value
        End Set
    End Property

    Public Property SetupFileLocation() As String
        Get

            Return sqlExpressSetupFileLocation
        End Get
        Set(ByVal value As String)
            sqlExpressSetupFileLocation = value
        End Set
    End Property

    Public Property SqlInstallSharedDirectory() As String
        Get
            Return installSqlSharedDir
        End Get
        Set(ByVal value As String)
            installSqlSharedDir = value
        End Set
    End Property
    Public Property SqlDataDirectory() As String
        Get
            Return installSqlDataDir
        End Get
        Set(ByVal value As String)
            installSqlDataDir = value
        End Set
    End Property
    Public Property AutostartSQLService() As Boolean
        Get
            Return sqlAutoStart
        End Get
        Set(ByVal value As Boolean)
            sqlAutoStart = value
        End Set
    End Property
    Public Property AutostartSQLBrowserService() As Boolean
        Get
            Return sqlBrowserAutoStart
        End Get
        Set(ByVal value As Boolean)
            sqlBrowserAutoStart = value
        End Set
    End Property
    Public Property SqlBrowserAccountName() As String
        Get
            Return sqlBrowserAccount
        End Get
        Set(ByVal value As String)
            sqlBrowserAccount = value
        End Set
    End Property
    Public Property SqlBrowserPassword() As String
        Get
            Return m_sqlBrowserPassword
        End Get
        Set(ByVal value As String)
            m_sqlBrowserPassword = value
        End Set
    End Property
    'Defaults to LocalSystem
    Public Property SqlServiceAccountName() As String
        Get
            Return sqlAccount
        End Get
        Set(ByVal value As String)
            sqlAccount = value
        End Set
    End Property
    Public Property SqlServicePassword() As String
        Get
            Return sqlPassword
        End Get
        Set(ByVal value As String)
            sqlPassword = value
        End Set
    End Property
    Public Property UseSQLSecurityMode() As Boolean
        Get
            Return sqlSecurityMode
        End Get
        Set(ByVal value As Boolean)
            sqlSecurityMode = value
        End Set
    End Property
    Public WriteOnly Property SysadminPassword() As String
        Set(ByVal value As String)
            saPassword = value
        End Set
    End Property
    Public Property Collation() As String
        Get
            Return sqlCollation
        End Get
        Set(ByVal value As String)
            sqlCollation = value
        End Set
    End Property
    Public Property DisableNetworkProtocols() As Boolean
        Get
            Return m_disableNetworkProtocols
        End Get
        Set(ByVal value As Boolean)
            m_disableNetworkProtocols = value
        End Set
    End Property
    Public Property ReportErrors() As Boolean
        Get
            Return errorReporting
        End Get
        Set(ByVal value As Boolean)
            errorReporting = value
        End Set
    End Property
    Public Property SqlInstallDirectory() As String
        Get
            Return installSqlDir
        End Get
        Set(ByVal value As String)
            installSqlDir = value
        End Set
    End Property
    Public Property ZipExtractorFilePath() As String
        Get
            Return _zipExtractFilePath
        End Get
        Set(ByVal value As String)
            _zipExtractFilePath = value
        End Set
    End Property

#End Region

    Public Function IsExpressInstalled() As Boolean
        Using Key As RegistryKey = Registry.LocalMachine.OpenSubKey("Software\Microsoft\Microsoft SQL Server\", False)
            If Key Is Nothing Then
                Return False
            End If
            Dim strNames As String()
            strNames = Key.GetSubKeyNames()

            'If we cannot find a SQL Server registry key, we don't have SQL Server Express installed
            If strNames.Length = 0 Then
                Return False
            End If

            For Each s As String In strNames
                If s.StartsWith("MSSQL.") Then
                    'Check to see if the edition is "Express Edition"
                    Using KeyEdition As RegistryKey = Key.OpenSubKey(s.ToString() & "\Setup\", False)
                        If DirectCast(KeyEdition.GetValue("Edition"), String) = "Express Edition" Then
                            'If there is at least one instance of SQL Server Express installed, return true
                            Return True
                        End If
                    End Using
                End If
            Next
        End Using
        Return False
    End Function

    Public Function EnumSQLInstances(ByRef strInstanceArray As String(), ByRef strEditionArray As String(), ByRef strVersionArray As String()) As Integer
        Using Key As RegistryKey = Registry.LocalMachine.OpenSubKey("Software\Microsoft\Microsoft SQL Server\", False)
            If Key Is Nothing Then
                Return 0
            End If
            Dim strNames As String()
            strNames = Key.GetSubKeyNames()

            'If we can not find a SQL Server registry key, we return 0 for none
            If strNames.Length = 0 Then
                Return 0
            End If

            'How many instances do we have?
            Dim iNumberOfInstances As Integer = 0

            For Each s As String In strNames
                If s.StartsWith("MSSQL.") Then
                    iNumberOfInstances += 1
                End If
            Next

            'Reallocate the string arrays to the new number of instances
            strInstanceArray = New String(iNumberOfInstances - 1) {}
            strVersionArray = New String(iNumberOfInstances - 1) {}
            strEditionArray = New String(iNumberOfInstances - 1) {}
            Dim iCounter As Integer = 0

            For Each s As String In strNames
                If s.StartsWith("MSSQL.") Then
                    'Get Instance name
                    Using KeyInstanceName As RegistryKey = Key.OpenSubKey(s.ToString(), False)
                        strInstanceArray(iCounter) = DirectCast(KeyInstanceName.GetValue(""), String)
                    End Using

                    'Get Edition
                    Using KeySetup As RegistryKey = Key.OpenSubKey(s.ToString() & "\Setup\", False)
                        strEditionArray(iCounter) = DirectCast(KeySetup.GetValue("Edition"), String)
                        strVersionArray(iCounter) = DirectCast(KeySetup.GetValue("Version"), String)
                    End Using

                    iCounter += 1
                End If
            Next
            Return iCounter
        End Using
    End Function

    Private Function BuildCommandLine() As String
        Dim strCommandLine As New StringBuilder()

        If Not String.IsNullOrEmpty(installSqlDir) Then
            strCommandLine.Append(" INSTALLSQLDIR=""").Append(installSqlDir).Append("""")
        End If

        If Not String.IsNullOrEmpty(installSqlSharedDir) Then
            strCommandLine.Append(" INSTALLSQLSHAREDDIR=""").Append(installSqlSharedDir).Append("""")
        End If

        If Not String.IsNullOrEmpty(installSqlDataDir) Then
            strCommandLine.Append(" INSTALLSQLDATADIR=""").Append(installSqlDataDir).Append("""")
        End If

        If Not String.IsNullOrEmpty(addLocal) Then
            strCommandLine.Append(" ADDLOCAL=""").Append(addLocal).Append("""")
        End If

        If sqlAutoStart Then
            strCommandLine.Append(" SQLAUTOSTART=1")
        Else
            strCommandLine.Append(" SQLAUTOSTART=0")
        End If

        If sqlBrowserAutoStart Then
            strCommandLine.Append(" SQLBROWSERAUTOSTART=1")
        Else
            strCommandLine.Append(" SQLBROWSERAUTOSTART=0")
        End If

        If Not String.IsNullOrEmpty(sqlBrowserAccount) Then
            strCommandLine.Append(" SQLBROWSERACCOUNT=""").Append(sqlBrowserAccount).Append("""")
        End If

        If Not String.IsNullOrEmpty(m_sqlBrowserPassword) Then
            strCommandLine.Append(" SQLBROWSERPASSWORD=""").Append(m_sqlBrowserPassword).Append("""")
        End If

        If Not String.IsNullOrEmpty(sqlAccount) Then
            strCommandLine.Append(" SQLACCOUNT=""").Append(sqlAccount).Append("""")
        End If

        If Not String.IsNullOrEmpty(sqlPassword) Then
            strCommandLine.Append(" SQLPASSWORD=""").Append(sqlPassword).Append("""")
        End If

        strCommandLine.Append(" INSTANCENAME=" & InstanceName & "")

        If sqlSecurityMode = True Then
            strCommandLine.Append(" SECURITYMODE=SQL")
        End If

        If Not String.IsNullOrEmpty(saPassword) Then
            strCommandLine.Append(" SAPWD=""").Append(saPassword).Append("""")
        End If

        If Not String.IsNullOrEmpty(sqlCollation) Then
            strCommandLine.Append(" SQLCOLLATION=""").Append(sqlCollation).Append("""")
        End If

        If m_disableNetworkProtocols = True Then
            strCommandLine.Append(" DISABLENETWORKPROTOCOLS=1")
        Else
            strCommandLine.Append(" DISABLENETWORKPROTOCOLS=0")
        End If

        If errorReporting = True Then
            strCommandLine.Append(" ERRORREPORTING=1")
        Else
            strCommandLine.Append(" ERRORREPORTING=0")
        End If

        Return strCommandLine.ToString()
    End Function

    Public Function InstallExpress(ByVal isDisplay As Boolean, ByVal instanceName As String) As Boolean

        If extractSqlInstaller() = True Then


            'In both cases, we run Setup because we have the file.
            Dim sqlSetupProcess As New Process()
            Dim param As String = "/qn"
            If isDisplay Then
                param = "/qb"
            End If


            sqlSetupProcess.StartInfo.FileName = sqlExpressSetupFileLocation
            sqlSetupProcess.StartInfo.Arguments = param & BuildCommandLine()
            'To prevent logging of password 
            ' LogFile.WriteLog("Command is:" & sqlSetupProcess.StartInfo.FileName & " " & sqlSetupProcess.StartInfo.Arguments)
            '      /qn -- Specifies that setup run with no user interface.
            '                        /qb -- Specifies that setup show only the basic user interface. Only dialog boxes displaying progress information are displayed. Other dialog boxes, such as the dialog box that asks users if they want to restart at the end of the setup process, are not displayed.
            '                

            sqlSetupProcess.StartInfo.UseShellExecute = False

            sqlSetupProcess.Start()

            sqlSetupProcess.WaitForExit()

            Dim sqlPath As String = getSQLPath(instanceName)
            If sqlPath Is Nothing Then
                LogFile.WriteLog(instanceName & " is not installed successfully")
            Else
                LogFile.WriteLog(instanceName & " is installed successfully")

                If (m_b64bitOsDetected = True) Then
                    If (installXMO_x64() = False) Then
                        LogFile.WriteLog("Management Object Collection is not installed successfully")
                    Else
                        LogFile.WriteLog("Management Object Collection is installed successfully")
                    End If
                End If

            End If
        Else
            LogFile.WriteLog("SQL setup is not extracted")
        End If


    End Function

    Public Function installXMO_x64() As Boolean
        Try
            Dim xmoX64Process As New Process()
            Dim param As String = "/i SQLServer2005_XMO_x64.msi /qr"


            xmoX64Process.StartInfo.FileName = "msiexec"
            xmoX64Process.StartInfo.Arguments = param

            xmoX64Process.StartInfo.UseShellExecute = False

            xmoX64Process.Start()

            xmoX64Process.WaitForExit()

            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function
    Public Function extractSqlInstaller() As Boolean
        Try
            Dim extractSql As New Process()
            Dim param As String = "x -y SQLEXPR.7z"


            extractSql.StartInfo.FileName = ZipExtractorFilePath
            extractSql.StartInfo.Arguments = param

            extractSql.StartInfo.UseShellExecute = False

            extractSql.Start()

            extractSql.WaitForExit()

            Return True

        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function isDBExist(ByVal dbName As String, ByVal connectionString As String) As Boolean
        LogFile.WriteLog("Database Checking")
        Dim sqlConn As New SqlClient.SqlConnection(connectionString)
        Dim sqlCommand As SqlClient.SqlCommand
        Dim retValue As Boolean = False
        Try
            sqlCommand = New SqlClient.SqlCommand()
            sqlCommand.Connection = sqlConn
            sqlConn.Open()
            sqlCommand.CommandText = "SELECT Count(*) FROM sys.databases WHERE [name]='" & dbName & "'"
            retValue = sqlCommand.ExecuteScalar()
            sqlConn.Close()



        Catch ex As Exception

            LogFile.WriteLog("Database Checking : " & ex.Message)
        Finally
            If Not sqlConn.State = ConnectionState.Closed Then
                sqlConn.Close()
            End If
        End Try
        Return retValue
    End Function

    Public Function createDatabase(ByVal dbName As String, ByVal backupFileName As String, ByVal backupFilePath As String, ByVal sqlPath As String, ByVal instanceName As String, ByVal password As String) As Boolean

        Dim strAppPath As String
        strAppPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location())

        'Dim orgpath As String = backupFilePath & "\" & backupFileName
        'Dim destPath As String = sqlPath & "\Backup\" & backupFileName

        'My.Computer.FileSystem.CopyFile(orgpath, destPath, True)
        Dim mdfFilePath As String = sqlPath & "\Data\" & dbName & ".mdf "
        Dim ldfFilePath As String = sqlPath & "\Data\" & dbName & "_log.ldf "

        Dim installDB As Boolean = True

        Dim newConnectionString As String = "Data Source=localhost\" & instanceName & ";Initial Catalog=master;User Id=sa;password=" & password
        'connectionString = "Data Source=localhost\MyEXPRESS;Initial Catalog=master;User ID=sa;password=abc@123"

        If isDBExist(dbName, newConnectionString) Then
            Dim dialogResult As System.Windows.Forms.DialogResult = System.Windows.Forms.MessageBox.Show("Database already exists. Do you want to replace it?", "Replace Database", Windows.Forms.MessageBoxButtons.YesNo, Windows.Forms.MessageBoxIcon.Question, Windows.Forms.MessageBoxDefaultButton.Button1)

            If dialogResult = Windows.Forms.DialogResult.Yes Then
                LogFile.WriteLog("Database will be replace")
                installDB = True

            Else
                installDB = False

                'Launch DB Migration Utility
                'If Exit Code is 1007 then Update utility is failed.
                
            End If
        End If

        Do Until installDB = False

            If installDB Then
                KillAllConnection(newConnectionString, dbName)
                Dim sqlConn As New SqlClient.SqlConnection(newConnectionString)
                Dim sqlCommand As SqlClient.SqlCommand
                Try
                    sqlCommand = New SqlClient.SqlCommand()
                    sqlCommand.Connection = sqlConn
                    sqlConn.Open()

                    sqlCommand.CommandText = "USE master ALTER DATABASE " & dbName & " SET SINGLE_USER WITH ROLLBACK IMMEDIATE  RESTORE DATABASE  " & dbName & " FROM DISK='" & strAppPath & "\" & backupFilePath & "\" & backupFileName & "'  WITH REPLACE,RECOVERY,MOVE '" & dbName & "' to '" & mdfFilePath & "',move '" & dbName & "_log' to '" & ldfFilePath & "'"
                    'sqlCommand.CommandText = "RESTORE DATABASE  " & dbName & " FROM DISK='" & strAppPath & "\" & backupFilePath & "\" & backupFileName & "' With move '" & dbName & "' to '" & mdfFilePath & "',move '" & dbName & "_log' to '" & ldfFilePath & "',REPLACE "
                    sqlCommand.ExecuteNonQuery()
                    sqlConn.Close()
                    LogFile.WriteLog(dbName & " is successfully created.")
                    installDB = False

                Catch ex As SqlClient.SqlException

                    If ex.Message.Contains("RESTORE DATABASE successfully") Then
                        LogFile.WriteLog("RESTORE DATABASE successfully")
                        installDB = True
                    Else
                        LogFile.WriteLog(ex.Message & " Error number is:" & ex.Number)
                        Dim dialogueResult As System.Windows.Forms.DialogResult = System.Windows.Forms.MessageBox.Show("Failed to replace database. Close all other applications and try again." & vbCrLf & "If problem persists please contact support team", "Error in replace database", Windows.Forms.MessageBoxButtons.RetryCancel, Windows.Forms.MessageBoxIcon.Question, Windows.Forms.MessageBoxDefaultButton.Button1)
                        If dialogueResult = Windows.Forms.DialogResult.Retry Then
                            installDB = True
                        Else
                            installDB = False
                        End If
                    End If

                Catch ex As Exception
                    LogFile.WriteLog(ex.Message)
                    LogFile.WriteLog(ex.StackTrace)
                    Dim dialogueResult As System.Windows.Forms.DialogResult = System.Windows.Forms.MessageBox.Show("Failed to replace database. Close all other applications and try again." & vbCrLf & "If problem persists please contact support team", "Database Installation Error", Windows.Forms.MessageBoxButtons.RetryCancel, Windows.Forms.MessageBoxIcon.Question, Windows.Forms.MessageBoxDefaultButton.Button1)

                    If dialogueResult = Windows.Forms.DialogResult.Retry Then
                        installDB = True
                    Else
                        installDB = False
                    End If
                Finally
                    If Not sqlConn.State = ConnectionState.Closed Then
                        sqlConn.Close()
                    End If
                End Try
            End If
        Loop
    End Function
    Private Sub KillAllConnection(ByVal newConnectionString As String, ByVal dbName As String)

        Dim sqlConn As New SqlClient.SqlConnection(newConnectionString)
        Dim sqlCommand As SqlClient.SqlCommand

        sqlCommand = New SqlClient.SqlCommand()
        sqlCommand.Connection = sqlConn
        sqlConn.Open()

        sqlCommand.CommandText = "" &
"" &
"SET NOCOUNT ON " &
"DECLARE @DBName varchar(50) " &
"DECLARE @spidstr varchar(8000) " &
"DECLARE @ConnKilled smallint " &
"SET @ConnKilled=0 " &
"SET @spidstr = '' " &
"Set @DBName = 'database_new' " &
"IF db_id(@DBName) < 4 " &
"BEGIN " &
"PRINT 'Connections to system databases cannot be killed' " &
"RETURN " &
"END " &
"SELECT @spidstr=coalesce(@spidstr,',' )+'kill '+convert(varchar, spid)+ '; ' " &
"FROM master..sysprocesses WHERE dbid=db_id(@DBName) " &
"IF LEN(@spidstr) > 0 " &
"BEGIN " &
"EXEC(@spidstr) " &
"SELECT @ConnKilled = COUNT(1) " &
"FROM master..sysprocesses WHERE dbid=db_id(@DBName) " &
"END"

        sqlCommand.ExecuteNonQuery()
        sqlConn.Close()
    End Sub
  
    Public Function getSQLPath(ByVal instanceName As String) As String
        Dim Key As RegistryKey = Registry.LocalMachine.OpenSubKey("Software\Microsoft\Microsoft SQL Server\", False)
        If Key Is Nothing Then
            Return Nothing
        End If

        Dim strNames As String()
        strNames = Key.GetSubKeyNames()

        'If we can not find a SQL Server registry key, we return 0 for none
        If strNames.Length = 0 Then
            Return Nothing
        End If

        Dim instanceKey As RegistryKey = Registry.LocalMachine.OpenSubKey("Software\Microsoft\Microsoft SQL Server\" & instanceName, False)

        If instanceKey Is Nothing Then
            Return Nothing
        End If

        strNames = instanceKey.GetSubKeyNames()

        If instanceKey Is Nothing Then
            Return Nothing
        End If


        Dim strSQLPath As String = ""
        Dim instanceSql As String

        Dim setupkey As RegistryKey = Registry.LocalMachine.OpenSubKey("Software\Microsoft\Microsoft SQL Server\Instance Names\SQL", False)

        'If it is nothing then try other specially for 64 bit os.
        If (setupkey Is Nothing) Then
            setupkey = Registry.LocalMachine.OpenSubKey("Software\Wow6432Node\Microsoft\Microsoft SQL Server\Instance Names\SQL", False)
        End If

        instanceSql = DirectCast(setupkey.GetValue(instanceName), String)

        Dim instanceSqlKey As RegistryKey = Registry.LocalMachine.OpenSubKey("Software\Microsoft\Microsoft SQL Server\" & instanceSql & "\Setup", False)

        'If it is nothing then try other specially for 64 bit os.
        If (instanceSqlKey Is Nothing) Then
            instanceSqlKey = Registry.LocalMachine.OpenSubKey("Software\Wow6432Node\Microsoft\Microsoft SQL Server\" & instanceSql & "\Setup", False)
            m_b64bitOsDetected = True
        End If

        strSQLPath = DirectCast(instanceSqlKey.GetValue("SQLDataRoot"), String)

        Return strSQLPath
    End Function

End Class

Class Program

    Public Shared Sub Main(ByVal args() As String)
        Dim EI As New EmbeddedInstall()
        Dim isInstall As Boolean = True

        Dim DBNAME As String = args(0) '"database_new"
        Dim INSTANCENAME As String = args(1) '"MyEXPRESS"
        Dim BACKUPFILENAME As String = args(2) '"database_new.bak"
        Dim BACKUPFILEPATH As String = args(3) '".\"
        Dim SQLSERVERSETUPPATH As String = args(4) '".\SQLEXPR32_SP3_setup\setup.exe"
        Dim SAPWD As String = args(5) '"abc@123"
        Dim EXTRACTORPATH As String = args(6) '.\\7z.exe

        Console.WriteLine("Checking for database service please wait...")
        Dim sqlPath As String = EI.getSQLPath(INSTANCENAME)
        LogFile.WriteLog("gesqlpath " & sqlPath)
        If sqlPath Is Nothing Then
            LogFile.WriteLog("There is no SQL Server Express instances " & INSTANCENAME & " is installed." & vbLf & vbLf)
            isInstall = False
        Else
            LogFile.WriteLog("An instance " & INSTANCENAME & " of SQL Server Express is installed." & vbLf & vbLf)
            EI.createDatabase(DBNAME, BACKUPFILENAME, BACKUPFILEPATH, sqlPath, INSTANCENAME, SAPWD)
        End If

        If isInstall = False Then
            LogFile.WriteLog(vbLf & "Installing SQL Server 2005 Express Edition" & vbLf)

            EI.AutostartSQLBrowserService = True
            EI.AutostartSQLService = True
            'by default this collation will set if we write this for diffrent culture this will be diffrent
            'EI.Collation = "SQL_Latin1_General_CP1_CI_AS"

            EI.ZipExtractorFilePath = EXTRACTORPATH
            EI.DisableNetworkProtocols = False
            EI.InstanceName = INSTANCENAME
            EI.ReportErrors = True
            EI.SetupFileLocation = SQLSERVERSETUPPATH
            'Provide location for the Express setup file
            EI.SqlBrowserAccountName = ""
            'Blank means LocalSystem
            EI.SqlBrowserPassword = ""
            ' N/A
            EI.SqlDataDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) & "\Microsoft SQL Server\"
            EI.SqlInstallDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
            EI.SqlInstallSharedDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) & "\"
            'If sqlServiceAccount vl none then by default it vl install

            'For English,Simplified Chinese,Traditional Chinese,Korean,Japanese
            ' EI.SqlServiceAccountName = "NT AUTHORITY\NETWORK SERVICE"

            'For German
            'EI.SqlServiceAccountName = "NT-AUTORITÄT\NETZWERKDIENST"

            'For French
            'EI.SqlServiceAccountName = "AUTORITE NT\SERVICE RÉSEAU"

            'For Italian
            'EI.SqlServiceAccountName = "NT AUTHORITY\SERVIZIO DI RETE"

            'For Spanish
            'EI.SqlServiceAccountName = "NT AUTHORITY\SERVICIO DE RED"

            'For Russian
            'EI.SqlServiceAccountName = "NT AUTHORITY\NETWORK SERVICE"


            'Blank means Localsystem
            EI.SqlServicePassword = ""
            ' N/A
            EI.SysadminPassword = SAPWD
            '<<Supply a secure sysadmin password>>
            EI.UseSQLSecurityMode = True



            EI.InstallExpress(True, INSTANCENAME)
            sqlPath = EI.getSQLPath(INSTANCENAME)

            EI.createDatabase(DBNAME, BACKUPFILENAME, BACKUPFILEPATH, sqlPath, INSTANCENAME, SAPWD)


        End If

        'LogFile.WriteLog(vbLf & "Installing custom application" & vbLf)


        'If you need to run another MSI install, remove the following comment lines 
        'and fill in information about your MSI

        'Process myProcess = new Process();
        '            myProcess.StartInfo.FileName = "";//<<Insert the path to your MSI file here>>
        '            myProcess.StartInfo.Arguments = ""; //<<Insert any command line parameters here>>
        '            myProcess.StartInfo.UseShellExecute = false;
        '            myProcess.Start();


    End Sub

End Class



