Imports System
Imports System.IO
Imports System.Data.OracleClient
Imports System.Web.Mail
Imports System.Collections.Specialized
Module Module1
    Dim recipientsfile, reportsfolder, sastablesfolder As String
    Dim appEvents As New EventLog
    Sub Main()
        If ConfigOK() Then
            Try
                Dim c As New Comparison(GetCurrentState(), GetLastState())
                If c.NoChange Then
                    SendMail(getRecipients, "The Oracle tables did not change.", c.LongChangesReport)
                Else
                    Dim workingdirectory As String = Directory.CreateDirectory(reportsfolder & Now.ToString("yyyy-MM-dd-hhmmss")).FullName & "\"
                    Dim attachments As String = workingdirectory & "Changes_Report.txt," & workingdirectory & "Convert_Job.log"

                    ConvertChangedTables(workingdirectory, c.ChangedTables)

                    WriteChangesReport(workingdirectory, c)

                    WriteCurrentStateFile(c.Today)

                    AppendToLog(c.ShortChangesReport.Replace(vbCrLf, " "))

                    SendMail(getRecipients, "The Oracle tables changed", _
                            "Hello, the Oracle tables changed.  Attached you will find a report of the change(s) and " & _
                            "a log of the job that converted the tables to SAS datasets." & _
                            vbCrLf & vbCrLf & "Change(s):" & vbCrLf & vbCrLf & _
                            c.ShortChangesReport & vbCrLf & vbCrLf & "federico", attachments)
                End If
            Catch ex As Exception
                Select Case ex.GetType.ToString
                    Case "System.NullReferenceException"
                        Exit Sub
                    Case "System.Security.SecurityException"
                        Console.Write("Check .NET security: KB320268")
                        Exit Sub
                    Case Else
                        Console.Write(ex.GetType.ToString)
                        Exit Sub
                End Select
            End Try
        Else
            SendMail("user@domain.edu", "Error in tables config file.", "", "")
        End If
    End Sub
    Function ConfigOK() As Boolean
        ConfigOK = False
        Dim ConfigFile As String = "tables.config"
        If Not File.Exists(ConfigFile) Then
            appEvents.WriteEntry("GetFiles", "Configuration file not found at: " & ConfigFile, _
                                    EventLogEntryType.Error)
        Else
            Dim sfile As New StreamReader(ConfigFile)
            Dim line As String
            Dim value As String = ""
            Dim param As String = ""
            Do
                line = sfile.ReadLine()
                If line <> "" Then
                    If Not line.StartsWith("#") Then
                        If InStr(line, "=") > 0 Then
                            param = LCase(Trim(Left(line, line.IndexOf("="))))
                            value = Trim(Right(line, line.Length - line.IndexOf("=") - 1))
                            Select Case param
                                Case "recipients"
                                    recipientsfile = value
                                Case "reports"
                                    reportsfolder = value
                                Case "sastables"
                                    sastablesfolder = value
                                Case Else
                                    appEvents.WriteEntry("GetFiles", "Unrecognized parameter: " & _
                                                        param, EventLogEntryType.Warning)
                            End Select
                        End If
                    End If
                End If
            Loop Until line Is Nothing
            sfile.Close()
            If Not File.Exists(recipientsfile) Then
                appEvents.WriteEntry("tables", "Recipients file was missing in the config file.", _
                                        EventLogEntryType.Error)
                Exit Function
            ElseIf Not Directory.Exists(reportsfolder) Then
                appEvents.WriteEntry("tables", "Reports folder was missing in the config file.", _
                                        EventLogEntryType.Error)
                Exit Function
            ElseIf Not Directory.Exists(sastablesfolder) Then
                appEvents.WriteEntry("tables", "SAS tables folder was missing in the config file.", _
                                        EventLogEntryType.Error)
                Exit Function
            End If
            ConfigOK = True
        End If
    End Function
    Function GetCurrentState() As String
        GetCurrentState = ""
        Try
            Dim conn As New OracleConnection("User ID=xxxx;Data Source=xxxx;Password=xxxx")
            Dim cmd As New OracleCommand("select table_name from user_tables", conn)
            conn.Open()
            Dim tablesRdr As OracleDataReader = cmd.ExecuteReader()
            While tablesRdr.Read()
                Dim atable As String = tablesRdr.GetString(0)
                Dim cntCmd As New OracleCommand("select count(*) from " & atable, conn)
                Dim cntRdr As OracleDataReader = cntCmd.ExecuteReader
                cntRdr.Read()
                GetCurrentState += atable & "," & cntRdr.GetDecimal(0) & ControlChars.NewLine
                cntRdr.Close()
            End While
            tablesRdr.Close()
            conn.Close()
        Catch e As OracleException
            Console.Write("Oracle Error:" & vbLf & e.Message)
        End Try
    End Function
    Function GetLastState() As String
        GetLastState = ""
        Try
            Dim sr As StreamReader = File.OpenText("state\Current.State.txt")
            Do While sr.Peek() >= 0
                GetLastState += sr.ReadLine & vbCrLf
            Loop
            sr.Close()
        Catch e As Exception
            Console.WriteLine("Read Current.State.txt failed: {0}", e.ToString())
        End Try
    End Function
    Function getRecipients() As String
        Try
            Dim sr As StreamReader = File.OpenText(recipientsfile)
            Do While sr.Peek() >= 0
                Dim tmp As String = sr.ReadLine
                If tmp <> "" Then
                    If (tmp.Chars(0) <> "#") And (InStr(tmp, "@")) Then
                        getRecipients += tmp & ";"
                    End If
                End If
            Loop
            sr.Close()
            getRecipients = getRecipients.Remove(getRecipients.Length - 1, 1)
        Catch e As Exception
            Console.WriteLine("The process failed: {0}", e.ToString())
        End Try
    End Function
    Sub WriteCurrentStateFile(ByVal contents As String)
        Try
            Dim currFile As String = "state\Current.State.txt"
            Dim oldFile As String = "state\State.Created." & _
                            File.GetCreationTime(currFile).ToString("yyyy-MM-dd-hhmmss") & ".txt"
            File.Move(currFile, oldFile)
            Dim sw As StreamWriter = File.CreateText(currFile)
            sw.Write(contents)
            sw.Flush()
            sw.Close()
            File.SetCreationTime(currFile, Now)
        Catch e As Exception
            Console.WriteLine("Write new Current State failed: {0}", e.ToString())
        End Try
    End Sub
    Sub WriteChangesReport(ByVal rpath As String, ByVal rpts As Comparison)
        Try
            Dim rpt As String = ""
            rpt += "Changes Report for: " & Now.ToString("MM-dd-yyyy") & vbCrLf & vbCrLf
            rpt += "Brief Summary:" & vbCrLf & vbCrLf & rpts.ShortChangesReport & vbCrLf & vbCrLf
            rpt += "Long Comparison:" & vbCrLf & vbCrLf & rpts.LongChangesReport
            Dim sw As StreamWriter = File.CreateText(rpath & "Changes_Report.txt")
            sw.Write(rpt)
            sw.Flush()
            sw.Close()
        Catch ex As Exception
            Console.WriteLine("Write Changes report failed: {0}", ex.ToString())
        End Try
    End Sub
    Sub ConvertChangedTables(ByVal rpath As String, ByVal changedtables As StringCollection)
        Dim cmds As String = "libname xxxx oracle user=xxxx password=xxxx path=xxxx;" & vbCrLf &
                            "libname saslib """ & sastablesfolder & """;" & vbCrLf
        For Each changedtable As String In changedtables
            cmds += "data saslib." & changedtable & "; set uc02." & changedtable & "; run;" & vbCrLf
        Next
        Try
            Dim sw1 As StreamWriter = File.CreateText(rpath & "Convert_Job.sas")
            With sw1
                .Write(cmds)
                .Flush()
                .Close()
            End With
            Dim proc As New Process
            With proc.StartInfo
                .FileName = "sas"
                .WorkingDirectory = rpath
                .UseShellExecute = False
                .CreateNoWindow = True
                .Arguments = "Convert_Job.sas"
            End With
            Try
                proc.Start()
                proc.WaitForExit()
            Catch ex As Exception
                Console.WriteLine("There was an error running a sas job: " & ex.Message)
            End Try
        Catch ex As Exception
            Console.WriteLine("Convert tables failed: {0}", ex.ToString())
        End Try
    End Sub
    Sub AppendToLog(ByVal entry As String)
        Try
            Dim w As StreamWriter = File.AppendText(reportsfolder & "Tables.Change.Log.txt")
            w.WriteLine(Now.ToString("yyyy-MM-dd") & vbTab & entry)
            w.Flush()
            w.Close()
        Catch ex As Exception
            Console.WriteLine("Append to log failed: {0}", ex.ToString())
        End Try
    End Sub
    Sub SendMail(ByVal sTo As String, ByVal sSbj As String, ByVal sBody As String, Optional ByVal sAttach As String = "")
        Try
            Dim Message As MailMessage = New MailMessage
            Message.To = sTo
            Message.From = "user@domain.org"
            Message.Subject = sSbj
            Message.Body = sBody
            If sAttach <> "" Then
                Dim delim As Char = ","
                Dim sSubstr As String
                For Each sSubstr In sAttach.Split(delim)
                    If File.Exists(sSubstr) Then
                        Dim myAttachment As MailAttachment = New MailAttachment(sSubstr)
                        Message.Attachments.Add(myAttachment)
                    End If
                Next
            End If
            Try
                SmtpMail.SmtpServer = "smtphost"
                SmtpMail.Send(Message)
            Catch ehttp As System.Web.HttpException
                Console.WriteLine("0", ehttp.Message)
                Console.WriteLine("Here is the full error message")
                Console.Write("0", ehttp.ToString())
            End Try
        Catch e As System.Exception
            Console.WriteLine("Unknown Exception occurred 0", e.Message)
            Console.WriteLine("0", e.ToString())
        End Try
    End Sub
End Module
