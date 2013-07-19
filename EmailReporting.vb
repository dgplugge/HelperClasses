' dgp rev 9/26/08 Email Reporting of Errors and Feedback
Imports System.Net.Mail
Imports Scripting
'Imports Microsoft.Office.Interop

' dgp rev 9/26/08
Public Class EmailReporting

    Private EmailFlag As Boolean = True
    Private EmailSender As String

    Private mAttachFlag As Boolean
    Private mAttachError As Boolean

    Private mLogger As HelperClasses.Logger
    Public Property Logger() As HelperClasses.Logger
        Get
            Return mLogger
        End Get
        Set(ByVal value As HelperClasses.Logger)
            mLogger = value
        End Set
    End Property

    Private Sub AppendStatus(ByVal txt As String)

        If mLogger Is Nothing Then Exit Sub
        mLogger.Log_Info(txt)

    End Sub

    Private mMessage As String = ""
    Public Property Message() As String
        Get
            Return mMessage
        End Get
        Set(ByVal value As String)
            mMessage = value
        End Set
    End Property

    ' dgp rev 9/26/08 Feedback Reporting
    Public Function Scan_Files() As Boolean

        Dim item
        Dim today
        Dim fdate

        AttachList = New ArrayList
        today = Format(Now(), "yyyyMMdd")
        For Each item In System.IO.Directory.GetFiles(HelperClasses.Utility.MyAppPath)
            fdate = Format(System.IO.File.GetCreationTime(item), "yyyyMMdd")
            If (fdate = today) Then AttachList.Add(item)
        Next
        mMessage = mMessage + vbCrLf + AttachList.Count.ToString + " log file(s)"

        Scan_Files = (AttachList.Count > 0)

    End Function

    ' dgp rev 9/26/08
    Private mEmailText As String
    Public Property EmailText() As String
        Get
            Return mEmailText
        End Get
        Set(ByVal value As String)
            mEmailText = value
        End Set
    End Property

    Private objMess As MailMessage
    Private objAttach As Attachment
    Private objMail As SmtpClient

    Private mAttachList As ArrayList
    Public Property AttachList() As ArrayList
        Get
            Return mAttachList
        End Get
        Set(ByVal value As ArrayList)
            mAttachList = value
        End Set
    End Property

    Private TextFlag As Boolean = False
    Private UserName As String = Environment.GetEnvironmentVariable("Username")
    Private MachineName As String = Environment.MachineName
    Private file_list As ArrayList
    Private mExcept As Exception
    Public ReadOnly Property EmailException() As Exception
        Get
            Return mExcept
        End Get
    End Property

    ' dgp rev 10/12/07 
    Private Function PrepNewFormat(ByVal keyword) As Boolean

        EmailSender = UserName + "@mail.nih.gov"

        mAttachError = False
        mAttachFlag = False
        PrepNewFormat = True
        objMess = New MailMessage
        objMess.To.Add(New MailAddress("plugged@mail.nih.gov"))
        objMess.From = New MailAddress(EmailSender)
        objMess.Subject = "EIB: " + keyword
        objMess.Body = mEmailText + vbCrLf + "From: " + MachineName + vbCrLf

        objMail = New SmtpClient("mailfwd.nih.gov", 25)
        objMail.UseDefaultCredentials = True

        If (mAttachList Is Nothing) Then Exit Function
        mAttachFlag = True
        Dim item
        For Each item In mAttachList
            Try
                objMess.Attachments.Add(New Attachment(item))
            Catch ex As Exception
                mExcept = ex
                mAttachError = True
                Err.Clear()
            End Try
        Next

    End Function



    ' dgp rev 10/12/07 
    Private Function Prep_Message() As Boolean

        EmailSender = UserName + "@mail.nih.gov"

        mAttachError = False
        mAttachFlag = False
        Prep_Message = True
        objMess = New MailMessage
        objMess.To.Add(New MailAddress("plugge@usa.net"))
        objMess.To.Add(New MailAddress("plugged@mail.nih.gov"))
        objMess.From = New MailAddress(EmailSender)
        objMess.Subject = "Report from " + UserName
        objMess.Body = mEmailText + vbCrLf + "From: " + MachineName + vbCrLf

        objMail = New SmtpClient("mailfwd.nih.gov", 25)
        objMail.UseDefaultCredentials = True

        If (mAttachList Is Nothing) Then Exit Function
        mAttachFlag = True
        Dim item
        For Each item In mAttachList
            Try
                objMess.Attachments.Add(New Attachment(item))
            Catch ex As Exception
                mExcept = ex
                mAttachError = True
                Err.Clear()
            End Try
        Next

    End Function

    ' dgp rev 9/26/08 
    Private Function mSend()

        mSend = True
        Try
            objMail.Send(objMess)
        Catch ex As System.Net.Mail.SmtpException
            mExcept = ex
            mSend = False
        Catch ex As Exception
            mExcept = ex
            mSend = False
        End Try

    End Function

    ' dgp rev 8/2/07 Send the log files in a message
    Private Function Send_SMTP(ByVal attach As String) As Boolean

        Send_SMTP = True
        Dim VB_Flag As Boolean = True

        Dim objMess As New System.Net.Mail.MailMessage(UserName + "@mail.nih.gov", "plugged@mail.nih.gov", "Error Report", mMessage)
        Dim objMail As New System.Net.Mail.SmtpClient("mailfwd.nih.gov", 25)

        objMail.UseDefaultCredentials = True

        Dim uname As String = HelperClasses.Utility.Unique_Name()

        If (System.IO.File.Exists(attach)) Then
            Try
                ' dgp rev 7/27/07 log is open, so copy log to another file, then mail
                Dim mail_file As String = System.IO.Path.Combine(System.IO.Directory.GetParent(attach).Name, "pvwave" + uname + ".txt")
                System.IO.File.Copy(attach, mail_file, True)

                Dim attached As New System.Net.Mail.Attachment(mail_file)
                mMessage = mMessage + vbCrLf + "Attachment PVW Added"
                objMess.Attachments.Add(attached)
            Catch ex As Exception
                Send_SMTP = False
                mMessage = mMessage + vbCrLf + ex.Message
            End Try
        End If

        If (Logger.Active) Then
            Try
                ' dgp rev 7/27/07 log is open, so copy log to another file, then mail
                Logger.Flush()
                Dim fil
                fil = System.IO.Path.GetFileName(Logger.FilePath)
                Dim mail_file As String = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Logger.FilePath), "vbnet" + uname + ".txt")
                fil.Copy(mail_file, True)

                Dim attached As New System.Net.Mail.Attachment(mail_file)
                mMessage = mMessage + vbCrLf + "Attachment VB.Net Added"
                objMess.Attachments.Add(attached)
                VB_Flag = True
            Catch ex As Exception
                mMessage = mMessage + vbCrLf + ex.Message
                Send_SMTP = False
                VB_Flag = False
            End Try
        End If

        Try

            objMail.Send(objMess)
            mMessage = mMessage + vbCrLf + "Message Sent"

        Catch ex As System.Net.Mail.SmtpException

            Send_SMTP = False
            mMessage = mMessage + vbCrLf + "SMTP Error: " + ex.Message

        Catch ex As Exception

            Send_SMTP = False
            mMessage = mMessage + vbCrLf + ex.Message

        End Try

    End Function

    Private mSubject As String = ""

    ' dgp rev 8/2/07 Send the log files in a message
    Public Function SendAttach(ByVal file As String) As Boolean

        EmailSender = UserName + "@mail.nih.gov"

        mAttachError = False
        mAttachFlag = False
        objMess = New MailMessage
        objMess.To.Add(New MailAddress("plugge@usa.net"))
        objMess.To.Add(New MailAddress("plugged@mail.nih.gov"))
        objMess.From = New MailAddress(EmailSender)
        objMess.Subject = "Report from " + UserName
        objMess.Body = mEmailText + vbCrLf + "From: " + MachineName + vbCrLf

        objMail = New SmtpClient("mailfwd.nih.gov", 25)
        objMail.UseDefaultCredentials = True

        If (Not System.IO.File.Exists(file)) Then Exit Function
        mAttachFlag = True
        Try
            objMess.Attachments.Add(New Attachment(file))
        Catch ex As Exception
            mExcept = ex
            mAttachError = True
            Err.Clear()
            Exit Function
        End Try

        SendAttach = mSend()

    End Function

    ' dgp rev 9/20/2011 Send Version information
    Public Function SendLog(ByVal filespec As String) As Boolean

        Dim name = System.IO.Path.GetFileName(filespec)
        EmailSender = UserName + "@mail.nih.gov"
        objMess = New MailMessage
        objMess.To.Add(New MailAddress("plugged@mail.nih.gov"))
        objMess.From = New MailAddress(EmailSender)
        objMess.Subject = String.Format("Log {0} for {1}", name, UserName)
        objMess.Attachments.Add(New Attachment(filespec))

        Dim body = Now.Date.ToLongDateString + vbCrLf + "From: " + UserName + vbCrLf
        body = body + String.Format("Log: {0}", filespec)
        objMess.Body = body

        objMail = New SmtpClient("mailfwd.nih.gov", 25)
        objMail.UseDefaultCredentials = True

        SendLog = mSend()

    End Function



    ' dgp rev 9/20/2011 Send Version information
    Public Function SendVersionLog(ByVal ver As String) As Boolean

        Dim name = System.IO.Path.GetFileName(ver)
        EmailSender = UserName + "@mail.nih.gov"
        objMess = New MailMessage
        objMess.To.Add(New MailAddress("plugged@mail.nih.gov"))
        objMess.From = New MailAddress(EmailSender)
        objMess.Subject = String.Format("Version {0} for {1}", name, UserName)

        Dim body = Now.Date.ToLongDateString + vbCrLf + "From: " + UserName + vbCrLf
        body = body + String.Format("Version: {0}", ver)
        objMess.Body = body

        objMail = New SmtpClient("mailfwd.nih.gov", 25)
        objMail.UseDefaultCredentials = True

        SendVersionLog = mSend()

    End Function

    ' dgp rev 9/26/08 
    Private Function SendNewFormat() As Boolean

        SendNewFormat = False
        Try
            objMail.Send(objMess)
            SendNewFormat = True
        Catch ex As System.Net.Mail.SmtpException
            mExcept = ex
        Catch ex As Exception
            mExcept = ex
        End Try

    End Function



    ' dgp rev 8/2/07 Send the log files in a message
    Public Function SendReport(ByVal keyword) As Boolean

        PrepNewFormat(keyword)

        SendReport = mSend()

    End Function



    ' dgp rev 8/2/07 Send the log files in a message
    Public Function SendReport() As Boolean

        Prep_Message()

        SendReport = mSend()

    End Function

End Class
