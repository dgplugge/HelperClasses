' Name:     Logger Module   
' Author:   Donald G Plugge
' Date:     2/28/08
' Purpose:  Module to facilitate event logging to a text file
' 
Imports System.IO

Public Class Logger

    ' dgp rev 2/22/08 File System
    Private mActiveFlag As Boolean = False
    Private mValidPath As String
    Private mFilePath As String
    Private mSW As System.IO.StreamWriter
    Private mPrefix As String = ""

    ' dgp rev 2/29/08 complete path
    Public ReadOnly Property FilePath() As String
        Get
            Return mFilePath
        End Get
    End Property

    ' dgp rev 2/29/08 is the log active
    Public ReadOnly Property Active() As Boolean
        Get
            Return mActiveFlag
        End Get
    End Property

    ' dgp rev 2/28/08 write logging info if logging is enabled
    Public Sub Log_Info(ByVal text As String)

        If (mActiveFlag) Then
            If (mSW.BaseStream.CanWrite) Then
                mSW.WriteLine(text)
            End If
        End If

    End Sub

    ' dgp rev 11/15/06 setup for application logging
    Private Sub Start_Logging(ByVal Prefix As String)

        Dim test_path As String = Utility.MyAppPath

        If (Utility.Create_Tree(test_path)) Then
            mValidPath = test_path
        Else
            mValidPath = CurDir()
        End If

        mFilePath = System.IO.Path.Combine(mValidPath, Prefix + Format(Now(), "yyMMddhhmmss") + ".log")

        Try
            mSW = New StreamWriter(mFilePath, True)
            mActiveFlag = True
            Log_Info(Format(Now(), "MMM dd yyyy hh mm"))
        Catch ex As Exception
            mActiveFlag = False
        End Try



    End Sub

    ' dgp rev 11/13/06 clean up before exiting
    Public Sub Clean_Up()

        ' close the log file
        mSW.Close()
        mActiveFlag = False

    End Sub

    ' dgp rev 2/29/08 Flush the writer
    Public Sub Flush()

        ' close the log file
        mSW.Flush()

    End Sub

    ' dgp rev 7/24/08 Start logging
    Public Sub New(ByVal prefix As String)

        Start_Logging(prefix)

    End Sub

End Class