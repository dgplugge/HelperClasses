Imports Microsoft.Win32
Imports System.IO

Public Class HelperApps

    ' dgp rev 12/3/09 CSDiff helper info
    Private Shared mCSDiff
    Private Shared mCSDiff_ExePath As String = "Software\ComponentSoftware\CSDiff"
    Private Shared mTextPad
    Private Shared mCSDiff_BaseSpec
    Private Shared mCSDiff_CompSpec

    Public Class ModDateCompare

        Implements IComparer

        ' Calls CaseInsensitiveComparer.Compare with the parameters reversed.
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
           Implements IComparer.Compare

            Return ModDateCompare(x, y)

        End Function

        Private Shared Function ModDateCompare(ByVal x As Object, ByVal y As Object) As Integer

            Dim File1 As FileInfo
            Dim File2 As FileInfo
            File1 = DirectCast(x, FileInfo)
            File2 = DirectCast(y, FileInfo)
            ModDateCompare = DateTime.Compare(File1.LastWriteTime, File2.LastWriteTime)
        End Function

    End Class



    Public Shared Function IsValidFileName(ByVal filename As String) As Boolean

        Dim i As Integer

        ' Validate
        IsValidFileName = True
        ' Can't be empty
        If String.IsNullOrEmpty(filename) Then
            IsValidFileName = False
        End If
        ' Check invalid characters
        For i = 0 To Path.GetInvalidFileNameChars.GetUpperBound(0)
            If filename.IndexOf(Path.GetInvalidFileNameChars(i)) >= 0 Then
                ' Illegal character
                IsValidFileName = False
                Exit For
            End If
        Next i

    End Function

    ' dgp rev 4/12/07 Create a Unique Name
    Public Shared Function Unique_Name() As String

        ' dgp rev 3/5/08 change the date to the 24 hour format for proper 
        ' ordering in pulldown list
        Return Format(Now(), "yyMMddHHmm")

    End Function

    ' dgp rev 12/3/09
    Public Shared Property CSDiff_BaseSpec As String
        Get
            Return mCSDiff_BaseSpec
        End Get
        Set(ByVal value As String)
            mCSDiff_BaseSpec = value
        End Set
    End Property

    ' dgp rev 12/3/09
    Public Shared Property CSDiff_CompSpec As String
        Get
            Return mCSDiff_CompSpec
        End Get
        Set(ByVal value As String)
            mCSDiff_CompSpec = value
        End Set
    End Property


    Private Shared mCSDiff_ErrFlg As Boolean = False

    ' dgp rev 9/28/07 
    Private Sub Compare_CSDiff(ByVal PathA As String, ByVal PathB As String)

        Dim exe_path As String = "Software\ComponentSoftware\CSDiff"
        Dim RegKey As RegistryKey
        RegKey = Registry.CurrentUser.OpenSubKey(exe_path, True)
        Dim CSDiff_Path As String = RegKey.GetValue("CSDiffPath")

        RegKey = Registry.CurrentUser.OpenSubKey("Software\ComponentSoftware\CSDiff\HistoryCombo", True)

        PathA = System.IO.Path.Combine(PathA, "PRO")
        PathB = System.IO.Path.Combine(PathB, "PRO")

        RegKey.SetValue("CSDiffBaseFolder", PathA)
        RegKey.SetValue("CSDiffComparedFolder", PathB)

        Process.Start(System.IO.Path.Combine(CSDiff_Path, "CSDiff.exe"))

    End Sub

    ' dgp rev 4/26/07 Get the Diff between VB6
    Public Shared Sub Compare_CSDiff(ByVal BaseSpec, ByVal CompSpec)

        mCSDiff_ErrFlg = True
        Try

            Dim exe_path As String = "Software\ComponentSoftware\CSDiff"
            Dim RegKey As RegistryKey
            RegKey = Registry.CurrentUser.OpenSubKey(exe_path, True)
            If RegKey Is Nothing Then Return

            Dim CSDiff_Path As String = RegKey.GetValue("CSDiffPath")

            RegKey = Registry.CurrentUser.OpenSubKey("Software\ComponentSoftware\CSDiff\HistoryCombo", True)
            RegKey.SetValue("CSDiffBaseFile", BaseSpec)
            RegKey.SetValue("CSDiffComparedFile", CompSpec)
            RegKey.SetValue("CSDiffBaseFolder", BaseSpec)
            RegKey.SetValue("CSDiffComparedFolder", CompSpec)

            RegKey = Registry.CurrentUser.OpenSubKey("Software\ComponentSoftware\CSDiff", True)
            RegKey.SetValue("CSDiffFileTypes", "*.pro")
            Process.Start(System.IO.Path.Combine(CSDiff_Path, "CSDiff.exe"))

        Catch ex As Exception
            mCSDiff_ErrFlg = True
            Return
        End Try

        mCSDiff_ErrFlg = False
        Return

    End Sub



End Class
