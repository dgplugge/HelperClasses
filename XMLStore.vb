' Name:     XMLStore
' Author:   Donald G Plugge
' Date:     5/2/2012
' Purpose:  XML Store of information
Imports System.Threading
Imports Microsoft.Win32
Imports System.IO
Imports System.Xml.Linq
Imports HelperClasses
Imports System.Linq

Public Class XMLStore

    Private mMessage As String

    Private mElementName = "Settings"
    Public Property ElementName As String
        Get
            Return mElementName
        End Get
        Set(ByVal value As String)
            mElementName = value
        End Set
    End Property

    ' dgp rev 5/2/2012 
    Private Function SingleLogElement() As XElement

        Try
            SingleLogElement =
                New XElement("Item", New XAttribute("Timestamp", DateTime.Now),
                                  New XElement("Date", Now.Date.ToLongDateString),
                                  New XElement("User", Environment.UserName))

        Catch ex As Exception
            mMessage = ex.Message
            SingleLogElement =
                New XElement("Item", New XAttribute("Timestamp", DateTime.Now),
                                  New XElement("Error", ex.Message),
                                  New XElement("Date", Now.Date.ToLongDateString),
                                  New XElement("User", Environment.UserName))
        End Try

    End Function

    ' dgp rev 5/2/2012 
    Private Function SingleLogElement(NewKey As String, NewVal As String) As XElement

        Try
            SingleLogElement =
                New XElement("Item", New XAttribute("Timestamp", DateTime.Now),
                                  New XElement("Date", Now.Date.ToLongDateString),
                                  New XElement(NewKey, NewVal),
                                  New XElement("User", Environment.UserName))

            Return SingleLogElement
        Catch ex As Exception
            mMessage = ex.Message
        End Try
        Try
            SingleLogElement =
                New XElement("Item", New XAttribute("Timestamp", DateTime.Now),
                                  New XElement("Error", mMessage),
                                  New XElement("Date", Now.Date.ToLongDateString),
                                  New XElement("Key", NewKey),
                                  New XElement("Value", NewVal),
                                  New XElement("User", Environment.UserName))
            Return SingleLogElement
        Catch ex As Exception
            Return New XElement("Creation", "Error")
        End Try

    End Function

    ' dgp rev 5/2/2012 
    Private Function SingleLogElement(NewInfo As String) As XElement

        Try
            SingleLogElement =
                New XElement("Item", New XAttribute("Timestamp", DateTime.Now),
                                  New XElement("Date", Now.Date.ToLongDateString),
                                  New XElement("Info", NewInfo),
                                  New XElement("User", Environment.UserName))

        Catch ex As Exception
            mMessage = ex.Message
            SingleLogElement =
                New XElement("Item", New XAttribute("Timestamp", DateTime.Now),
                                  New XElement("Error", ex.Message),
                                  New XElement("Date", Now.Date.ToLongDateString),
                                  New XElement("Info", NewInfo),
                                  New XElement("User", Environment.UserName))
        End Try

    End Function

    ' dgp rev 5/2/2012 Path to Server
    Private mShareName As String = "Distribution"
    Public ReadOnly Property GlobalShareName As String
        Get
            Return mShareName
        End Get
    End Property

    ' dgp rev 5/2/2012 
    Private mGlobalServerPath = Nothing
    Public ReadOnly Property GlobalServerPath() As String
        Get
            If mGlobalServerPath Is Nothing Then
                mGlobalServerPath = String.Format("\\{0}\{1}", "NT-EIB-10-6B16", GlobalShareName)
                mGlobalServerPath = System.IO.Path.Combine(mGlobalServerPath, "Versions")
            End If
            Return mGlobalServerPath
        End Get
    End Property

    ' dgp rev 5/2/2012 
    Private mXMLFileName = Nothing
    Public ReadOnly Property XMLFileName() As String
        Get
            If mXMLFileName Is Nothing Then mXMLFileName = "Persistent.XML"
            Return mXMLFileName
        End Get
    End Property

    ' dgp rev 5/2/2012 
    Public ReadOnly Property XMLFullSpec As String
        Get
            Return System.IO.Path.Combine(GlobalServerPath, XMLFileName)
        End Get
    End Property

    ' dgp rev 5/2/2012 
    Public Function ClearData() As Boolean

        ClearData = True
        If System.IO.File.Exists(XMLFullSpec) Then
            Try
                System.IO.File.Delete(XMLFullSpec)
            Catch ex As Exception
                ClearData = False
            End Try
        End If

    End Function

    ' dgp rev 5/31/2012
    Private Function SaveXML(CurElement As XElement) As Boolean

        SaveXML = False
        If Utility.Create_Tree(GlobalServerPath) Then
            SaveXML = True
            Try
                CurElement.Save(XMLFullSpec)
            Catch ex As Exception
                SaveXML = False
            End Try
        End If

    End Function

    ' dgp rev 5/31/2012
    Private Function SaveXML() As Boolean

        SaveXML = False
        If Utility.Create_Tree(GlobalServerPath) Then
            SaveXML = True
            Try
                AppendXElement.Save(XMLFullSpec)
            Catch ex As Exception
                SaveXML = False
            End Try
        End If

    End Function

    ' dgp rev 5/3/2012 
    Private mGroupName As String = "FlowControl"
    Public Property GroupName As String
        Get
            Return mGroupName
        End Get
        Set(value As String)
            mGroupName = value
        End Set
    End Property

    ' dgp rev 5/3/2012 
    Private mCurXMLDoc As XElement

    ' dgp rev 5/3/2012 
    Public ReadOnly Property CurXMLDoc As XElement
        Get
            If System.IO.File.Exists(CurFullSpec) Then
                mCurXMLDoc = XElement.Load(CurFullSpec)
                mGroupName = mCurXMLDoc.Name.LocalName.ToString
            Else
                mGroupName = System.IO.Path.GetFileNameWithoutExtension(CurFullSpec)
                mCurXMLDoc = New XElement(GroupName)
            End If
            Return mCurXMLDoc
        End Get
    End Property
    ' dgp rev 5/4/2012 
    Public Function RemoveElement(key As String) As Boolean

        RemoveElement = False

        Dim orig As XElement = CurXMLDoc

        Dim items = (From item In orig.Descendants("Item") Where item.Attribute("Timestamp") = key Select item).ToList

        If items IsNot Nothing Then

            items(0).remove()
            SaveXML(orig)
        End If

    End Function

    Private mFound

    ' dgp rev 5/4/2012 
    Public Function ModifyElement(key As String, NewEle As XElement) As Boolean

        ModifyElement = False

        Dim orig As XElement = CurXMLDoc

        Dim items = (From item In orig.Descendants("Item") Where item.Attribute("Timestamp") = key Select item).ToList

        If items IsNot Nothing Then
            items(0).ReplaceWith(NewEle)
            SaveXML(orig)
        End If

    End Function

    ' dgp rev 5/4/2012 
    Public ReadOnly Property SelectElement(key As String) As XElement
        Get
            Try
                Dim found = (From info In CurXMLDoc.Descendants("Item") Where (info.Attribute("Timestamp").Value = key) Select info).ToList
                If found IsNot Nothing Then
                    Return found(0)
                End If
            Catch ex As Exception

            End Try

            Return New XElement("Nothing")

        End Get
    End Property

    ' dgp rev 5/2/2012 
    Private mSingleXElement As XElement
    Public ReadOnly Property AppendXElement As XElement
        Get
            Try
                Dim lastEntry As XElement = CurXMLDoc.Element("Item")
                If lastEntry Is Nothing Then
                    mSingleXElement = New XElement(GroupName, SingleLogElement)
                Else
                    mSingleXElement = CurXMLDoc
                    mSingleXElement.Add(SingleLogElement)
                End If
                SaveXML(mSingleXElement)

            Catch ex As Exception

            End Try

            Return mSingleXElement
        End Get
    End Property

    ' dgp rev 5/3/2012
    Public ReadOnly Property AppendXElement(NewKey As String, NewVal As String) As XElement
        Get
            Try
                Dim lastEntry As XElement = CurXMLDoc.Element("Item")
                If lastEntry Is Nothing Then
                    mSingleXElement = New XElement(GroupName, SingleLogElement(NewKey, NewVal))
                Else
                    mSingleXElement = CurXMLDoc
                    mSingleXElement.Add(SingleLogElement(NewKey, NewVal))
                End If
                SaveXML(mSingleXElement)

            Catch ex As Exception

            End Try

            Return mSingleXElement
        End Get
    End Property

    ' dgp rev 5/3/2012
    Public ReadOnly Property AppendXElement(NewInfo As String) As XElement
        Get
            Try
                Dim lastEntry As XElement = CurXMLDoc.Element("Item")
                If lastEntry Is Nothing Then
                    mSingleXElement = New XElement(GroupName, SingleLogElement(NewInfo))
                Else
                    mSingleXElement = CurXMLDoc
                    mSingleXElement.Add(SingleLogElement(NewInfo))
                End If
                SaveXML(mSingleXElement)

            Catch ex As Exception

            End Try

            Return mSingleXElement
        End Get
    End Property

    ' Instance portion
    ' dgp rev 5/2/2012 Path to Server
    Private mCurShareName As String = "Upload"
    Public Property CurShareName As String
        Get
            Return mCurShareName
        End Get
        Set(value As String)
            mCurShareName = value
        End Set

    End Property

    Private mValid As Boolean = False

    ' dgp rev 5/2/2012 
    Private mCurServerPath = Nothing
    Private mCurFileName = String.Format("{0}.xml", Environment.UserName)
    Public Property CurFileName As String
        Get
            Return mCurFileName
        End Get
        Set(value As String)
            mCurFileName = String.Format("{0}.xml", value.ToLower.Replace(".xml", ""))
        End Set
    End Property

    ' dgp rev 5/23/2012 
    Public ReadOnly Property CurServerPath() As String
        Get
            If mCurServerPath Is Nothing Then
                mCurServerPath = String.Format("\\{0}\{1}", "NT-EIB-10-6B16", CurShareName)
                mCurServerPath = System.IO.Path.Combine(mGlobalServerPath, "Tracking")
            End If
            Return mCurServerPath
        End Get
    End Property

    ' dgp rev 8/29/2012 modify specific user keyword
    Public ReadOnly Property RemoveKey(key As String) As Boolean
        Get
            Dim orig As XElement = CurXMLDoc
            Dim lst = (From item In orig.Elements(ElementName).Attributes(key) Select item).ToList
            If lst.count = 0 Then Return True
            Dim remov = (From item In orig.Elements(ElementName).Attributes(key) Select item.Parent).First
            remov.remove()
            orig.Save(CurFullSpec)
            Return True
        End Get
    End Property

    ' dgp rev 8/29/2012 modify specific user keyword
    Public ReadOnly Property RemoveKey(user As String, key As String) As Boolean
        Get
            Dim orig As XElement = CurXMLDoc
            Dim lst = (From item In orig.Elements(user).Elements(ElementName).Attributes(key) Select item).ToList
            If lst.count = 0 Then Return True
            Dim remov = (From item In orig.Elements(user).Elements(ElementName).Attributes(key) Select item.Parent).First
            remov.remove()
            orig.Save(CurFullSpec)
            Return True
        End Get
    End Property

    ' dgp rev 8/29/2012 modify specific user keyword
    Public ReadOnly Property AddValue(key As String, newval As String) As Boolean
        Get
            Dim orig As XElement = CurXMLDoc
            Dim items = (From item In orig.Elements(ElementName).Attributes(key) Select item).ToList
            If items.count = 0 Then
                Dim newele As XElement = New XElement(ElementName, New XAttribute(key, newval))
                orig.Add(newele)
                orig.Save(CurFullSpec)
            Else
                orig.Elements(ElementName).Attributes(key).First.Value = newval
                orig.Save(CurFullSpec)
            End If
            Return True
        End Get
    End Property

    ' dgp rev 8/29/2012 modify specific user keyword
    Public ReadOnly Property AddUpgrade(ByVal user As String, ByVal Uninstall As String, ByVal Install As String) As Boolean
        Get
            Dim orig As XElement = CurXMLDoc

            If CheckUser(user) Then
                Dim items = (From item In orig.Elements(user).Elements(SoftwareUpgrade.UpgradeStore.CurElementName).Attributes("Uninstall") Select item).ToList
                If items.Count = 0 Then
                    Dim newele As XElement = New XElement(SoftwareUpgrade.UpgradeStore.CurElementName, New XAttribute("Uninstall", Uninstall), New XAttribute("Install", Install))
                    orig.Element(user).Add(newele)
                    orig.Save(CurFullSpec)
                Else
                    orig.Elements(user).Elements(SoftwareUpgrade.UpgradeStore.CurElementName).Attributes("Uninstall").First.Value = Uninstall
                    orig.Elements(user).Elements(SoftwareUpgrade.UpgradeStore.CurElementName).Attributes("Install").First.Value = Install
                    orig.Save(CurFullSpec)
                End If
            Else
                Dim newele As XElement = New XElement(SoftwareUpgrade.UpgradeStore.CurElementName, New XAttribute("Uninstall", Uninstall), New XAttribute("Install", Install))
                orig.Add(newele)
                orig.Save(CurFullSpec)
            End If
            Return True
        End Get
    End Property

    ' dgp rev 8/29/2012 modify specific user keyword
    Public ReadOnly Property AddValue(user As String, key As String, newval As String) As Boolean
        Get
            Dim orig As XElement = CurXMLDoc

            If CheckUser(user) Then
                Dim items = (From item In orig.Elements(user).Elements(ElementName).Attributes(key) Select item).ToList
                If items.count = 0 Then
                    Dim newele As XElement = New XElement(ElementName, New XAttribute(key, newval))
                    orig.Element(user).Add(newele)
                    orig.Save(CurFullSpec)
                Else
                    orig.Elements(user).Elements(ElementName).Attributes(key).First.Value = newval
                    orig.Save(CurFullSpec)
                End If
            Else
                Dim newele As XElement = New XElement(user, New XElement(ElementName, New XAttribute(key, newval)))
                orig.Add(newele)
                orig.Save(CurFullSpec)
            End If
            Return True
        End Get
    End Property

    ' dgp rev 8/29/2012 modify specific user keyword
    Public ReadOnly Property ModifyValue(user As String, key As String, newval As String) As Boolean
        Get
            Dim orig As XElement = CurXMLDoc
            Dim items = (From item In orig.Elements(user).Elements(ElementName).Attributes(key) Select item)
            If items Is Nothing Then Return False
            orig.Elements(user).Elements(ElementName).Attributes(key).First.Value = newval
            orig.Save(CurFullSpec)
            Return True
        End Get
    End Property

    ' dgp rev 8/29/2012 modify global keyword
    Public ReadOnly Property ModifyValue(key As String, newval As String) As Boolean
        Get
            Dim orig As XElement = CurXMLDoc
            Dim items = (From item In orig.Elements(ElementName).Attributes(key) Select item)
            If items Is Nothing Then Return False
            orig.Elements(ElementName).Attributes(key).First.Value = newval
            orig.Save(CurFullSpec)
            Return True
        End Get
    End Property

    ' dgp rev 8/29/2012 global key value
    Public ReadOnly Property GetValue(key As String) As String
        Get
            Dim orig As XElement = CurXMLDoc

            Try
                Dim items = (From item In orig.Elements(ElementName).Attributes(key) Select item.Value).ToList
                If items IsNot Nothing Then
                    Return items(0).ToString
                End If
            Catch ex As Exception
                Return ""
            End Try
            Return ""
        End Get
    End Property

    Public Function CheckUser(ByVal user As String) As Boolean

        CheckUser = False
        Dim orig As XElement = CurXMLDoc

        Dim items = (From item In orig.Elements(user)).ToList
        Return Not (items.count = 0)

    End Function

    ' dgp rev 8/29/2012 specific user key value
    Public ReadOnly Property GetValue(ByVal user As String, ByVal key As String) As String
        Get
            Dim orig As XElement = CurXMLDoc

            If Not CheckUser(user) Then Return ""

            Try
                Dim items = (From item In orig.Elements(user).Elements(ElementName).Attributes(key) Select item.Value).ToList
                If items.count = 0 Then Return ""
                Return items(0).ToString
            Catch ex As Exception
                Return ""
            End Try
        End Get
    End Property



    ' dgp rev 8/29/2012 specific user key value
    Public ReadOnly Property GetAttributes(ByVal user As String, ByVal key As String) As ArrayList
        Get
            Dim orig As XElement = CurXMLDoc

            Dim Attribes As New ArrayList
            If Not CheckUser(user) Then Return Attribes

            Try
                Dim items = (From item In orig.Elements(user).Elements(key).Attributes Select item).ToList
                If items.Count = 0 Then Return Attribes
                Return Attribes
            Catch ex As Exception
                Return Attribes
            End Try
        End Get
    End Property

    ' dgp rev 8/29/2012 retrieve list of users
    Public ReadOnly Property GetUsers As ArrayList
        Get
            Dim orig As XElement = CurXMLDoc
            Dim NewArray As ArrayList = New ArrayList

            Dim items = (From item In orig.Elements() Where Not item.Name = ElementName Select item).ToList
            If items IsNot Nothing Then
                If items.count > 0 Then
                    Dim item As XElement
                    For Each item In items
                        NewArray.Add(item.Name.LocalName)
                    Next
                End If
            End If
            Return NewArray
        End Get
    End Property

    ' dgp rev 8/29/2012 retrieve list of keys for select user
    Public ReadOnly Property GetKeys(ByVal user As String) As ArrayList
        Get
            Dim orig As XElement = CurXMLDoc
            Dim NewArray As ArrayList = New ArrayList

            Dim items = (From item In orig.Elements(user) Select item).ToList
            If items IsNot Nothing Then
                If items.count > 0 Then
                    Dim item As XElement
                    For Each item In items.Elements
                        NewArray.Add(item.Name.LocalName)
                    Next
                End If
            End If
            Return NewArray
        End Get
    End Property

    ' dgp rev 8/29/2012 retrieve list of keys for select user
    Public ReadOnly Property GetSubKeys(ByVal user As String) As ArrayList
        Get
            Dim orig As XElement = CurXMLDoc
            Dim NewArray As ArrayList = New ArrayList

            Dim items = (From item In orig.Elements(user).Elements(ElementName) Select item).ToList
            If items IsNot Nothing Then
                If items.Count > 0 Then
                    Dim item As XElement
                    For Each item In items
                        NewArray.Add(item.Attributes.ElementAt(0).Name.LocalName)
                    Next
                End If
            End If
            Return NewArray
        End Get
    End Property



    ' dgp rev 8/29/2012 retrieve list of global keys
    Public ReadOnly Property GetKeys As ArrayList
        Get
            Dim orig As XElement = CurXMLDoc
            Dim NewArray As ArrayList = New ArrayList

            Dim items = (From item In orig.Elements(ElementName) Select item).ToList
            If items IsNot Nothing Then
                If items.count > 0 Then
                    Dim item As XElement
                    For Each item In items
                        NewArray.Add(item.Attributes.ElementAt(0).Name.LocalName)
                    Next
                End If
            End If
            Return NewArray
        End Get
    End Property

    ' dgp rev 5/23/2012 
    Public ReadOnly Property Exists(user As String, key As String) As String
        Get
            Dim orig As XElement = CurXMLDoc
            Dim lst = (From item In orig.Elements(user).Elements(ElementName).Attributes(key) Select item).ToList
            Return Not (lst.count = 0)
        End Get
    End Property

    ' dgp rev 5/23/2012 
    Public ReadOnly Property Exists(key As String) As String
        Get
            Dim orig As XElement = CurXMLDoc
            Dim lst = (From item In orig.Elements(ElementName).Attributes(key) Select item).ToList
            Return Not (lst.count = 0)
        End Get
    End Property

    Private mCurUser As String = Environment.UserName
    Public ReadOnly Property CurUser As String
        Get
            Return mCurUser
        End Get
    End Property

    ' dgp rev 5/3/2012
    Public ReadOnly Property CurAppendXElement(NewKey As String, NewVal As String) As XElement
        Get
            Try
                Dim lastEntry As XElement = CurXMLDoc.Element(CurUser)
                If lastEntry Is Nothing Then
                    mSingleXElement = New XElement(GroupName, SingleLogElement(NewKey, NewVal))
                Else
                    mSingleXElement = CurXMLDoc
                    mSingleXElement.Add(SingleLogElement(NewKey, NewVal))
                End If
                SaveXML(mSingleXElement)

            Catch ex As Exception

            End Try

            Return mSingleXElement
        End Get
    End Property

    ' dgp rev 5/23/2012 
    Public ReadOnly Property PutSetting(key As String, val As String) As Boolean
        Get
            Dim info = CurAppendXElement(key, val)
            Return True
        End Get
    End Property

    Private mNewFilename = Nothing
    Public Property NewFilename As String
        Get
            Return mNewFilename
        End Get
        Set(value As String)
            mNewFilename = value
        End Set
    End Property

    ' dgp rev 5/2/2012 
    Public ReadOnly Property CurFullSpec As String
        Get
            If mNewFilename Is Nothing Then Return mFullFileSpec
            Return mNewFilename
        End Get
    End Property

    Private mFullFileSpec As String

    ' dgp rev 5/23/2012 New XML persistence file
    Public Sub New(FullFileSpec As String)

        mFullFileSpec = FullFileSpec
        Dim uri As Uri = New Uri(FullFileSpec)
        If uri.IsUnc Then
        Else
        End If
        CurShareName = uri.AbsolutePath

        CurFileName = System.IO.Path.GetFileNameWithoutExtension(FullFileSpec)
        mXMLFileName = System.IO.Path.GetFileNameWithoutExtension(FullFileSpec)
        mValid = System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(FullFileSpec))

    End Sub

    ' dgp rev 5/23/2012 New XML persistence file
    Public Sub New(Share As String, Filename As String)

        Dim uri As Uri = New Uri(Share)
        If uri.IsUnc Then
        Else
        End If
        mFullFileSpec = System.IO.Path.Combine(Share, String.Format("{0}.xml", Filename.ToLower.Replace(".xml", "")))
        CurShareName = uri.AbsolutePath

        CurFileName = Filename
        mXMLFileName = Filename
        mValid = System.IO.Directory.Exists(Share)

    End Sub
End Class
