Imports System.Xml
Imports HelperClasses

' dgp rev 5/2/07 Class to handle persistence
Public Class Dynamic

    Dim xmlDoc As New XmlDocument

    Private mDoc_Path As String
    Private mDoc_Name As String
    Private mDoc_Spec As String

    Private mValid As Boolean = False

    Public ReadOnly Property XMLPath() As String
        Get
            Return mDoc_Spec
        End Get
    End Property

    Public ReadOnly Property Valid() As Boolean
        Get
            Return mValid
        End Get
    End Property

    ' dgp rev 2/19/08 Flag XML File as New
    Private mNewFlag As Boolean = True
    Public ReadOnly Property NewFlag() As Boolean
        Get
            Return mNewFlag
        End Get
    End Property

    ' dgp rev 3/5/09 Initialize XML doc
    Private Function Init() As Boolean

        If (Utility.Create_Tree(mDoc_Path)) Then
            mDoc_Spec = System.IO.Path.Combine(mDoc_Path, mDoc_Name)
            mNewFlag = Not System.IO.File.Exists(mDoc_Spec)

            Try
                xmlDoc.Load(mDoc_Spec)
                Init = True
            Catch
                xmlDoc.LoadXml("<dynamic></dynamic>")
                Init = False
            End Try
        Else
            Init = False
        End If

    End Function

    ' dgp rev 5/2/07 Create a new instance with default settings
    Public Sub New()

        mDoc_Name = "Dynamic.XML"
        mDoc_Path = Utility.MyAppPath

        mValid = Init()

    End Sub
    ' dgp rev 5/1/07 New Instance with path provided
    Public Sub New(ByVal Name As String)

        If (Not Name.ToLower.Contains(".xml")) Then
            Name = Name + ".xml"
        End If

        mDoc_Name = Name
        mDoc_Path = Utility.MyAppPath

        mValid = Init()

    End Sub
    ' dgp rev 5/1/07 New Instance with path and name provided
    Public Sub New(ByVal Path As String, ByVal Name As String)

        If (Not Name.ToLower.Contains(".xml")) Then
            Name = Name + ".xml"
        End If
        mDoc_Name = Name
        mDoc_Path = Path

        mValid = Init()

    End Sub
    ' dgp rev 4/16/07 Get dynamic without Default
    Public Function GetSetting(ByVal xPath As String) As String

        Dim XmlNode As XmlNode = xmlDoc.SelectSingleNode("dynamic/" + xPath)
        If (XmlNode Is Nothing) Then
            Return ""
        Else
            Return XmlNode.InnerText
        End If

    End Function
    ' dgp rev 4/16/07 does the setting keyword exist
    Public Function Exists(ByVal xPath As String) As Boolean

        Dim XmlNode As XmlNode = xmlDoc.SelectSingleNode("dynamic/" + xPath)
        Return (Not XmlNode Is Nothing)

    End Function
    ' dgp rev 5/2/07 Get Setting as String Value
    Public Function GetSetting(ByVal xPath As String, ByVal defaultValue As String) As String

        Dim XmlNode As XmlNode = xmlDoc.SelectSingleNode("dynamic/" + xPath)
        If (XmlNode Is Nothing) Then
            Return defaultValue
        Else
            Return XmlNode.InnerText
        End If

    End Function

    ' dgp rev 5/2/07 Get Setting as String Value
    Public Function RemoveSetting(ByVal xPath As String) As Boolean

        Dim XmlNode As XmlNode = xmlDoc.SelectSingleNode("dynamic/" + xPath)
        If (XmlNode Is Nothing) Then
            Return True
        Else
            XmlNode.ParentNode.RemoveChild(XmlNode)
            XmlNode = xmlDoc.SelectSingleNode("dynamic/" + xPath)
            xmlDoc.Save(Me.mDoc_Spec)
        End If

    End Function
    ' dgp rev 5/2/07 Save Integer Value
    Public Sub PutSetting(ByVal xPath As String, ByVal value As Integer)

        PutSetting(xPath, Convert.ToString(value))

    End Sub

    ' dgp rev 5/2/07 Save String Value 
    Public Sub PutList(ByVal xPath As String, ByVal NewItem As String, ByVal ht As Hashtable)

        Dim xmlnode As XmlNode
        Dim XmlNodeList As XmlNodeList = xmlDoc.SelectNodes("dynamic/" + xPath)

        Dim xmlelement As XmlElement = xmlDoc.CreateElement(xPath)
        Dim xmltext As XmlText
        Dim xmlattr As XmlAttribute
        Dim item As DictionaryEntry
        xmltext = xmlDoc.CreateTextNode(NewItem)
        For Each item In ht
            xmlattr = xmlDoc.CreateAttribute(item.Key)
            xmlattr.Value = item.Value
            xmlelement.Attributes.Append(xmlattr)
        Next

        xmlelement.AppendChild(xmltext)
        If (XmlNodeList.Count = 0) Then
            XmlNodeList = xmlDoc.SelectNodes("dynamic")
            xmlnode = XmlNodeList.Item(0)
            xmlnode.InsertAfter(xmlelement, xmlnode.LastChild)
        Else
            xmlnode = XmlNodeList.Item(0)
            xmlnode.ParentNode.InsertAfter(xmlelement, xmlnode.ParentNode.LastChild)
        End If

        Try
            xmlDoc.Save(mDoc_Spec)
        Catch ex As Exception

        End Try
    End Sub


    ' dgp rev 5/2/07 Save String Value 
    Public Sub PutList(ByVal xPath As String, ByVal item As DictionaryEntry)

        Dim xmlnode As XmlNode
        Dim XmlNodeList As XmlNodeList = xmlDoc.SelectNodes("dynamic/" + xPath)

        Dim xmlelement As XmlElement = xmlDoc.CreateElement(xPath)
        Dim xmltext As XmlText = xmlDoc.CreateTextNode(item.Key)
        Dim xmlattr As XmlAttribute = xmlDoc.CreateAttribute("File")
        xmlattr.Value = item.Value

        xmlelement.Attributes.Append(xmlattr)
        xmlelement.AppendChild(xmltext)
        If (XmlNodeList.Count = 0) Then
            XmlNodeList = xmlDoc.SelectNodes("dynamic")
            xmlnode = XmlNodeList.Item(0)
            xmlnode.InsertAfter(xmlelement, xmlnode.LastChild)
        Else
            xmlnode = XmlNodeList.Item(0)
            xmlnode.ParentNode.InsertAfter(xmlelement, xmlnode.ParentNode.LastChild)
        End If

        Try
            xmlDoc.Save(mDoc_Spec)
        Catch ex As Exception

        End Try
    End Sub


    ' dgp rev 5/2/07 Save String Value 
    Public Sub PutList(ByVal xPath As String, ByVal value As XmlNode)

        Dim xmlnode As XmlNode
        Dim XmlNodeList As XmlNodeList = xmlDoc.SelectNodes("dynamic/" + xPath)

        If (XmlNodeList.Count = 0) Then
            XmlNodeList = xmlDoc.SelectNodes("dynamic")
            xmlnode = XmlNodeList.Item(0)
            xmlnode.InsertAfter(value, xmlnode.LastChild)
        Else
            xmlnode = XmlNodeList.Item(0)
            xmlnode.ParentNode.InsertAfter(value, xmlnode.ParentNode.LastChild)
        End If

        Try
            xmlDoc.Save(mDoc_Spec)
        Catch ex As Exception

        End Try
    End Sub

    ' dgp rev 5/2/07 Save String Value 
    Public Sub PutList(ByVal xPath As String, ByVal value As String)

        Dim xmlnode As XmlNode
        Dim XmlNodeList As XmlNodeList = xmlDoc.SelectNodes("dynamic/" + xPath)

        Dim xmlelement As XmlElement = xmlDoc.CreateElement(xPath)
        Dim xmltext As XmlText = xmlDoc.CreateTextNode(value)

        xmlelement.AppendChild(xmltext)
        If (XmlNodeList.Count = 0) Then
            XmlNodeList = xmlDoc.SelectNodes("dynamic")
            xmlnode = XmlNodeList.Item(0)
            xmlnode.InsertAfter(xmlelement, xmlnode.LastChild)
        Else
            xmlnode = XmlNodeList.Item(0)
            xmlnode.ParentNode.InsertAfter(xmlelement, xmlnode.ParentNode.LastChild)
        End If

        Try
            xmlDoc.Save(mDoc_Spec)
        Catch ex As Exception

        End Try
    End Sub



    ' dgp rev 5/2/07 Save String Value 
    Public Sub PutSetting(ByVal xPath As String, ByVal value As String)

        Dim XmlNode As XmlNode = xmlDoc.SelectSingleNode("dynamic/" + xPath)

        If (XmlNode Is Nothing) Then XmlNode = createMissingNode("dynamic/" + xPath)
        XmlNode.InnerText = value
        Try
            xmlDoc.Save(mDoc_Spec)
        Catch ex As Exception

        End Try
    End Sub

    ' dgp rev 5/2/07 Create any missing nodes
    Public Function GetKeys() As ArrayList

        Dim testNode As XmlNode = Nothing
        Dim currentNode As XmlNode = xmlDoc.SelectSingleNode("dynamic")

        Dim keys As New ArrayList
        For Each testNode In currentNode.SelectNodes("*")
            If (testNode.ChildNodes.Count > 0) Then keys.Add(testNode.Name)
        Next

        Return keys

    End Function

    ' dgp rev 5/2/07 Create any missing nodes
    Private Function createMissingNode(ByVal xPath As String) As XmlNode

        Dim xPathSections() As String = xPath.Split("/")
        Dim currentXPath As String = ""
        Dim testNode As XmlNode = Nothing
        Dim currentNode As XmlNode = xmlDoc.SelectSingleNode("dynamic")
        Dim xPathSection As String
        For Each xPathSection In xPathSections
            currentXPath += xPathSection
            testNode = xmlDoc.SelectSingleNode(currentXPath)
            If (testNode Is Nothing) Then currentNode.InnerXml += "<" + xPathSection + "></" + xPathSection + ">"
            currentNode = xmlDoc.SelectSingleNode(currentXPath)
            currentXPath += "/"
        Next
        Return currentNode

    End Function

End Class
