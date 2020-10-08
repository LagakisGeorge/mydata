VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   5520
      Left            =   4200
      TabIndex        =   1
      Top             =   480
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   1920
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim doc as XmlDocument = new XmlDocument()
'doc.Load ("config.xml")
'Dim root As XmlElement = doc.DocumentElement
'PackageName = root.GetAttribute("id")
'ProjectVersion = root.GetAttribute("version")
'ProjectName = root.GetElementsByTagName("name").Item(0).InnerText


    Set objXML = CreateObject("Msxml.DOMDocument")
    objXML.async = True
    objXML.Load "C:\txtfiles\apantreq.xml"

    Dim nodeList As IXMLDOMNodeList
    Dim node As IXMLDOMNode

    Set nodeList = objXML.selectNodes("*")

    For Each node In nodeList
        Print node.nodeName  ' this works'
        'print node.appendChild(
        'Print node.n
        Call printNode(node)     'here is the problem explained below'
    Next node
End Sub

Public Sub printNode(node As IXMLDOMNode)
    Dim xmlNode As IXMLDOMNode

    If node.hasChildNodes Then
        'List1.AddItem node.childNodes.Item(0)
        For Each xmlNode In node.childNodes
            Call printNode(xmlNode)
        Next xmlNode
       ' Print node.nodeName
List1.AddItem node.nodeName + "*" + node.Text
        
    End If
End Sub




Sub test()






    Dim objXML, arrNodes, nodesXML, i

    Set objXML = CreateObject("MSXML2.DOMDocument.6.0")
    With objXML
        .SetProperty "SelectionLanguage", "XPath"
        .SetProperty "SelectionNamespaces", "xmlns:s='http://www.w3.org/ns/widgets'"
        .ValidateOnParse = True
        .async = False
        .Load "C:\txtfiles\apantreq.xml"
    End With

    arrNodes = Array("/s:widget/s:preference[@name='android-minSdkVersion']/@value", _
                    "/s:widget/s:preference[@name='android-versionCode']/@value", _
                    "/s:widget/s:preference[@name='android-installLocation']/@value", _
                    "/s:widget/s:preference[@name='android-targetSdkVersion']/@value", _
                    "/s:widget/s:preference[@name='orientation']/@value", _
                    "/s:widget/s:preference[@name='fullscreen']/@value", _
                    "/s:widget/@id", _
                    "/s:widget/@version", _
                    "/s:widget/s:name")

    For i = LBound(arrNodes) To UBound(arrNodes)
        Set nodesXML = objXML.DocumentElement.SelectSingleNode(arrNodes(i))
        MsgBox nodesXML.Text
    Next

    Set nodesXML = Nothing: Set objXML = Nothing
End Sub

