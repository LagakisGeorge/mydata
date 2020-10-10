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

Dim SQL As String

Dim gdb As New ADODB.Connection
Dim mark As String

Dim AFM As String

Dim issueDate As String
Dim aa As String
Dim invoiceType As String


Dim paytype As String
Dim payaji As String

Dim totalNetValue As String
Dim totalVatAmount As String
Dim totalGrossValue As String

Dim classificationType As String

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
gdb.Open "DSN=MERCSQL;"
    Set nodeList = objXML.selectNodes("RequestedDoc/invoicesDoc/invoice")

    For Each node In nodeList
        Print node.nodeName  ' this works'
        List1.AddItem node.nodeName + "*******" + node.Text  ' INVOICE
       
        '       Print node.n
        Call printNode(node)     'here is the problem explained below'
    
    If Len(Trim(payaji)) = 0 Then payaji = "0"
    If Len(Trim(totalGrossValue)) = 0 Then totalGrossValue = "0"
    

SQL = "INSERT INTO APESTALMENA  ([MARK],[AFM],[ISSUEDATE],[AA],[TYPOS],[PAYTYPE],[PAYAJI],[TOTALNETVALUE],[TOTALVATAMOUNT],[TOTALGROSSVALUE],[CLASSIFICATIONTYPE])"
SQL = SQL + " Values ('" + mark + "','" + AFM + "','" + issueDate + "','" + aa + "','" + invoiceType + "','" + paytype + "'," + Replace(payaji, ",", ".") + "," + Replace(totalNetValue, ",", ".") + "," + Replace(totalVatAmount, ",", ".") + "," + Replace(totalGrossValue, ",", ".") + ",'" + classificationType + "')"

gdb.Execute SQL

    
    
    
    
    
    Next node
End Sub

Public Sub printNode(node As IXMLDOMNode)
    Dim xmlNode As IXMLDOMNode
     Dim xml2Node As IXMLDOMNode
      Dim xml3Node As IXMLDOMNode
     
    If node.hasChildNodes Then
       
        For Each xmlNode In node.childNodes
           List1.AddItem xmlNode.nodeName + "*" + xmlNode.Text  ' MARK - ISSUER -COUNTERPART-HEADER
           ' Print xmlNode.nodeName
           
           If xmlNode.nodeName = "mark" Then mark = xmlNode.Text
           
            If UCase(xmlNode.nodeName) = "COUNTERPART" Then
                   For Each xml2Node In xmlNode.childNodes
                       If UCase(xml2Node.nodeName) = "VATNUMBER" Then AFM = xml2Node.Text
                        'Call printNode(xmlNode)
                    Next
            End If
            If (xmlNode.nodeName) = "invoiceHeader" Then
                   For Each xml2Node In xmlNode.childNodes
                       If (xml2Node.nodeName) = "issueDate" Then issueDate = xml2Node.Text
                       If (xml2Node.nodeName) = "aa" Then aa = xml2Node.Text
                        If (xml2Node.nodeName) = "invoiceType" Then invoiceType = xml2Node.Text
                        'Call printNode(xmlNode) invoiceType
                    Next
            End If
            If (xmlNode.nodeName) = "paymentMethods" Then
                   For Each xml2Node In xmlNode.childNodes
                       If (xml2Node.nodeName) = "type" Then paytype = xml2Node.Text
                       If (xml2Node.nodeName) = "amount" Then payaji = xml2Node.Text
                        'Call printNode(xmlNode)
                    Next
            End If
            If (xmlNode.nodeName) = "invoiceDetails" Then
                   For Each xml2Node In xmlNode.childNodes
                       If (xml2Node.nodeName) = "lineNumber" Then
                          'add eggtim line
                       End If
                       
                    Next
            End If
            'invoiceSummary
            If (xmlNode.nodeName) = "invoiceSummary" Then
                   For Each xml2Node In xmlNode.childNodes
                       If (xml2Node.nodeName) = "totalNetValue" Then totalNetValue = xml2Node.Text
                       If (xml2Node.nodeName) = "totalVatAmount" Then totalVatAmount = xml2Node.Text
                       
                       If (xml2Node.nodeName) = "incomeClassification" Then
                            For Each xml3Node In xml2Node.childNodes
                       
                                If (xml3Node.nodeName) = "icls:classificationType" Then
                                      classificationType = xml3Node.Text
                                End If
                       
                       
                       
                            Next
                       
                       End If
                       
                        'Call printNode(xmlNode)
                    Next
            End If
            
            
'             <totalNetValue>53.65</totalNetValue>
'        <totalVatAmount>12.88</totalVatAmount>
'        <totalWithheldAmount>0</totalWithheldAmount>
'        <totalFeesAmount>0</totalFeesAmount>
'        <totalStampDutyAmount>0</totalStampDutyAmount>
'        <totalOtherTaxesAmount>0</totalOtherTaxesAmount>
'        <totalDeductionsAmount>0</totalDeductionsAmount>
'        <totalGrossValue>66.53</totalGrossValue>
'        <incomeClassification>
'          <icls:classificationType>E3_561_001</icls:classificationType>
'          <icls:classificationCategory>category1_1</icls:classificationCategory>
'          <icls:amount>53.65</icls:amount>
'        </incomeClassification>
'      </invoiceSummary>
            
            
            
            
            
            
            
            
           ' Call print2Node(xmlNode)
        Next xmlNode
       ' Print node.nodeName
'List1.AddItem node.nodeName + "*" + node.Text
        
    End If
End Sub


Public Sub print2Node(node As IXMLDOMNode)
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
        .setProperty "SelectionLanguage", "XPath"
        .setProperty "SelectionNamespaces", "xmlns:s='http://www.w3.org/ns/widgets'"
        .validateOnParse = True
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
        Set nodesXML = objXML.documentElement.selectSingleNode(arrNodes(i))
        MsgBox nodesXML.Text
    Next

    Set nodesXML = Nothing: Set objXML = Nothing
End Sub

