VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFF80&
   Caption         =   "Form1"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker EOS 
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   4920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      Format          =   136118273
      CurrentDate     =   44118
   End
   Begin MSComCtl2.DTPicker APO 
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   4920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      Format          =   136118273
      CurrentDate     =   44118
   End
   Begin VB.CommandButton cmdCommand6 
      Caption         =   "Command3"
      Height          =   360
      Left            =   8520
      TabIndex        =   7
      Top             =   6360
      Width           =   990
   End
   Begin VB.CommandButton cmdDOMDocumentUTF8 
      BackColor       =   &H0000FF00&
      Caption         =   "DOMDocument UTF8 τεστ"
      Height          =   720
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6360
      Width           =   2775
   End
   Begin VB.CommandButton cmdCommand5 
      BackColor       =   &H0000FF00&
      Caption         =   "DOMDocument UTF8"
      Height          =   720
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   2775
   End
   Begin VB.CommandButton cmdCommand4 
      Caption         =   "DOMDocument CREATE XML"
      Height          =   600
      Left            =   1560
      TabIndex        =   4
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton cmdCommand3 
      Caption         =   "CREATE XML WITHOUT TEXTWRITER"
      Height          =   360
      Left            =   1560
      TabIndex        =   3
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "SQLSERVER TO XML"
      Height          =   495
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   5520
      Left            =   4200
      TabIndex        =   1
      Top             =   480
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "xml to SQLSERVER"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
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

 Dim docStock As MSXML2.DOMDocument

Dim classificationType As String

Private Sub cmdCommand3_Click()
'Set a reference to "Microsoft XML, v3.0"

'<?xml version="1.0" encoding="UTF-8"?>
'<log>
'<error>
' <DateError>12/10/2020</DateError>
'  <ErrorCode>1562</ErrorCode>
'</error>
'
'</log>


Dim I As Integer

    Randomize
    
    I = Int(Rnd() * 5000) + 1
    LogError Now, I
End Sub

Private Sub LogError(ByVal ErrDate As Date, ByVal ErrCode As Long)

    Dim objDoc                   As MSXML2.DOMDocument

    Dim xmlProcessingInstruction As MSXML2.IXMLDOMProcessingInstruction

    Dim objRoot                  As IXMLDOMNode

    Dim objNode                  As IXMLDOMNode

    Dim objNodeErr               As IXMLDOMNode

    Dim objNodeDetails           As IXMLDOMNode

    Dim strLogPath               As String

    strLogPath = "c:\ErrorLog.xml"

    'Create an XML Document object
    Set objDoc = New MSXML2.DOMDocument
    
    'Check if the log exists. If it does open it otherwise create it
    If Dir(strLogPath) = "" Then
        'Creating standard headers
        Set xmlProcessingInstruction = objDoc.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
        objDoc.appendChild xmlProcessingInstruction
        Set xmlProcessingInstruction = Nothing
        
        'Create the Root Node
        Set objRoot = objDoc.createElement("log")
        objDoc.appendChild objRoot
        Set objRoot = Nothing
    Else
        objDoc.Load strLogPath
    End If
    
    'Get the Root Node
    Set objRoot = objDoc.selectSingleNode("log")
    
    'Add an error node to the root node
    Set objNodeErr = objDoc.createElement("error")
    objRoot.appendChild objNodeErr
    Set objRoot = Nothing
    
    'Add the details to the error node
    Set objNodeDetails = objDoc.createElement("DateError")
    objNodeDetails.Text = FormatDateTime(ErrDate, 2)
    objNodeErr.appendChild objNodeDetails
    
    Set objNodeDetails = objDoc.createElement("ErrorCode")
    objNodeDetails.Text = ErrCode
    objNodeErr.appendChild objNodeDetails
    
    'clean up
    Set objNodeDetails = Nothing
    Set objNodeErr = Nothing
    Set objNode = Nothing
    
    'Save the log
    objDoc.save strLogPath

    Set objDoc = Nothing

End Sub

Private Sub cmdCommand4_Click()
' This procedure creates XML document
' and saves it to disk.
' Requires msxml.dll (Go to Project --> References and
' and choose Microsoft XML version 2.0, or whatever the
' current version you have installed)
' The example given below will write the following XML
' documents.
'
' <Family>
'    <Member Relationship="Father">
'       <Name>Some Guy</name>
'    </member>
' </family>
'
'but it should be clear how to modify the code
'to create your own documents

   
   Dim objDom As DOMDocument
   Dim objRootElem As IXMLDOMElement
   Dim objMemberElem As IXMLDOMElement
   Dim objMemberRel As IXMLDOMAttribute
   Dim objMemberName As IXMLDOMElement
   
   Set objDom = New DOMDocument
   
   ' Creates root element
   Set objRootElem = objDom.createElement("Family")
   objDom.appendChild objRootElem
   
   ' Creates Member element
   Set objMemberElem = objDom.createElement("Member")
   objRootElem.appendChild objMemberElem
   
   ' Creates Attribute to the Member Element
   Set objMemberRel = objDom.createAttribute("Relationship")
   objMemberRel.nodeValue = "Father"
   objMemberElem.setAttributeNode objMemberRel
   
   ' Create element under Member element, and
   ' gives value "some guy"
   Set objMemberName = objDom.createElement("Name")
   objMemberElem.appendChild objMemberName
   objMemberName.Text = "Some Guy"

   ' Saves XML data to disk.
   objDom.save ("c:\andrew.xml")

End Sub

Private Sub cmdCommand5_Click()

'<?xml version="1.0" encoding="utf-8"?>
'<ArrayOfStock xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">

'<Stock>
'<ProductCode>12345</ProductCode>
'<ProductPrice>10.32</ProductPrice>
'</Stock>

'<Stock>
'<ProductCode>₯45632</ProductCode>
'<ProductPrice>5.43</ProductPrice>
'</Stock>

'</ArrayOfStock>






 Dim varStock As Variant
   'global       Dim docStock As MSXML2.DOMDocument
    Dim elemRoot As MSXML2.IXMLDOMElement
    Dim elemStock As MSXML2.IXMLDOMElement
    Dim elemField As MSXML2.IXMLDOMElement
    Dim I As Integer
    
    varStock = Array(Array("ΕΥΡΩ12345", 10.32), _
                     Array("₯45632", 5.43)) 'Yen sign used here to show Unicode.
    
    Set docStock = New MSXML2.DOMDocument
    With docStock
        .appendChild .createProcessingInstruction("xml", _
                                                  "version=""1.0"" encoding=""utf-8""")
        Set elemRoot = .createElement("ArrayOfStock")
        With elemRoot
            .setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
            .setAttribute "xmlns:xsd", "http://www.w3.org/2001/XMLSchema"
            
            For I = 0 To UBound(varStock)
                Set elemStock = docStock.createElement("Stock")
                With elemStock
                    Set elemField = docStock.createElement("ProductCode")
                    elemField.Text = CStr(varStock(I)(0))
                    .appendChild elemField
                    Set elemField = docStock.createElement("ProductPrice")
                    elemField.Text = CStr(varStock(I)(1))
                    .appendChild elemField
                End With
                .appendChild elemStock
            Next
        End With
        Set .documentElement = elemRoot
        On Error Resume Next
        Kill "C:\created.xml"
        On Error GoTo 0
        .save "C:\created.xml"
    End With
    
    
    
    
    
    
End Sub

' Add formatting to the document.
Private Sub FormatXmlDocument(ByVal xml_doc As DOMDocument)
    FormatXmlNode xml_doc.documentElement, 0
End Sub

' Add formatting to this element. Indent it and add a
' carriage return before its children. Then recursively
' format the children with increased indentation.
Private Sub FormatXmlNode(ByVal node As IXMLDOMNode, ByVal _
    indent As Integer)
Dim child As IXMLDOMNode
Dim text_only As Boolean

    ' Do nothing if this is a text node.
    If TypeOf node Is IXMLDOMText Then Exit Sub

    ' See if this node contains only text.
    text_only = True
    If node.hasChildNodes Then
        For Each child In node.childNodes
            If Not (TypeOf child Is IXMLDOMText) Then
                text_only = False
                Exit For
            End If
        Next child
    End If

    ' Process child nodes.
    If node.hasChildNodes Then
        ' Add a carriage return before the children.
        If Not text_only Then
            node.insertBefore _
                node.ownerDocument.createTextNode(vbCrLf), _
                node.firstChild
        End If

        ' Format the children.
        For Each child In node.childNodes
            FormatXmlNode child, indent + 2
        Next child
    End If

    ' Format this element.
    If indent > 0 Then
        ' Indent before this element.
        node.parentNode.insertBefore _
            node.ownerDocument.createTextNode(Space$(indent)), _
 _
            node

        ' Indent after the last child node.
        If Not text_only Then _
            node.appendChild _
                node.ownerDocument.createTextNode(Space$(indent))

        ' Add a carriage return after this node.
        If node.nextSibling Is Nothing Then
            node.parentNode.appendChild _
                node.ownerDocument.createTextNode(vbCrLf)
        Else
            node.parentNode.insertBefore _
                node.ownerDocument.createTextNode(vbCrLf), _
                node.nextSibling
        End If
    End If
End Sub








Private Sub cmdDOMDocumentUTF8_Click()


 Dim varStock As Variant
  '  Dim docStock As MSXML2.DOMDocument
    Dim elemRoot As MSXML2.IXMLDOMElement
    Dim elemStock As MSXML2.IXMLDOMElement
    Dim elemField As MSXML2.IXMLDOMElement
    Dim I As Integer
    
    varStock = Array(Array("ΕΥΡΩ12345", 10.32), _
                     Array("₯45632", 5.43)) 'Yen sign used here to show Unicode.
    
    Set docStock = New MSXML2.DOMDocument
    With docStock
        .appendChild .createProcessingInstruction("xml", _
                                                  "version=""1.0"" encoding=""utf-8""")
        Set elemRoot = .createElement("ArrayOfStock")
      '  With elemRoot/////////////////////////////////////////////////////////////////////////////////////////
            elemRoot.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
            elemRoot.setAttribute "xmlns:xsd", "http://www.w3.org/2001/XMLSchema"
            
            For I = 0 To UBound(varStock)
                Set elemStock = docStock.createElement("Stock")
               ' With elemStock----------------------------------
               
'                    add_c Stock, "ProductCode", "11111"
'                     add_c Stock, "ProductPrice", "1.35"
               
                    Set elemField = docStock.createElement("ProductCode")
                       
                       
                       ' δημιουργω εσοχη
                        Set elem2Field = docStock.createElement("Product2Price")
                            elem2Field.Text = "22222"
                            elemField.appendChild elem2Field

                       
                       
                       
                       
                     elemStock.appendChild elemField
                    
                    
                    
                    
                    
                    
                    Set elemField = docStock.createElement("ProductPrice"): elemField.Text = "22222": elemStock.appendChild elemField

                    
               ' End With--------------------------------------
                elemRoot.appendChild elemStock
            Next
      '  End With  /////////////////////////////////////////////////////////////////////////////////////////////
        Set .documentElement = elemRoot
        On Error Resume Next
        Kill "C:\created.xml"
        On Error GoTo 0
        .save "C:\created.xml"
    End With
    
End Sub

'Sub add_1(elem As MSXML2.IXMLDOMElement, klados As String)
'
'    Set elem = docStock.createElement(klados)
'
'End Sub


Sub add_c(ByRef mParent As MSXML2.IXMLDOMElement, mChild As String, mText As String)
Dim elemField  As MSXML2.IXMLDOMElement
 Set elemField = docStock.createElement(mChild)
                    elemField.Text = mText
                    elemField.appendChild elemField



End Sub





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
gdb.open "DSN=MERCSQL;"
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


Sub TEST()








    Dim objXML, arrNodes, nodesXML, I

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

    For I = LBound(arrNodes) To UBound(arrNodes)
        Set nodesXML = objXML.documentElement.selectSingleNode(arrNodes(I))
        MsgBox nodesXML.Text
    Next

    Set nodesXML = Nothing: Set objXML = Nothing
End Sub

Private Sub Command2_Click()
   Dim n As Integer: n = ToXMLsub
   
End Sub

Function ToXMLsub() As Integer

    '===?G??O ?? XML G?? ?? ????S?????? =================================================================================
    'WHERE (ENTITYMARK IS NULL OR ENTITYMARK='ERROR' ) AND
    'Left(ATIM, 1) In     (  " + PAR + "  )    And
    'HME>='" + Format(APO.Value, "MM/dd/yyyy") + "'  AND HME<='" + Format(EOS.Value, "MM/dd/yyyy") + "'  "
    '<correlatedInvoices>400000017716190</correlatedInvoices>
    ToXMLsub = 1

    Dim pol As String, polepis, ago, AGOEPIS As String

    pol = "": polepis = "": ago = "": AGOEPIS = ""
    'If checkServer(0) Then
    ' MsgBox("OK")
    'End If
    gdb.Execute "UPDATE TIM SET ENTITY=0,ENTLINEN=0"

    Dim DUM As Boolean: DUM = Get_AJ_ASCII(pol, polepis, ago, AGOEPIS)

    Dim PAR, SYNT As String

    PAR = pol + polepis

    Dim SQL As String

    SYNT = ""
    SQL = "SELECT  ID_NUM, AJ1  ,AJ2 , AJ3,AJ4,AJ5,AJI,FPA1,FPA2,FPA3,FPA4,ATIM,"
    SQL = SQL + "HME,PEL.EPO,PEL.AFM,KPE,PEL.DIE,"    '"CONVERT(CHAR(10),HME,3) AS HMEP
    SQL = SQL + "PEL.EPA,PEL.POL,AJ6,FPA6,AJ7,FPA7,TRP,ISNULL(APALAGIFPA,0) AS APALAGIFPA ,ISNULL(PEL.XRVMA,'') AS TK "

    SQL = SQL + "   FROM TIM INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD "
    SQL = SQL + " WHERE (ENTITYMARK IS NULL OR ENTITYMARK='ERROR' ) AND    LEFT(ATIM,1) IN     (  " + PAR + "  )    and HME>='" + Format(APO.Value, "MM/dd/yyyy") + "'  AND HME<='" + Format(EOS.Value, "MM/dd/yyyy") + "'  "
    SQL = SQL + "  AND AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7>0  " + SYNT
    SQL = SQL + " order by HME"       '  OR INCMARK IS NULL OR INCMARK='ERROR'

    Dim sqldt As New ADODB.Recordset

    '  SQL = "SELECT  top 20  AJ1 ,AJ2  from TIM  order by HME"

    sqldt.open SQL, gdb, adOpenDynamic, adLockOptimistic

    If sqldt.EOF Then
        MsgBox ("ΔΕΝ ΒΡΕΘΗΚΑΝ ΕΓΓΡΑΦΕΣ")
        ToXMLsub = 0

        Exit Function

    End If

    Dim varStock As Variant

    '  Dim docStock As MSXML2.DOMDocument
    Dim elemRoot As MSXML2.IXMLDOMElement

    Dim INVOICE As MSXML2.IXMLDOMElement

    Dim elemField As MSXML2.IXMLDOMElement

    Dim I As Integer
    
    '<InvoicesDoc   xmlns="http://www.aade.gr/myDATA/invoice/v1.0"
    ' xsi:schemaLocation="http://www.aade.gr/myDATA/invoice/v1.0 schema.xsd"
    ' xmlns:N1="https://www.aade.gr/myDATA/incomeClassificaton/v1.0"
    ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    Set docStock = New MSXML2.DOMDocument

    With docStock
        .appendChild .createProcessingInstruction("xml", "version=""1.0"" encoding=""utf-8""")
        Set elemRoot = .createElement("InvoicesDoc")

        With elemRoot '/////////////////////////////////////////////////////////////////////////////////////////
            .setAttribute "xmlns", "http://www.aade.gr/myDATA/invoice/v1.0"
            .setAttribute "xsi:schemaLocation", "http://www.aade.gr/myDATA/invoice/v1.0 schema.xsd"
            .setAttribute "xmlns:n1", "https://www.aade.gr/myDATA/incomeClassificaton/v1.0"
            .setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
            
            'Dim I As Long:
            I = -1
            
            Do While Not sqldt.EOF
            
                I = I + 1
            
                '  If checkIntegrity(I) = False Then

                '    MsgBox (" ΠΡΟΒΛΗΜΑ ΣΤΟ ΠΑΡΑΣΤΑΤΙΚΟ " + sqldt(0)("ATIM").ToString)
                '  End If

                Dim EGGTIM As New ADODB.Recordset

                Me.Caption = "ΠΑΡΑΣΤΑΤΙΚΑ " + Str(I)
                EGGTIM.open "SELECT KODE,POSO,TIMM,EKPT,FPA,ISNULL(KAU_AJIA,0) AS KAU_AJIA,ISNULL(MIK_AJIA,0) AS MIK_AJIA FROM EGGTIM WHERE TIMM<>0 AND POSO<>0 AND ID_NUM=" + Str(sqldt("ID_NUM")), gdb, adOpenDynamic, adLockOptimistic
                'Dim DUM As New DataTable
                gdb.Execute "UPDATE TIM SET ENTITY=" + Str(I + 1) + " WHERE ID_NUM=" + Str(sqldt("ID_NUM"))

                sumNet = sqldt("aj1") + sqldt("aj2") + sqldt("aj3") + sqldt("aj4") + sqldt("aj5") + sqldt("aj6") + sqldt("aj7")
                sumFpa = sqldt("fpa1") + sqldt("fpa2") + sqldt("fpa3") + sqldt("fpa4") + sqldt("fpa6") + sqldt("fpa7")

                '1.1;E3_561_001;category1_1
                ctypos = FINDTYPOS(Mid(sqldt("ATIM"), 1, 1)) ' Split(tmpStr, ":")(0)

                If Len(Trim(Split(ctypos, ";")(1))) = 0 Or Len(Trim(Split(ctypos, ";")(2))) = 0 Then
                    'writer.Close()
                    MsgBox ("δεν εχουν ορισθει παραμετροι ΜΥDATA στο παρ/κό " + sqldt("ATIM"))

                    Exit Function

                End If

                ' writer.WriteComment (sqldt("ATIM") + " " + Format(sqldt("HME"), "dd/MM/yyyy"))
            
                ' For I = 0 To UBound(varStock)
                Set INVOICE = docStock.createElement("invoice")

                With INVOICE  '----------------------------------
                      
                    Set elemField = docStock.createElement("uid"): elemField.Text = Str(I + 1): INVOICE.appendChild elemField
                      
                    Set elemField = docStock.createElement("issuer") ' δημιουργω εσοχη
                    Set elem2Field = docStock.createElement("vatNumber"): elem2Field.Text = "028783755": elemField.appendChild elem2Field
                    Set elem2Field = docStock.createElement("country"): elem2Field.Text = "GR": elemField.appendChild elem2Field
                    Set elem2Field = docStock.createElement("branch"): elem2Field.Text = "0": elemField.appendChild elem2Field
                    .appendChild elemField
                                
                    If Mid(Split(ctypos, ";")(0), 1, 2) = "11" Then
                        'lianikh den xreiazetai pelaths
                    Else
                                
                        Set elemField = docStock.createElement("counterpart") ' δημιουργω εσοχ
                        Set elem2Field = docStock.createElement("vatNumber"): elem2Field.Text = Trim(sqldt("AFM")): elemField.appendChild elem2Field
                        Set elem2Field = docStock.createElement("country"): elem2Field.Text = "GR": elemField.appendChild elem2Field
                        Set elem2Field = docStock.createElement("branch"): elem2Field.Text = "0": elemField.appendChild elem2Field
                            
                        Set elem2Field = docStock.createElement("address") ' δημιουργω εσοχ' δημιουργω εσοχ  sqlDT(i)("XRVMA")= TK
                        Set elem3Field = docStock.createElement("postalCode"): elem3Field.Text = sqldt("TK"): elem2Field.appendChild elem3Field
                        Set elem3Field = docStock.createElement("city"): elem3Field.Text = sqldt("POL"): elem2Field.appendChild elem3Field
                        elemField.appendChild elem2Field
                        .appendChild elemField
                    End If

                    Set elemField = docStock.createElement("invoiceHeader") ' δημιουργω εσοχη
                    Set elem2Field = docStock.createElement("series"): elem2Field.Text = "0": elemField.appendChild elem2Field
                    Set elem2Field = docStock.createElement("aa"): elem2Field.Text = Mid(sqldt("ATIM"), 2, 6): elemField.appendChild elem2Field
                    Set elem2Field = docStock.createElement("issueDate"): elem2Field.Text = Format(sqldt("hme"), "yyyy-MM-dd"): elemField.appendChild elem2Field
                    Set elem2Field = docStock.createElement("invoiceType"): elem2Field.Text = Split(ctypos, ";")(0): elemField.appendChild elem2Field
                    Set elem2Field = docStock.createElement("currency"): elem2Field.Text = "EUR": elemField.appendChild elem2Field
                    Set elem2Field = docStock.createElement("exchangeRate"): elem2Field.Text = "1.0": elemField.appendChild elem2Field
                                                       
                    .appendChild elemField
                        
                    Dim cTrp As String: cTrp = FindTRP(Mid(sqldt("TRP"), 1, 1))

                    If Len(cTrp) = 0 Then
                        MsgBox ("ΔΕΝ ΕΧΩ ΑΝΤΙΣΤΟΙΧΙΣΗ ΣΤΟΝ ΤΡΟΠΟ ΠΛΗΡΩΜΗΣ " + sqldt(I)("TRP"))
             
                        Exit Function

                    End If

                    Set elemField = docStock.createElement("paymentMethods") ' δημιουργω εσοχ
                            
                    Set elem2Field = docStock.createElement("paymentMethodDetails") ' δημιουργω εσοχ' δημιουργω εσοχ
                    Set elem3Field = docStock.createElement("type"): elem3Field.Text = cTrp: elem2Field.appendChild elem3Field
                    Set elem3Field = docStock.createElement("amount"): elem3Field.Text = Format(sumNet + sumFpa, "######0.##"): elem2Field.appendChild elem3Field
                    elemField.appendChild elem2Field
                            
                    .appendChild elemField




                     Dim SYN_KAU, SYN_FPA As Double
                        SYN_KAU = 0
                        SYN_FPA = 0
                        Dim fpaRow As Double



                    Dim L As Integer: L = 0

                    Do While Not EGGTIM.EOF
                        'For n = 1 To 3 ' SEIRES TIMOLOGIOY
                        L = L + 1
                        Dim AJ As Single

                        If IsNull(EGGTIM("TIMM")) Then
                            AJ = 0
                        Else
                            AJ = EGGTIM("KAU_AJIA")  ' Math.Round(EGGTIM(L)("POSO") * EGGTIM(L)("TIMM") * (1 - EGGTIM(L)("EKPT") / 100), 2)
                        End If

                        Dim VAT As String

                        '1 ΦΠΑ συντελεστής 24% 24%
                        '2 ΦΠΑ συντελεστής 13% 13%
                        '3 ΦΠΑ συντελεστής 6% 6%
                        '4 ΦΠΑ συντελεστής 17% 17%
                        '5 ΦΠΑ συντελεστής 9% 9%
                        '6 ΦΠΑ συντελεστής 4% 4%
                        '7 ’νευ Φ.Π.Α. 0%
                        '8 Εγγραφές χωρίς ΦΠΑ  (πχ Μισθοδοσία, Αποσβέσεις)

                        If EGGTIM("FPA") = 1 Then '13%
                            VAT = "2"
                            SYN_KAU = SYN_KAU + AJ
                            fpaRow = EGGTIM("MIK_AJIA") - EGGTIM("KAU_AJIA") ' AJ * 0.13
                            SYN_FPA = SYN_FPA + fpaRow

                            'ElseIf EGGTIM(L)("FPA") = 2 Then
                            '   VAT = "1"
                        ElseIf EGGTIM("FPA") = 2 Then
                            VAT = "1"
                            SYN_KAU = SYN_KAU + AJ
                            fpaRow = EGGTIM("MIK_AJIA") - EGGTIM("KAU_AJIA") 'AJ * 0.24
                            SYN_FPA = SYN_FPA + fpaRow

                            ' SYN_FPA = SYN_FPA + AJ * 0.24

                        ElseIf EGGTIM("FPA") = 5 Then
                            VAT = "7"
                            SYN_KAU = SYN_KAU + AJ
                            fpaRow = 0
                            SYN_FPA = SYN_FPA + fpaRow

                        ElseIf EGGTIM("FPA") = 6 Then
                            VAT = "1"
                            SYN_KAU = SYN_KAU + AJ
                            ' SYN_FPA = SYN_FPA + AJ * 0.24

                            fpaRow = EGGTIM("MIK_AJIA") - EGGTIM("KAU_AJIA")  ' AJ * 0.24
                            SYN_FPA = SYN_FPA + fpaRow

                        ElseIf EGGTIM("FPA") = 4 Then
                            VAT = "4"
                            SYN_KAU = SYN_KAU + AJ
                            fpaRow = EGGTIM("MIK_AJIA") - EGGTIM("KAU_AJIA")  ' AJ * 0.06
                            SYN_FPA = SYN_FPA + fpaRow
                        Else ' If EGGTIM(L)("FPA") = 2 Then
                            VAT = "1"
                        End If

                        '-----------------------------------------------  invoiceDetails
                        
                        Set elemField = docStock.createElement("invoiceDetails") ' δημιουργω εσοχ
                        Set elem2Field = docStock.createElement("lineNumber"): elem2Field.Text = Str(L): elemField.appendChild elem2Field
                        Set elem2Field = docStock.createElement("netValue"): elem2Field.Text = Format(AJ, "######0.##"): elemField.appendChild elem2Field
                        Set elem2Field = docStock.createElement("vatCategory"): elem2Field.Text = VAT: elemField.appendChild elem2Field
                        Set elem2Field = docStock.createElement("vatAmount"): elem2Field.Text = Format(fpaRow, "######0.##"): elemField.appendChild elem2Field
                            
                          If fpaRow = 0 Then
                                Set elem2Field = docStock.createElement("vatExemptionCategory"): elem2Field.Text = "1": elemField.appendChild elem2Field
                           End If
                        Set elem2Field = docStock.createElement("incomeClassification") ' δημιουργω εσοχ' δημιουργω εσοχ
                        Set elem3Field = docStock.createElement("n1:classificationType"): elem3Field.Text = Split(ctypos, ";")(1): elem2Field.appendChild elem3Field
                        Set elem3Field = docStock.createElement("n1:classificationCategory"): elem3Field.Text = Split(ctypos, ";")(2): elem2Field.appendChild elem3Field
                        Set elem3Field = docStock.createElement("n1:amount"): elem3Field.Text = Format(AJ, "######0.##"): elem2Field.appendChild elem3Field
                        elemField.appendChild elem2Field
                            
                        .appendChild elemField
                        
                        EGGTIM.MoveNext
                        'Next
                    Loop
      
                     gdb.Execute "UPDATE TIM SET AADEKAU=" + Replace(Format(SYN_KAU, "######0.#####"), ",", ".") + ",AADEFPA=" + Replace(Format(SYN_FPA, "######0.#####"), ",", ".") + " WHERE ID_NUM=" + Str(sqldt("ID_NUM"))
  
      
                    '------------------------------------------------ InvoiceSummary
                        
                    Set elemField = docStock.createElement("invoiceSummary") ' δημιουργω εσοχ
                    Set elem2Field = docStock.createElement("totalNetValue"): elem2Field.Text = Format(SYN_KAU, "######0.##"): elemField.appendChild elem2Field
                    Set elem2Field = docStock.createElement("totalVatAmount"): elem2Field.Text = Format(SYN_FPA, "######0.##"): elemField.appendChild elem2Field
                    Set elem2Field = docStock.createElement("totalWithheldAmount"): elem2Field.Text = "0.00": elemField.appendChild elem2Field
                            
                    Set elem2Field = docStock.createElement("totalFeesAmount"): elem2Field.Text = "0.00": elemField.appendChild elem2Field
                    Set elem2Field = docStock.createElement("totalStampDutyAmount"): elem2Field.Text = "0.00": elemField.appendChild elem2Field
                    Set elem2Field = docStock.createElement("totalOtherTaxesAmount"): elem2Field.Text = "0.00": elemField.appendChild elem2Field
                            
                    Set elem2Field = docStock.createElement("totalDeductionsAmount"): elem2Field.Text = "0.00": elemField.appendChild elem2Field
                            
                    Set elem2Field = docStock.createElement("totalGrossValue"): elem2Field.Text = Format(SYN_KAU + SYN_FPA, "######0.##"): elemField.appendChild elem2Field
                            
                    Set elem2Field = docStock.createElement("incomeClassification") ' δημιουργω εσοχ' δημιουργω εσοχ
                    Set elem3Field = docStock.createElement("n1:classificationType"): elem3Field.Text = Split(ctypos, ";")(1): elem2Field.appendChild elem3Field
                    Set elem3Field = docStock.createElement("n1:classificationCategory"): elem3Field.Text = Split(ctypos, ";")(2): elem2Field.appendChild elem3Field
                    Set elem3Field = docStock.createElement("n1:amount"): elem3Field.Text = Format(SYN_KAU, "######0.##"): elem2Field.appendChild elem3Field
                    elemField.appendChild elem2Field
                            
                    .appendChild elemField
                    
                End With  '--------------------------------------

                .appendChild INVOICE
                
                EGGTIM.Close
                
                sqldt.MoveNext
            Loop

            ' Next
        End With ' /////////////////////////////////////////////////////////////////////////////////////////////

        Set .documentElement = elemRoot

        On Error Resume Next

        Kill "C:\txtfiles\inv.xml"

        On Error GoTo 0
        
        FormatXmlDocument docStock ' βαζει κενα να ειναι ευκολο στο διαβασμα
        
        .save "C:\txtfiles\inv.xml"
    End With

End Function
        
'        Sub TEST22()
'
'
'        Dim I As Integer
'        Dim sumNet As Single
'        Dim sumFpa As Single
'        Dim ctypos As String
'
'        '======================================= ΠΕΡΠΑΤΑΩ ΤΟ ΤΙΜ ====================================
'        For I = 0 To sqldt.Rows.Count - 1
'
'
'
'            If checkIntegrity(I) = False Then
'
'                MsgBox (" ΠΡΟΒΛΗΜΑ ΣΤΟ ΠΑΡΑΣΤΑΤΙΚΟ " + sqldt(0)("ATIM").ToString)
'            End If
'
'
'
'            Dim EGGTIM As New DataTable
'            Me.Text = "ΠΑΡΑΣΤΑΤΙΚΑ " + Str(I)
'            ExecuteSQLQuery("SELECT KODE,POSO,TIMM,EKPT,FPA,ISNULL(KAU_AJIA,0) AS KAU_AJIA,ISNULL(MIK_AJIA,0) AS MIK_AJIA FROM EGGTIM WHERE TIMM<>0 AND POSO<>0 AND ID_NUM=" +SQLDT("ID_NUM").ToString, EGGTIM)
'            Dim DUM As New DataTable
'            ExecuteSQLQuery("UPDATE TIM SET ENTITY=" + Str(i + 1) + " WHERE ID_NUM=" +SQLDT("ID_NUM").ToString, DUM)
'
'
'
'
'            sumNet = sqldt("aj1") + sqldt("aj2") + sqldt("aj3") + sqldt("aj4") + sqldt("aj5") + sqldt("aj6") + sqldt("aj7")
'            sumFpa = sqldt("fpa1") + sqldt("fpa2") + sqldt("fpa3") + sqldt("fpa4") + sqldt("fpa6") + sqldt("fpa7")
'
'            '1.1;E3_561_001;category1_1
'            ctypos = FINDTYPOS(Mid(sqldt(I)("ATIM"), 1, 1)) ' Split(tmpStr, ":")(0)
'            If Len(Trim(Split(ctypos, ";")(1))) = 0 Or Len(Trim(Split(ctypos, ";")(2))) = 0 Then
'                writer.Close()
'                MsgBox ("δεν εχουν ορισθει παραμετροι ΜΥDATA στο παρ/κό " + sqldt("ATIM"))
'                Exit Sub
'
'            End If
'
'
'            writer.WriteComment (sqldt(I)("ATIM") + " " + Format(sqldt(I)("HME"), "dd/MM/yyyy"))
'            writer.WriteStartElement ("invoice")
'            crNode("uid", Str(i + 1), writer)
'            'crNode("mark", "", writer)
'
'
'            '---------------------------------------------εκδοτης
'            writer.WriteStartElement ("issuer")
'            crNode("vatNumber", "028783755", writer)
'            crNode("country", "GR", writer)
'            crNode("branch", "0", writer)
'            ' writer.WriteStartElement("address")
'            ' crNode("postalCode", """66100""", writer)
'            ' crNode("city", """ΔΡΑΜΑ""", writer)
'            ' writer.WriteEndElement() '/address
'            writer.WriteEndElement() '/issuer
'
'            '--------------------------------------------- πελατης
'            If Mid(Split(ctypos, ";")(0), 1, 2) = "11" Then
'                'lianikh den xreiazetai pelaths
'            Else
'                ' End If
'                writer.WriteStartElement ("counterpart")
'                crNode("vatNumber", Trim(sqlDT(i)("AFM")), writer)  ' crNode("vatNumber", "026677115", writer)
'                crNode("country", "GR", writer)
'                crNode("branch", "0", writer)
'                writer.WriteStartElement ("address")
'                crNode("postalCode", """66100""", writer)  'crNode("postalCode", """66100""", writer)
'                crNode("city",SQLDT("POL"), writer)  ' crNode("city", """ΔΡΑΜΑ""", writer)
'                writer.WriteEndElement() ' /address
'                writer.WriteEndElement() ' /counterpart
'            End If
'
'
'            '----------------------------------------------- header
'            writer.WriteStartElement ("invoiceHeader")
'            crNode("series", "0", writer)
'            crNode("aa", Mid(sqlDT(i)("ATIM"), 2, 6), writer)   '  crNode("aa", "15", writer)
'            crNode("issueDate", Format(sqlDT(i)("hme"), "yyyy-MM-dd"), writer) ' crNode("issueDate", "2019-12-15", writer)
'
'
'            'ctypos = FINDTYPOS(Mid(sqlDT(i)("ATIM"), 1, 1)) ' Split(tmpStr, ":")(0)
'            'If Len(Trim(Split(ctypos, ";")(1))) = 0 Or Len(Trim(Split(ctypos, ";")(2))) = 0 Then
'            '    writer.Close()
'            '    MsgBox("δεν εχουν ορισθει παραμετροι ΜΥDATA στο παρ/κό " +SQLDT("ATIM"))
'            '    Exit Function
'
'            'End If
'
'
'            crNode("invoiceType", Split(ctypos, ";")(0), writer)   ' ειδος παραστατικού
'            crNode("currency", "EUR", writer)
'            crNode("exchangeRate", "1.0", writer)
'            writer.WriteEndElement() ' /invoiceHeader
'
'            '          <paymentMethods>
'            '   <paymentMethodDetails>
'            '       <type>3</type>
'            '       <amount>66.53</amount>
'            '   </paymentMethodDetails>
'            '</paymentMethods>
'            '----------------------------------------------- paymentMethods
'            writer.WriteStartElement ("paymentMethods")
'            writer.WriteStartElement ("paymentMethodDetails")
'            Dim cTrp As String = FindTRP(Mid(sqlDT(i)("TRP"), 1, 1))
'            If Len(cTrp) = 0 Then
'                MsgBox ("ΔΕΝ ΕΧΩ ΑΝΤΙΣΤΟΙΧΙΣΗ ΣΤΟΝ ΤΡΟΠΟ ΠΛΗΡΩΜΗΣ " + sqldt("TRP"))
'                writer.Close()
'                Exit Sub
'            End If
'            crNode("type", cTrp, writer)
'            crNode("amount", Format(sumNet + sumFpa, "######0.##"), writer)   '  crNode("aa", "15", writer)
'
'            writer.WriteEndElement() ' /paymentMethodDetails
'            writer.WriteEndElement() ' /paymentMethods
'
'
'            Dim SYN_KAU, SYN_FPA As Double
'            SYN_KAU = 0
'            SYN_FPA = 0
'            Dim fpaRow As Double
'
'            '======================================= ΠΕΡΠΑΤΑΩ ΤΟ EGGΤΙΜ ====================================
'            For L As Integer = 0 To EGGTIM.Rows.Count - 1
'
'                Dim AJ As Single
'                If IsDBNull(EGGTIM(L)("TIMM")) Then
'                    AJ = 0
'                Else
'                    AJ = EGGTIM(L)("KAU_AJIA")  ' Math.Round(EGGTIM(L)("POSO") * EGGTIM(L)("TIMM") * (1 - EGGTIM(L)("EKPT") / 100), 2)
'                End If
'
'                Dim VAT As String
'                '1 ΦΠΑ συντελεστής 24% 24%
'                '2 ΦΠΑ συντελεστής 13% 13%
'                '3 ΦΠΑ συντελεστής 6% 6%
'                '4 ΦΠΑ συντελεστής 17% 17%
'                '5 ΦΠΑ συντελεστής 9% 9%
'                '6 ΦΠΑ συντελεστής 4% 4%
'                '7 ’νευ Φ.Π.Α. 0%
'                '8 Εγγραφές χωρίς ΦΠΑ  (πχ Μισθοδοσία, Αποσβέσεις)
'
'
'
'                If EGGTIM(L)("FPA") = 1 Then '13%
'                    VAT = "2"
'                    SYN_KAU = SYN_KAU + AJ
'                    fpaRow = EGGTIM(L)("MIK_AJIA") - EGGTIM(L)("KAU_AJIA") ' AJ * 0.13
'                    SYN_FPA = SYN_FPA + fpaRow
'
'                    'ElseIf EGGTIM(L)("FPA") = 2 Then
'                    '   VAT = "1"
'                ElseIf EGGTIM(L)("FPA") = 2 Then
'                    VAT = "1"
'                    SYN_KAU = SYN_KAU + AJ
'                    fpaRow = EGGTIM(L)("MIK_AJIA") - EGGTIM(L)("KAU_AJIA") 'AJ * 0.24
'                    SYN_FPA = SYN_FPA + fpaRow
'
'                    ' SYN_FPA = SYN_FPA + AJ * 0.24
'
'                ElseIf EGGTIM(L)("FPA") = 5 Then
'                    VAT = "7"
'                    SYN_KAU = SYN_KAU + AJ
'                    fpaRow = 0
'                    SYN_FPA = SYN_FPA + fpaRow
'
'                ElseIf EGGTIM(L)("FPA") = 6 Then
'                    VAT = "1"
'                    SYN_KAU = SYN_KAU + AJ
'                    ' SYN_FPA = SYN_FPA + AJ * 0.24
'
'                    fpaRow = EGGTIM(L)("MIK_AJIA") - EGGTIM(L)("KAU_AJIA")  ' AJ * 0.24
'                    SYN_FPA = SYN_FPA + fpaRow
'
'
'
'
'                ElseIf EGGTIM(L)("FPA") = 4 Then
'                    VAT = "4"
'                    SYN_KAU = SYN_KAU + AJ
'                    fpaRow = EGGTIM(L)("MIK_AJIA") - EGGTIM(L)("KAU_AJIA")  ' AJ * 0.06
'                    SYN_FPA = SYN_FPA + fpaRow
'                Else ' If EGGTIM(L)("FPA") = 2 Then
'                    VAT = "1"
'                End If
'                '-----------------------------------------------  invoiceDetails
'                writer.WriteStartElement ("invoiceDetails")
'                crNode("lineNumber", Str(L + 1), writer) '  crNode("lineNumber", "1", writer)
'                '    crNode("quantity", EGGTIM(L)("POSO").ToString, writer)
'                '    crNode("measurementUnit", "1", writer)
'                crNode("netValue", Format(AJ, "######0.##"), writer)  ' crNode("netValue", "100", writer)
'
'                crNode("vatCategory", VAT, writer) '1=24%   2=13%
'
'                crNode("vatAmount", Format(fpaRow, "######0.##"), writer)  ' c
'
'                If fpaRow = 0 Then
'
'                    crNode("vatExemptionCategory", "1", writer) 'APALAGIFPA
'
'                End If
'
'                writer.WriteStartElement ("incomeClassification")
'                crNode("N1:classificationType", Split(ctypos, ";")(1), writer)
'                crNode("N1:classificationCategory", Split(ctypos, ";")(2), writer)
'                crNode("N1:amount", Format(AJ, "######0.##"), writer)
'
'                writer.WriteEndElement() '/incomeClassification
'
'
'
'
'
'                '                <invoiceDetails>
'                '  <lineNumber> 1</lineNumber>
'                '  <netValue>1185</netValue>
'                '  <vatCategory>7</vatCategory>
'                '  <vatAmount>0</vatAmount>
'                '<vatExemptionCategory>1</vatExemptionCategory>
'                '  <incomeClassification>
'                '    <N1:classificationType> E3_561_001</N1:classificationType>
'                '    <N1:classificationCategory>category1_1</N1:classificationCategory>
'                '    <N1:amount> 1185</N1:amount>
'                '  </incomeClassification>
'
'                '</invoiceDetails>
'
'
'
'                '               <incomeClassification>
'                '<N1:classificationType> E3_561_001</N1:classificationType>
'                '            <N1:classificationCategory>category1_1</N1:classificationCategory>
'                '<N1:amount> 100.0</N1:amount>
'                '    </incomeClassification>
'
'
'                writer.WriteEndElement()   ' /invoiceDetails
'            Next
'
'            ExecuteSQLQuery "UPDATE TIM SET AADEKAU=" + Replace(Format(SYN_KAU, "######0.#####"), ",", ".") + ",AADEFPA=" + Replace(Format(SYN_FPA, "######0.#####"), ",", ".") + " WHERE ID_NUM=" + sqldt("ID_NUM").ToString, DUM
'            '------------------------------------------------ InvoiceSummary
'            writer.WriteStartElement ("invoiceSummary")
'            crNode("totalNetValue", Format(SYN_KAU, "######0.##"), writer)  ' crNode("totalNetValue", "100", writer)
'            crNode("totalVatAmount", Format(SYN_FPA, "######0.##"), writer)  '  crNode("totalVatAmount", "24", writer)
'            crNode("totalWithheldAmount", "0", writer)
'            crNode("totalFeesAmount", "0", writer)
'            crNode("totalStampDutyAmount", "0", writer)
'            crNode("totalOtherTaxesAmount", "0", writer)
'            crNode("totalDeductionsAmount", "0", writer)
'            crNode("totalGrossValue", Format(SYN_KAU + SYN_FPA, "######0.##"), writer)
'
'
'            writer.WriteStartElement ("incomeClassification")
'            crNode("N1:classificationType", Split(ctypos, ";")(1), writer)
'            crNode("N1:classificationCategory", Split(ctypos, ";")(2), writer)
'            crNode("N1:amount", Format(SYN_KAU, "######0.##"), writer)
'            writer.WriteEndElement() '  /invoicesummary
'
'
'
'
'            writer.WriteEndElement() '  /invoicesummary
'            '=========================================================
'            writer.WriteEndElement() ' / Invoice
'
'        Next
'
'
'
'
'
'        writer.WriteEndElement() 'InvoicesDoc
'
'
'
'
'
'
'
'
'        writer.WriteEndDocument()
'        writer.Close()
'        '  MsgBox("ok")
'
'
'        ListBox2.Items.Clear()
'
'
'
'        '------ τοπικος ελεγχος xml που τον καταργησα γιατι μπηκε και το "https://www.aade.gr/myDATA/incomeClassificaton/v1.0
'        'FileOpen(1, "C:\TXTFILES\CHECKXSD.TXT", OpenMode.Output)
'        ''        Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'        'Dim myDocument As New XmlDocument
'        'myDocument.Load(ff) ' m_filename)  ' "C:\somefile.xml"
'        'myDocument.Schemas.Axmlns:N1dd("http://www.aade.gr/myDATA/invoice/v1.0", "c:\txtfiles\invoicesDoc-v0.6.xsd") 'namespace here or empty string
'        'Dim eventHandler As ValidationEventHandler = New ValidationEventHandler(AddressOf ValidationEventHandler)
'        'myDocument.Validate(eventHandler)
'        ''       MsgBox("ok ελεγχος")
'        'For n As Integer = 0 To ListBox2.Items.Count - 1
'        '    PrintLine(1, ListBox2.Items(n).ToString)
'        'Next
'        'FileClose(1)
'
'
'
'
'
'
'        paint_ergasies(DataGridView1, "SELECT   ATIM,HME,ENTITY,AADEKAU,AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7 AS KAUTIM,AADEFPA,FPA1+FPA2+FPA3+FPA4+FPA6+FPA7 AS FPATIM,ENTITYUID,ENTITYMARK FROM TIM WHERE ENTITY>0")
'
'    End Sub

Sub ExecuteSQLQuery(SQL As String, ByRef RR As ADODB.Recordset)
    If InStr(UCase(SQL), "SELECT") > 0 Then
        RR.open SQL, gdb, adOpenDynamic, adLockOptimistic
    
    Else  ' EXECUTE
    
        gdb.Execute SQL
    
    End If
    



End Sub
Function Get_AJ_ASCII(ByRef pol As String, ByVal polepis As String, ByVal ago As String, ByVal AGOEPIS As String) As Boolean

        '<EhHeader>


        '</EhHeader>



        ' Dim R As New ADODB.Recordset
        Dim x As String

        'If gConnect = "Access" Then
        '   Set db = OpenDatabase(gDir, False, False)
        'Else
        '   Set db = OpenDatabase(gDir, False, False, gConnect)
        'End If
Dim sqlDT2 As New ADODB.Recordset
       sqlDT2.open "select POL,EIDOS,AJIA_APOU from PARASTAT", gdb, adOpenDynamic, adLockOptimistic

        pol = " "

        Dim row As Integer
        Do While Not sqlDT2.EOF
           ' For row = 0 To sqlDT2.Rows.Count - 1

            If IsNull(sqlDT2("eidos")) Or IsNull(sqlDT2("pol")) Or IsNull(sqlDT2("ajia_apou")) Then

            Else

                If sqlDT2("pol") = "1" And sqlDT2("ajia_apou") = "3" Then
                    pol = pol + "'" + sqlDT2("eidos") + "',"
                End If

                If sqlDT2("pol") = "1" And sqlDT2("ajia_apou") = "4" Then
                    polepis = polepis + "'" + sqlDT2("eidos") + "',"
                End If

                If sqlDT2("pol") = "2" And sqlDT2("ajia_apou") = "1" Then
                    ago = ago + "'" + sqlDT2("eidos") + "',"
                End If

                If sqlDT2("pol") = "2" And sqlDT2("ajia_apou") = "2" Then
                    AGOEPIS = AGOEPIS + "'" + sqlDT2("eidos") + "',"
                End If

DoEvents


            End If
            sqlDT2.MoveNext
        Loop  'Next

240:    pol = Mid(pol, 1, Len(pol) - 1)

250:    If Len(polepis) > 0 Then
260:        polepis = Mid(polepis, 1, Len(polepis) - 1)
        Else
270:        polepis = "999"  'ME KENO DHMIOYRGEI PROBLHMA
        End If

280:
290:    ago = Mid(ago, 1, Len(ago) - 1)
300:    Get_AJ_ASCII = True

350:    If Len(AGOEPIS) > 0 Then
360:        AGOEPIS = Mid(AGOEPIS, 1, Len(AGOEPIS) - 1)
        Else
370:        AGOEPIS = "999" 'ME KENO DHMIOYRGEI PROBLHMA
        End If


    End Function

Private Sub Form_Load()
  gdb.open "DSN=MERCSQL;DATABASE=EMPMYDATA"
End Sub



  Function FINDTYPOS(C As String) As String
        Dim sqlDT4 As New ADODB.Recordset
        sqlDT4.open "select ETIK from PARASTAT where EIDOS='" + C + "'", gdb, adOpenDynamic, adLockOptimistic
        If IsNull(sqlDT4(0)) Then
            FINDTYPOS = ""
        Else
            FINDTYPOS = sqlDT4(0)
        End If


        FINDTYPOS = Trim(FINDTYPOS)

        ' If InStr(FINDTYPOS, ";") = 0 Then
        FINDTYPOS = FINDTYPOS + ";;" ' gia na mhn skaei to split()
        ' Else


        ' End If


    End Function
    Function FindTRP(C As String) As String
        Dim sqlDT4 As New ADODB.Recordset
        sqlDT4.open "select N1 from PINAKES where TYPOS=12 AND  AYJON=" + C + "", gdb, adOpenDynamic, adLockOptimistic
        If IsNull(sqlDT4(0)) Then
            FindTRP = ""
        Else
            FindTRP = sqlDT4(0)
        End If

    End Function

