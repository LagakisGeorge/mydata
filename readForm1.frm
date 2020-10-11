VERSION 5.00
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
    Dim docStock As MSXML2.DOMDocument
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
   ToXMLsub
   
End Sub
Private Sub ToXMLsub()

        '===ΒΓΑΖΩ ΤΟ XML ΓΙΑ ΤΑ ΠΑΡΑΣΤΑΤΙΚΑ =================================================================================
        'WHERE (ENTITYMARK IS NULL OR ENTITYMARK='ERROR' ) AND
        'Left(ATIM, 1) In     (  " + PAR + "  )    And
        'HME>='" + Format(APO.Value, "MM/dd/yyyy") + "'  AND HME<='" + Format(EOS.Value, "MM/dd/yyyy") + "'  "
        '<correlatedInvoices>400000017716190</correlatedInvoices>
       
Dim R As New ADODB.Recordset


        Dim pol, polepis, ago, AGOEPIS As String


       
        ExecuteSQLQuery "UPDATE TIM SET ENTITY=0,ENTLINEN=0", R

      '  Get_AJ_ASCII pol, polepis, ago, AGOEPIS

        Dim PAR, SYNT As String
        PAR = pol + polepis
        Dim SQL As String
        SYNT = ""
        SQL = "SELECT  ID_NUM, AJ1  ,AJ2 , AJ3,AJ4,AJ5,AJI,FPA1,FPA2,FPA3,FPA4,ATIM,"
        SQL = SQL + "HME,PEL.EPO,PEL.AFM,KPE,PEL.DIE,PEL.XRVMA"    '"CONVERT(CHAR(10),HME,3) AS HMEP
        SQL = SQL + ",PEL.EPA,PEL.POL,AJ6,FPA6,AJ7,FPA7,TRP,APALAGIFPA "

        SQL = SQL + "   FROM TIM INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD "
        SQL = SQL + " WHERE (ENTITYMARK IS NULL OR ENTITYMARK='ERROR' ) AND    LEFT(ATIM,1) IN     (  " + PAR + "  )    and HME>='" + Format(APO.Value, "MM/dd/yyyy") + "'  AND HME<='" + Format(EOS.Value, "MM/dd/yyyy") + "'  "
        SQL = SQL + "  AND AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7>0  " + SYNT
        SQL = SQL + " order by HME"       '  OR INCMARK IS NULL OR INCMARK='ERROR'




        '  SQL = "SELECT  top 20  AJ1 ,AJ2  from TIM  order by HME"

        ExecuteSQLQuery SQL, R

        If sqlDT.Rows.Count = 0 Then
            MsgBox ("ΔΕΝ ΒΡΕΘΗΚΑΝ ΕΓΓΡΑΦΕΣ")
           ' ToXMLsub = 0
            Exit Sub
        End If



        Dim ff As String
        ff = "c:\txtfiles\inv.xml" 'c:\mercvb\m" + Format(Now, "yyyyddmmHHMM") + ".export" ' "\\Logisthrio\333\pr.export" '
        Dim writer As New MXXMLWriter
        writer = MXXMLWriter(ff, System.Text.encoding.UTF8)
        writer.WriteStartDocument (True)
        writer.Formatting = Formatting.Indented
        writer.Indentation = 2
        writer.WriteStartElement ("InvoicesDoc")
        writer.WriteAttributeString "xmlns", "http://www.aade.gr/myDATA/invoice/v1.0"
        writer.WriteAttributeString "xsi:schemaLocation", "http://www.aade.gr/myDATA/invoice/v1.0 schema.xsd"
        writer.WriteAttributeString "xmlns:N1", "https://www.aade.gr/myDATA/incomeClassificaton/v1.0"
        writer.WriteAttributeString "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"

        writer.WriteEndDocument
        writer.Close
        
        End Sub
        
        Sub TEST22()
        

        Dim I As Integer
        Dim sumNet As Single
        Dim sumFpa As Single
        Dim ctypos As String

        '======================================= ΠΕΡΠΑΤΑΩ ΤΟ ΤΙΜ ====================================
        For I = 0 To sqlDT.Rows.Count - 1



            If checkIntegrity(I) = False Then

                MsgBox (" ΠΡΟΒΛΗΜΑ ΣΤΟ ΠΑΡΑΣΤΑΤΙΚΟ " + sqlDT(0)("ATIM").ToString)
            End If



            Dim EGGTIM As New DataTable
            Me.Text = "ΠΑΡΑΣΤΑΤΙΚΑ " + Str(I)
            ExecuteSQLQuery("SELECT KODE,POSO,TIMM,EKPT,FPA,ISNULL(KAU_AJIA,0) AS KAU_AJIA,ISNULL(MIK_AJIA,0) AS MIK_AJIA FROM EGGTIM WHERE TIMM<>0 AND POSO<>0 AND ID_NUM=" + sqlDT(i)("ID_NUM").ToString, EGGTIM)
            Dim DUM As New DataTable
            ExecuteSQLQuery("UPDATE TIM SET ENTITY=" + Str(i + 1) + " WHERE ID_NUM=" + sqlDT(i)("ID_NUM").ToString, DUM)




            sumNet = sqlDT(I)("aj1") + sqlDT(I)("aj2") + sqlDT(I)("aj3") + sqlDT(I)("aj4") + sqlDT(I)("aj5") + sqlDT(I)("aj6") + sqlDT(I)("aj7")
            sumFpa = sqlDT(I)("fpa1") + sqlDT(I)("fpa2") + sqlDT(I)("fpa3") + sqlDT(I)("fpa4") + sqlDT(I)("fpa6") + sqlDT(I)("fpa7")

            '1.1;E3_561_001;category1_1
            ctypos = FINDTYPOS(Mid(sqlDT(I)("ATIM"), 1, 1)) ' Split(tmpStr, ":")(0)
            If Len(Trim(Split(ctypos, ";")(1))) = 0 Or Len(Trim(Split(ctypos, ";")(2))) = 0 Then
                writer.Close()
                MsgBox ("δεν εχουν ορισθει παραμετροι ΜΥDATA στο παρ/κό " + sqlDT(I)("ATIM"))
                Exit Sub

            End If


            writer.WriteComment (sqlDT(I)("ATIM") + " " + Format(sqlDT(I)("HME"), "dd/MM/yyyy"))
            writer.WriteStartElement ("invoice")
            crNode("uid", Str(i + 1), writer)
            'crNode("mark", "", writer)


            '---------------------------------------------εκδοτης
            writer.WriteStartElement ("issuer")
            crNode("vatNumber", "028783755", writer)
            crNode("country", "GR", writer)
            crNode("branch", "0", writer)
            ' writer.WriteStartElement("address")
            ' crNode("postalCode", """66100""", writer)
            ' crNode("city", """ΔΡΑΜΑ""", writer)
            ' writer.WriteEndElement() '/address
            writer.WriteEndElement() '/issuer

            '--------------------------------------------- πελατης
            If Mid(Split(ctypos, ";")(0), 1, 2) = "11" Then
                'lianikh den xreiazetai pelaths
            Else
                ' End If
                writer.WriteStartElement ("counterpart")
                crNode("vatNumber", Trim(sqlDT(i)("AFM")), writer)  ' crNode("vatNumber", "026677115", writer)
                crNode("country", "GR", writer)
                crNode("branch", "0", writer)
                writer.WriteStartElement ("address")
                crNode("postalCode", """66100""", writer)  'crNode("postalCode", """66100""", writer)
                crNode("city", sqlDT(i)("POL"), writer)  ' crNode("city", """ΔΡΑΜΑ""", writer)
                writer.WriteEndElement() ' /address
                writer.WriteEndElement() ' /counterpart
            End If


            '----------------------------------------------- header
            writer.WriteStartElement ("invoiceHeader")
            crNode("series", "0", writer)
            crNode("aa", Mid(sqlDT(i)("ATIM"), 2, 6), writer)   '  crNode("aa", "15", writer)
            crNode("issueDate", Format(sqlDT(i)("hme"), "yyyy-MM-dd"), writer) ' crNode("issueDate", "2019-12-15", writer)


            'ctypos = FINDTYPOS(Mid(sqlDT(i)("ATIM"), 1, 1)) ' Split(tmpStr, ":")(0)
            'If Len(Trim(Split(ctypos, ";")(1))) = 0 Or Len(Trim(Split(ctypos, ";")(2))) = 0 Then
            '    writer.Close()
            '    MsgBox("δεν εχουν ορισθει παραμετροι ΜΥDATA στο παρ/κό " + sqlDT(i)("ATIM"))
            '    Exit Function

            'End If


            crNode("invoiceType", Split(ctypos, ";")(0), writer)   ' ειδος παραστατικού
            crNode("currency", "EUR", writer)
            crNode("exchangeRate", "1.0", writer)
            writer.WriteEndElement() ' /invoiceHeader

            '          <paymentMethods>
            '   <paymentMethodDetails>
            '       <type>3</type>
            '       <amount>66.53</amount>
            '   </paymentMethodDetails>
            '</paymentMethods>
            '----------------------------------------------- paymentMethods
            writer.WriteStartElement ("paymentMethods")
            writer.WriteStartElement ("paymentMethodDetails")
            Dim cTrp As String = FindTRP(Mid(sqlDT(i)("TRP"), 1, 1))
            If Len(cTrp) = 0 Then
                MsgBox ("ΔΕΝ ΕΧΩ ΑΝΤΙΣΤΟΙΧΙΣΗ ΣΤΟΝ ΤΡΟΠΟ ΠΛΗΡΩΜΗΣ " + sqlDT(I)("TRP"))
                writer.Close()
                Exit Sub
            End If
            crNode("type", cTrp, writer)
            crNode("amount", Format(sumNet + sumFpa, "######0.##"), writer)   '  crNode("aa", "15", writer)

            writer.WriteEndElement() ' /paymentMethodDetails
            writer.WriteEndElement() ' /paymentMethods


            Dim SYN_KAU, SYN_FPA As Double
            SYN_KAU = 0
            SYN_FPA = 0
            Dim fpaRow As Double

            '======================================= ΠΕΡΠΑΤΑΩ ΤΟ EGGΤΙΜ ====================================
            For L As Integer = 0 To EGGTIM.Rows.Count - 1

                Dim AJ As Single
                If IsDBNull(EGGTIM(L)("TIMM")) Then
                    AJ = 0
                Else
                    AJ = EGGTIM(L)("KAU_AJIA")  ' Math.Round(EGGTIM(L)("POSO") * EGGTIM(L)("TIMM") * (1 - EGGTIM(L)("EKPT") / 100), 2)
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



                If EGGTIM(L)("FPA") = 1 Then '13%
                    VAT = "2"
                    SYN_KAU = SYN_KAU + AJ
                    fpaRow = EGGTIM(L)("MIK_AJIA") - EGGTIM(L)("KAU_AJIA") ' AJ * 0.13
                    SYN_FPA = SYN_FPA + fpaRow

                    'ElseIf EGGTIM(L)("FPA") = 2 Then
                    '   VAT = "1"
                ElseIf EGGTIM(L)("FPA") = 2 Then
                    VAT = "1"
                    SYN_KAU = SYN_KAU + AJ
                    fpaRow = EGGTIM(L)("MIK_AJIA") - EGGTIM(L)("KAU_AJIA") 'AJ * 0.24
                    SYN_FPA = SYN_FPA + fpaRow

                    ' SYN_FPA = SYN_FPA + AJ * 0.24

                ElseIf EGGTIM(L)("FPA") = 5 Then
                    VAT = "7"
                    SYN_KAU = SYN_KAU + AJ
                    fpaRow = 0
                    SYN_FPA = SYN_FPA + fpaRow

                ElseIf EGGTIM(L)("FPA") = 6 Then
                    VAT = "1"
                    SYN_KAU = SYN_KAU + AJ
                    ' SYN_FPA = SYN_FPA + AJ * 0.24

                    fpaRow = EGGTIM(L)("MIK_AJIA") - EGGTIM(L)("KAU_AJIA")  ' AJ * 0.24
                    SYN_FPA = SYN_FPA + fpaRow




                ElseIf EGGTIM(L)("FPA") = 4 Then
                    VAT = "4"
                    SYN_KAU = SYN_KAU + AJ
                    fpaRow = EGGTIM(L)("MIK_AJIA") - EGGTIM(L)("KAU_AJIA")  ' AJ * 0.06
                    SYN_FPA = SYN_FPA + fpaRow
                Else ' If EGGTIM(L)("FPA") = 2 Then
                    VAT = "1"
                End If
                '-----------------------------------------------  invoiceDetails
                writer.WriteStartElement ("invoiceDetails")
                crNode("lineNumber", Str(L + 1), writer) '  crNode("lineNumber", "1", writer)
                '    crNode("quantity", EGGTIM(L)("POSO").ToString, writer)
                '    crNode("measurementUnit", "1", writer)
                crNode("netValue", Format(AJ, "######0.##"), writer)  ' crNode("netValue", "100", writer)

                crNode("vatCategory", VAT, writer) '1=24%   2=13%

                crNode("vatAmount", Format(fpaRow, "######0.##"), writer)  ' c

                If fpaRow = 0 Then

                    crNode("vatExemptionCategory", "1", writer) 'APALAGIFPA

                End If

                writer.WriteStartElement ("incomeClassification")
                crNode("N1:classificationType", Split(ctypos, ";")(1), writer)
                crNode("N1:classificationCategory", Split(ctypos, ";")(2), writer)
                crNode("N1:amount", Format(AJ, "######0.##"), writer)

                writer.WriteEndElement() '/incomeClassification





                '                <invoiceDetails>
                '  <lineNumber> 1</lineNumber>
                '  <netValue>1185</netValue>
                '  <vatCategory>7</vatCategory>
                '  <vatAmount>0</vatAmount>
                '<vatExemptionCategory>1</vatExemptionCategory>
                '  <incomeClassification>
                '    <N1:classificationType> E3_561_001</N1:classificationType>
                '    <N1:classificationCategory>category1_1</N1:classificationCategory>
                '    <N1:amount> 1185</N1:amount>
                '  </incomeClassification>

                '</invoiceDetails>



                '               <incomeClassification>
                '<N1:classificationType> E3_561_001</N1:classificationType>
                '            <N1:classificationCategory>category1_1</N1:classificationCategory>
                '<N1:amount> 100.0</N1:amount>
                '    </incomeClassification>


                writer.WriteEndElement()   ' /invoiceDetails
            Next

            ExecuteSQLQuery "UPDATE TIM SET AADEKAU=" + Replace(Format(SYN_KAU, "######0.#####"), ",", ".") + ",AADEFPA=" + Replace(Format(SYN_FPA, "######0.#####"), ",", ".") + " WHERE ID_NUM=" + sqlDT(I)("ID_NUM").ToString, DUM
            '------------------------------------------------ InvoiceSummary
            writer.WriteStartElement ("invoiceSummary")
            crNode("totalNetValue", Format(SYN_KAU, "######0.##"), writer)  ' crNode("totalNetValue", "100", writer)
            crNode("totalVatAmount", Format(SYN_FPA, "######0.##"), writer)  '  crNode("totalVatAmount", "24", writer)
            crNode("totalWithheldAmount", "0", writer)
            crNode("totalFeesAmount", "0", writer)
            crNode("totalStampDutyAmount", "0", writer)
            crNode("totalOtherTaxesAmount", "0", writer)
            crNode("totalDeductionsAmount", "0", writer)
            crNode("totalGrossValue", Format(SYN_KAU + SYN_FPA, "######0.##"), writer)


            writer.WriteStartElement ("incomeClassification")
            crNode("N1:classificationType", Split(ctypos, ";")(1), writer)
            crNode("N1:classificationCategory", Split(ctypos, ";")(2), writer)
            crNode("N1:amount", Format(SYN_KAU, "######0.##"), writer)
            writer.WriteEndElement() '  /invoicesummary




            writer.WriteEndElement() '  /invoicesummary
            '=========================================================
            writer.WriteEndElement() ' / Invoice

        Next





        writer.WriteEndElement() 'InvoicesDoc








        writer.WriteEndDocument()
        writer.Close()
        '  MsgBox("ok")


        ListBox2.Items.Clear()



        '------ τοπικος ελεγχος xml που τον καταργησα γιατι μπηκε και το "https://www.aade.gr/myDATA/incomeClassificaton/v1.0
        'FileOpen(1, "C:\TXTFILES\CHECKXSD.TXT", OpenMode.Output)
        ''        Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Dim myDocument As New XmlDocument
        'myDocument.Load(ff) ' m_filename)  ' "C:\somefile.xml"
        'myDocument.Schemas.Axmlns:N1dd("http://www.aade.gr/myDATA/invoice/v1.0", "c:\txtfiles\invoicesDoc-v0.6.xsd") 'namespace here or empty string
        'Dim eventHandler As ValidationEventHandler = New ValidationEventHandler(AddressOf ValidationEventHandler)
        'myDocument.Validate(eventHandler)
        ''       MsgBox("ok ελεγχος")
        'For n As Integer = 0 To ListBox2.Items.Count - 1
        '    PrintLine(1, ListBox2.Items(n).ToString)
        'Next
        'FileClose(1)






        paint_ergasies(DataGridView1, "SELECT   ATIM,HME,ENTITY,AADEKAU,AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7 AS KAUTIM,AADEFPA,FPA1+FPA2+FPA3+FPA4+FPA6+FPA7 AS FPATIM,ENTITYUID,ENTITYMARK FROM TIM WHERE ENTITY>0")

    End Sub

Sub ExecuteSQLQuery(SQL As String, ByRef RR As ADODB.Recordset)
    If InStr(UCase(SQL), "SELECT") > 0 Then
        RR.Open SQL, gdb, adOpenDynamic, adLockOptimistic
    
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
        ExecuteSQLQuery "select POL,EIDOS,AJIA_APOU from PARASTAT", sqlDT2

        pol = " "

        Dim row As Integer
        For row = 0 To sqlDT2.Rows.Count - 1

            If IsDBNull(sqlDT2.Rows(row)("eidos")) Or IsDBNull(sqlDT2.Rows(row)("pol")) Or IsDBNull(sqlDT2.Rows(row)("ajia_apou")) Then

            Else

                If sqlDT2.Rows(row)("pol") = "1" And sqlDT2.Rows(row)("ajia_apou") = "3" Then
                    pol = pol + "'" + sqlDT2.Rows(row)("eidos") + "',"
                End If

                If sqlDT2.Rows(row)("pol") = "1" And sqlDT2.Rows(row)("ajia_apou") = "4" Then
                    polepis = polepis + "'" + sqlDT2.Rows(row)("eidos") + "',"
                End If

                If sqlDT2.Rows(row)("pol") = "2" And sqlDT2.Rows(row)("ajia_apou") = "1" Then
                    ago = ago + "'" + sqlDT2.Rows(row)("eidos") + "',"
                End If

                If sqlDT2.Rows(row)("pol") = "2" And sqlDT2.Rows(row)("ajia_apou") = "2" Then
                    AGOEPIS = AGOEPIS + "'" + sqlDT2.Rows(row)("eidos") + "',"
                End If




            End If
        Next

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
