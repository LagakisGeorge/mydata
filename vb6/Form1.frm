VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "download"
      Height          =   735
      Left            =   1560
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "upload"
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'BASISMENO STH ROUTINA "vb6.txt" POY BRHKA STO https://mydata-dev.portal.azure-api.net/issues/5f1eb1a2c7573053a44ddec9

upload
End Sub





Sub upload()
'BALTE EDO TO USERNAME KAI TO SUBSCRIPTION KEY SAS
p_USER = "XXXXXXXX"
p_Key = "XXXXXXXXXXXXXXXXXXXXXX"

'KALESTE THN YPOROUTINA GIA NA FTIAXETE TO XML POY 8A STEILETE STO myData
Call create_xml(tXML)

    Const URL1 = "https://mydata-dev.azure-api.net/SendInvoices"

    'initialize
    Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
    Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
    XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
    'force TLS 1.2
    XMLServer.Option(9) = 2048
    XMLServer.Option(6) = True
    
    XMLServer.Open "POST", URL1, False
    XMLServer.setRequestHeader "aade-user-id", "glagakis2"
    XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", "555bc57c80634243958f62b629316aaa"
    XMLServer.send tXML
    Dim v As String
    Debug.Print XMLServer.Status
   
    v = XMLServer.responseText
 Debug.Print v

End Sub

Sub request()
'RequestDocs (GET) =========================

    Const URL2 = "https://mydata-dev.azure-api.net/RequestTransmittedDocs"  '/RequestDocs"

   'initialize
    Set XMLServer = CreateObject("WinHttp.WinHttpRequest.5.1")
    Set XMLReceive = CreateObject("Msxml2.DOMDocument.6.0")
    XMLServer.setTimeouts 5000, 60000, 10000, 10000
    
    'force TLS 1.2
    XMLServer.Option(9) = 2048
    XMLServer.Option(6) = True
    
    XMLServer.Open "GET", URL2 & "?mark=400000019698028", False
    XMLServer.setRequestHeader "aade-user-id", "glagakis2"
    XMLServer.setRequestHeader "Ocp-Apim-Subscription-Key", "555bc57c80634243958f62b629316aaa"
    XMLServer.send ""

    Debug.Print XMLServer.Status
    '  debugPrint XMLServer.responseText
 v = XMLServer.responseText
 Debug.Print v
    
    
End Sub


'SUBROUTINA GIA NA FTIAXETE TO XML
Sub create_xml(myData)

myData = ""
myData = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"
myData = myData + "<InvoicesDoc xmlns=""http://www.aade.gr/myDATA/invoice/v1.0"" xsi:schemaLocation=""http://www.aade.gr/myDATA/invoice/v1.0 schema.xsd"" xmlns:icls=""https://www.aade.gr/myDATA/incomeClassificaton/v1.0"" xmlns:ecls=""https://www.aade.gr/myDATA/expensesClassificaton/v1.0"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">"
myData = myData + "  <invoice>"
myData = myData + "    <issuer>"
myData = myData + "      <vatNumber>123456789</vatNumber>"
myData = myData + "      <country>GR</country>"
myData = myData + "      <branch>0</branch>"
myData = myData + "    </issuer>"
myData = myData + "    <counterpart>"
myData = myData + "      <vatNumber>012345678</vatNumber>"
myData = myData + "      <country>GR</country>"
myData = myData + "      <branch>0</branch>"
myData = myData + "      <address>"
myData = myData + "        <street>ODOS PELATH</street>"
myData = myData + "        <number>11</number>"
myData = myData + "        <postalCode>11111</postalCode>"
myData = myData + "        <city>PERIOXH PELATH</city>"
myData = myData + "      </address>"
myData = myData + "    </counterpart>"
myData = myData + "    <invoiceHeader>"
myData = myData + "      <series>TIM</series>"
myData = myData + "      <aa>000152</aa>"
myData = myData + "      <issueDate>2020-09-30</issueDate>"
myData = myData + "      <invoiceType>1.1</invoiceType>"
myData = myData + "      <currency>EUR</currency>"
myData = myData + "    </invoiceHeader>"
myData = myData + "    <paymentMethods>"
myData = myData + "      <paymentMethodDetails>"
myData = myData + "        <type>5</type>"
myData = myData + "        <amount>31.00</amount>"
myData = myData + "        </paymentMethodDetails>"
myData = myData + "    </paymentMethods>"
myData = myData + "    <invoiceDetails>"
myData = myData + "      <lineNumber>10001</lineNumber>"
myData = myData + "      <netValue>25.00</netValue>"
myData = myData + "      <vatCategory>1</vatCategory>"
myData = myData + "      <vatAmount>6.00</vatAmount>"
myData = myData + "      <incomeClassification>"
myData = myData + "        <icls:classificationType>E3_561_001</icls:classificationType>"
myData = myData + "        <icls:classificationCategory>category1_1</icls:classificationCategory>"
myData = myData + "        <icls:amount>25.00</icls:amount>"
myData = myData + "        <icls:id>1</icls:id>"
myData = myData + "      </incomeClassification>"
myData = myData + "    </invoiceDetails>"
myData = myData + "    <invoiceSummary>"
myData = myData + "      <totalNetValue>25.00</totalNetValue>"
myData = myData + "      <totalVatAmount>6.00</totalVatAmount>"
myData = myData + "      <totalWithheldAmount>0.00</totalWithheldAmount>"
myData = myData + "      <totalFeesAmount>0.00</totalFeesAmount>"
myData = myData + "      <totalStampDutyAmount>0.00</totalStampDutyAmount>"
myData = myData + "      <totalOtherTaxesAmount>0.00</totalOtherTaxesAmount>"
myData = myData + "      <totalDeductionsAmount>0.00</totalDeductionsAmount>"
myData = myData + "      <totalGrossValue>31.00</totalGrossValue>"
myData = myData + "      <incomeClassification>"
myData = myData + "        <icls:classificationType>E3_561_001</icls:classificationType>"
myData = myData + "        <icls:classificationCategory>category1_1</icls:classificationCategory>"
myData = myData + "        <icls:amount>25.00</icls:amount>"
myData = myData + "        <icls:id>1</icls:id>"
myData = myData + "      </incomeClassification>"
myData = myData + "    </invoiceSummary>"
myData = myData + "  </invoice>"
myData = myData + "</InvoicesDoc>"

End Sub

Private Sub Command2_Click()
 request
End Sub
