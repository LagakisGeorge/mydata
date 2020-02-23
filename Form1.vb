Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Xml
Imports System.Data.OleDb
Imports System.Xml.Schema
Imports System.Data.SqlClient
Imports System.Web
'Imports System.Net.Http.Headers
'Imports System.Text
'Imports System.Net.Http

'ΑΛΛΑΓΕΣ ΣΤΟ ΤΙΜ 
'    [ENTITYUID] [varchar](40) NULL,
'	[ENTITYMARK] [varchar](43) NULL,
'	[ENTITY] [int] NULL,
'	[AADEKAU] [float] NULL,
'	[AADEFPA] [float] NULL,
'	[ENTLINEN] [int] NULL,
'	[INCMARK] [nvarchar](43) NULL,

Public Class Form1


    Public openedFileStream As System.IO.Stream
    Public dataBytes() As Byte
    Public gConnect As String
    Public gSQLCon As String
    Public sqlDT As New DataTable
    Public sqlDT2 As New DataTable
    ' Public Property HttpUtility As Object
    '  SELECT ENTITY ,ATIM,ENTITYUID,ENTITYMARK ,HME FROM TIM WHERE ENTITY>0 ORDER BY ENTITY'

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MakeRequest()

    End Sub


    Private Async Sub MakeRequest()
        ' ΣΤΕΛΝΩ ΣΤΟ SENDINVOICES 
        ' "c:\txtfiles\inv.xml").ToString ' "--> εκει έχω αποθηκεύσει το xml που εφτιαξα"
        ' H APANTHSH EINAI STO  "c:\txtfiles\apantSendInv.XML")




        Dim client = New HttpClient()
        'Dim queryString = HttpUtility.ParseQueryString(String.Empty)
        Try
            client.DefaultRequestHeaders.Add("aade-user-id", "glagakis2")
            client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", "555bc57c80634243958f62b629316aaa")

            Dim uri = "https://mydata-dev.azure-api.net/SendInvoices" ' + queryString.ToString

            Dim response As HttpResponseMessage
            Dim xl = XDocument.Load("c:\txtfiles\inv.xml").ToString ' "--> εκει έχω αποθηκεύσει το xml που εφτιαξα"
            Dim byteData As Byte() = Encoding.UTF8.GetBytes(xl)

            Using content = New ByteArrayContent(byteData)
                content.Headers.ContentType = New MediaTypeHeaderValue("application/xml")
                response = Await client.PostAsync(uri, content)
                Dim result = Await response.Content.ReadAsStringAsync()
                Dim MF = "c:\txtfiles\apantSendInv" + Format(Now, "yyyyddMMHHmm") + ".xml"
                FileOpen(1, MF, OpenMode.Output)
                PrintLine(1, result.ToString)
                FileClose(1)
                TextBox2.Text = result.ToString
                ' "είναι το textbox πανω στη φόρμα που σου επιστρέφει το response xml"
                'Dim byteData2 As Byte() = File.ReadAllBytes("c:\txtfiles\inv.xml")
                Rename("c:\txtfiles\inv.xml", "c:\txtfiles\inv" + Format(Now, "yyyyddMMHHmm") + ".xml")
                FileCopy(MF, "c:\txtfiles\apantSendInv.XML")

            End Using


        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        MakeIncomeRequest()
    End Sub

    Private Async Sub MakeIncomeRequest()
        'ΠΑΙΡΝΩ ΤΟ ΑΡΧΕΙΟ  c:\txtfiles\inc.xml"
        ' ΚΑΙ ΣΤΕΛΝΩ ΤΟ SendIncomeClassification" 
        ' Η ΑΠΑΝΤΗΣΗ ΕΙΝΑΙ ΤΟ "c:\txtfiles\apantiNCOMe" + Format(Now, "yyyyddMMHHmm") + ".xml"





        ListBox2.Items.Clear()

        Dim client = New HttpClient()
        'Dim queryString = HttpUtility.ParseQueryString(String.Empty)
        Try
            client.DefaultRequestHeaders.Add("aade-user-id", "glagakis2")
            client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", "555bc57c80634243958f62b629316aaa")

            Dim uri = "https://mydata-dev.azure-api.net/SendIncomeClassification" ' + queryString.ToString

            Dim response As HttpResponseMessage
            Dim xl = XDocument.Load("c:\txtfiles\inc.xml").ToString ' "--> εκει έχω αποθηκεύσει το xml που εφτιαξα"
            Dim byteData As Byte() = Encoding.UTF8.GetBytes(xl)

            Using content = New ByteArrayContent(byteData)
                content.Headers.ContentType = New MediaTypeHeaderValue("application/xml")
                response = Await client.PostAsync(uri, content)
                Dim result = Await response.Content.ReadAsStringAsync()

                TextBox2.Text = result.ToString
                ' "είναι το textbox πανω στη φόρμα που σου επιστρέφει το response xml"

                Dim MF = "c:\txtfiles\ApantIncome" + Format(Now, "yyyyddMMHHmm") + ".xml"
                FileOpen(1, MF, OpenMode.Output)
                PrintLine(1, result.ToString)
                FileClose(1)
                TextBox2.Text = result.ToString
                ' "είναι το textbox πανω στη φόρμα που σου επιστρέφει το response xml"
                'Dim byteData2 As Byte() = File.ReadAllBytes("c:\txtfiles\inv.xml")
                ' Rename("c:\txtfiles\inv.xml", "c:\txtfiles\inv" + Format(Now, "yyyyddMMHHmm") + ".xml")
                FileCopy(MF, "c:\txtfiles\ApantIncome.XML")

            End Using


        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


    End Sub


    Public Function checkServer(ByVal check_path As Integer) As Boolean
        Dim c As String
        Dim tmpStr As String
        c = "Config.ini"


        Dim par As String = ""
        Dim mf As String
        mf = c   ' "c:\mercvb\err3.txt"
        If Len(Dir(UCase(mf))) = 0 Then
            par = ":(local)\sql2012:sa:12345678:1:EMP"    '" 'G','g','Ξ','D'  "
            par = InputBox("ΒΑΣΗ ΔΕΔΟΜΕΝΩΝ", , par)
        Else
            FileOpen(1, mf, OpenMode.Input)
            '   Input(1, par)
            par = LineInput(1)
            FileClose(1)
        End If
        If check_path = 1 Then
            par = InputBox("ΒΑΣΗ ΔΕΔΟΜΕΝΩΝ  (CONFIG.INI ΣΤΟΝ ΤΡΕΧΟΝΤΑ ΦΑΚΕΛΟ) ", ":Π.Χ. (local)\sql2012:sa:12345678:1:EMP", par)
        End If

        'Input = InputBox("Text:")

        If String.IsNullOrEmpty(par) Then
            ' Cancelled, or empty
            checkServer = False
            ' MsgBox("εξοδος λογω μη σύνδεσης με βάση δεδομένων")
            Exit Function
        Else
            ' Normal
        End If


        FileOpen(1, mf, OpenMode.Output)
        PrintLine(1, par)
        FileClose(1)





        ':(local)\sql2012:::2:EMP
        ':(local)\sql2012:sa:12345678:1:EMP





        Try

            ' With FrmSERVERSETTINGS
            OpenFileDialog1.FileName = c
            openedFileStream = OpenFileDialog1.OpenFile()
            'End With

            ReDim dataBytes(openedFileStream.Length - 1) 'Init 
            openedFileStream.Read(dataBytes, 0, openedFileStream.Length)
            openedFileStream.Close()
            tmpStr = par ' System.Text.Encoding.Unicode.GetString(dataBytes)

            '     With FrmSERVERSETTINGS
            If Val(Split(tmpStr, ":")(4)) = 1 Then
                'network
                'gConnect = "Provider=SQLOLEDB.1;" & _
                '           "Data Source=" & Split(tmpStr, ":")(0) & _
                '           ";Network=" & Split(tmpStr, ":")(1) & _
                '           ";Server=" & Split(tmpStr, ":")(1) & _
                '           ";Initial Catalog=" & Trim(Split(tmpStr, ":")(5)) & _
                '           ";User Id=" & Split(tmpStr, ":")(2) & _
                '           ";Password=" & Split(tmpStr, ":")(3)

                gConnect = "Provider=SQLOLEDB.1;;Password=" & Split(tmpStr, ":")(3) &
                ";Persist Security Info=True ;" &
                ";User Id=" & Split(tmpStr, ":")(2) &
                ";Initial Catalog=" & Trim(Split(tmpStr, ":")(5)) &
                ";Data Source=" & Split(tmpStr, ":")(1)

                ''   gConnect = "Provider=SQLOLEDB.1;;Password=" & Split(tmpStr, ":")(3) &
                gSQLCon = "Server=" + Split(tmpStr, ":")(1)
                gSQLCon = gSQLCon + ";Database=" + Trim(Split(tmpStr, ":")(5))
                gSQLCon = gSQLCon + ";Uid=" + Split(tmpStr, ":")(2) + ";Pwd=" + Split(tmpStr, ":")(3)



            Else
                'local
                'MsgBox(Split(tmpStr, ":")(1))
                '  gConnect = "Provider=SQLOLEDB;Server=" & Split(tmpStr, ":")(1) &
                '         ";Database=" & Split(tmpStr, ":")(5) & "; Trusted_Connection=yes;"

                '    gConSQL = "Data Source=" & Split(tmpStr, ":")(1) & ";Integrated Security=True;database=" & Split(tmpStr, ":")(5)
                'cnString = "Data Source=localhost\SQLEXPRESS;Integrated Security=True;database=YGEIA"

            End If
            'End With
            Dim sqlCon As New OleDbConnection
            '
            ' gConnect = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;PWD=12345678;Initial Catalog=D2014;Data Source=logisthrio\sqlexpress"
            'GDB.Open(gConnect)



            'OK
            'gConnect = "Provider=SQLOLEDB.1;;Password=12345678;Persist Security Info=True ;User Id=sa;Initial Catalog=EMP;Data Source=LOGISTHRIO\SQLEXPRESS"
            sqlCon.ConnectionString = gConnect
            sqlCon.Open()
            checkServer = True
            sqlCon.Close()

            '            Dim GDB As New ADODB.Connection

        Catch ex As Exception
            checkServer = False
            MsgBox("εξοδος λογω μη σύνδεσης με βάση δεδομένων")
            'End
        End Try
    End Function
    Function Get_AJ_ASCII(ByRef pol As String,
                          ByVal polepis As String,
                          ByVal ago As String,
                          ByVal AGOEPIS As String) As Boolean

        '<EhHeader>


        '</EhHeader>



        ' Dim R As New ADODB.Recordset
        Dim x As String

        'If gConnect = "Access" Then
        '   Set db = OpenDatabase(gDir, False, False)
        'Else
        '   Set db = OpenDatabase(gDir, False, False, gConnect)
        'End If

        ExecuteSQLQuery("select POL,EIDOS,AJIA_APOU from PARASTAT", sqlDT2)

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




    Private Sub ToXML_Click(sender As Object, e As EventArgs) Handles toXML.Click
        '===ΒΓΑΖΩ ΤΟ XML ΓΙΑ ΤΑ ΠΑΡΑΣΤΑΤΙΚΑ =================================================================================
        'WHERE (ENTITYMARK IS NULL OR ENTITYMARK='ERROR' ) AND   
        'Left(ATIM, 1) In     (  " + PAR + "  )    And 
        'HME>='" + Format(APO.Value, "MM/dd/yyyy") + "'  AND HME<='" + Format(EOS.Value, "MM/dd/yyyy") + "'  "





        Dim pol, polepis, ago, AGOEPIS As String


        If checkServer(0) Then
            ' MsgBox("OK")
        End If
        ExecuteSQLQuery("UPDATE TIM SET ENTITY=0")

        Get_AJ_ASCII(pol, polepis, ago, AGOEPIS)

        Dim PAR, SYNT As String
        PAR = pol + polepis
        Dim SQL As String
        SYNT = ""
        SQL = "SELECT top 1 ID_NUM, AJ1  ,AJ2 , AJ3,AJ4,AJ5,AJI,FPA1,FPA2,FPA3,FPA4,ATIM,"
        SQL = SQL + "HME,PEL.EPO,PEL.AFM,KPE,PEL.DIE,PEL.XRVMA"    '"CONVERT(CHAR(10),HME,3) AS HMEP
        SQL = SQL + ",PEL.EPA,PEL.POL,AJ6,FPA6,AJ7,FPA7 ,ID_NUM"

        SQL = SQL + "   FROM TIM INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD "
        SQL = SQL + " WHERE (ENTITYMARK IS NULL OR ENTITYMARK='ERROR' ) AND    LEFT(ATIM,1) IN     (  " + PAR + "  )    and HME>='" + Format(APO.Value, "MM/dd/yyyy") + "'  AND HME<='" + Format(EOS.Value, "MM/dd/yyyy") + "'  "
        SQL = SQL + "  AND AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7>0  " + SYNT
        SQL = SQL + " order by HME"




        '  SQL = "SELECT  top 20  AJ1 ,AJ2  from TIM  order by HME"

        ExecuteSQLQuery(SQL)

        If sqlDT.Rows.Count = 0 Then
            MsgBox("ΔΕΝ ΒΡΕΘΗΚΑΝ ΕΓΓΡΑΦΕΣ")
            Exit Sub
        End If



        Dim ff As String = "c:\txtfiles\inv.xml"  'c:\mercvb\m" + Format(Now, "yyyyddmmHHMM") + ".export" ' "\\Logisthrio\333\pr.export" '
        Dim writer As New XmlTextWriter(ff, System.Text.Encoding.UTF8)
        writer.WriteStartDocument(True)
        writer.Formatting = Formatting.Indented
        writer.Indentation = 2
        writer.WriteStartElement("InvoicesDoc")
        writer.WriteAttributeString("xmlns", "http://www.aade.gr/myDATA/invoice/v1.0")
        writer.WriteAttributeString("xsi:schemaLocation", "http://www.aade.gr/myDATA/invoice/v1.0 schema.xsd")
        writer.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")



        Dim i As Integer
        Dim sumNet As Single
        Dim sumFpa As Single

        For i = 0 To sqlDT.Rows.Count - 1
            Dim EGGTIM As New DataTable

            ExecuteSQLQuery("SELECT KODE,POSO,TIMM,EKPT,FPA FROM EGGTIM WHERE POSO>0 AND ID_NUM=" + sqlDT(i)("ID_NUM").ToString, EGGTIM)
            Dim DUM As New DataTable
            ExecuteSQLQuery("UPDATE TIM SET ENTITY=" + Str(i + 1) + " WHERE ID_NUM=" + sqlDT(i)("ID_NUM").ToString, DUM)




            sumNet = sqlDT(i)("aj1") + sqlDT(i)("aj2") + sqlDT(i)("aj3") + sqlDT(i)("aj4") + sqlDT(i)("aj5") + sqlDT(i)("aj6") + sqlDT(i)("aj7")
            sumFpa = sqlDT(i)("fpa1") + sqlDT(i)("fpa2") + sqlDT(i)("fpa3") + sqlDT(i)("fpa4") + sqlDT(i)("fpa6") + sqlDT(i)("fpa7")


            writer.WriteComment(sqlDT(i)("ATIM") + " " + Format(sqlDT(i)("HME"), "dd/MM/yyyy"))
            writer.WriteStartElement("invoice")
            crNode("uid", "", writer)
            crNode("mark", "", writer)


            '---------------------------------------------εκδοτης
            writer.WriteStartElement("issuer")
            crNode("vatNumber", "028783755", writer)
            crNode("country", "GR", writer)
            crNode("branch", "0", writer)
            writer.WriteStartElement("address")
            crNode("postalCode", """66100""", writer)
            crNode("city", """ΔΡΑΜΑ""", writer)
            writer.WriteEndElement() '/address
            writer.WriteEndElement() '/issuer

            '--------------------------------------------- πελατης
            writer.WriteStartElement("counterpart")
            crNode("vatNumber", sqlDT(i)("AFM"), writer)  ' crNode("vatNumber", "026677115", writer)
            crNode("country", "GR", writer)
            crNode("branch", "0", writer)
            writer.WriteStartElement("address")
            crNode("postalCode", sqlDT(i)("AFM"), writer)  'crNode("postalCode", """66100""", writer)
            crNode("city", sqlDT(i)("POL"), writer)  ' crNode("city", """ΔΡΑΜΑ""", writer)
            writer.WriteEndElement() ' /address
            writer.WriteEndElement() ' /counterpart



            '----------------------------------------------- header
            writer.WriteStartElement("invoiceHeader")
            crNode("series", "0", writer)
            crNode("aa", Mid(sqlDT(i)("ATIM"), 2, 6), writer)   '  crNode("aa", "15", writer)
            crNode("issueDate", Format(sqlDT(i)("hme"), "yyyy-MM-dd"), writer) ' crNode("issueDate", "2019-12-15", writer)
            crNode("invoiceType", "1.1", writer)   ' ειδος παραστατικού
            crNode("currency", "EUR", writer)
            crNode("exchangeRate", "1.0", writer)
            writer.WriteEndElement() ' /invoiceHeader



            Dim SYN_KAU, SYN_FPA As Double
            SYN_KAU = 0
            SYN_FPA = 0
            For L As Integer = 0 To EGGTIM.Rows.Count - 1

                Dim AJ As Single
                If IsDBNull(EGGTIM(L)("TIMM")) Then
                    AJ = 0
                Else
                    AJ = Math.Round(EGGTIM(L)("POSO") * EGGTIM(L)("TIMM") * (1 - EGGTIM(L)("EKPT") / 100), 2)
                End If

                Dim VAT As String
                '1 ΦΠΑ συντελεστής 24% 24% 
                '2 ΦΠΑ συντελεστής 13% 13% 
                '3 ΦΠΑ συντελεστής 6% 6% 
                '4 ΦΠΑ συντελεστής 17% 17% 
                '5 ΦΠΑ συντελεστής 9% 9% 
                '6 ΦΠΑ συντελεστής 4% 4% 
                '7 Άνευ Φ.Π.Α. 0%  
                '8 Εγγραφές χωρίς ΦΠΑ  (πχ Μισθοδοσία, Αποσβέσεις) 



                If EGGTIM(L)("FPA") = 1 Then '13%
                    VAT = "2"
                    SYN_KAU = SYN_KAU + AJ
                    SYN_FPA = SYN_FPA + AJ * 0.13
                ElseIf EGGTIM(L)("FPA") = 2 Then
                    VAT = "1"
                ElseIf EGGTIM(L)("FPA") = 2 Then
                    VAT = "1"
                    SYN_KAU = SYN_KAU + AJ
                    SYN_FPA = SYN_FPA + AJ * 0.24

                ElseIf EGGTIM(L)("FPA") = 5 Then
                    VAT = "7"
                    SYN_KAU = SYN_KAU + AJ

                ElseIf EGGTIM(L)("FPA") = 6 Then
                    VAT = "1"
                    SYN_KAU = SYN_KAU + AJ
                    SYN_FPA = SYN_FPA + AJ * 0.24
                ElseIf EGGTIM(L)("FPA") = 4 Then
                    VAT = "4"
                Else ' If EGGTIM(L)("FPA") = 2 Then
                    VAT = "1"
                End If
                '-----------------------------------------------  invoiceDetails
                writer.WriteStartElement("invoiceDetails")
                crNode("lineNumber", Str(L + 1), writer) '  crNode("lineNumber", "1", writer)
                crNode("quantity", EGGTIM(L)("POSO").ToString, writer)
                crNode("measurementUnit", "1", writer)
                crNode("netValue", Format(AJ, "######0.##"), writer)  ' crNode("netValue", "100", writer)
                crNode("vatCategory", VAT, writer) '1=24%   2=13%   
                writer.WriteEndElement()   ' /invoiceDetails
            Next

            ExecuteSQLQuery("UPDATE TIM SET AADEKAU=" + Format(SYN_KAU, "#######.#####") + ",AADEFPA=" + Format(SYN_FPA, "#######.#####") +
                            " WHERE ID_NUM=" + sqlDT(i)("ID_NUM").ToString, DUM)
            '------------------------------------------------ InvoiceSummary 
            writer.WriteStartElement("invoiceSummary")
            crNode("totalNetValue", Format(SYN_KAU, "#######.#####"), writer)  ' crNode("totalNetValue", "100", writer)
            crNode("totalVatAmount", Format(SYN_FPA, "#######.#####"), writer)  '  crNode("totalVatAmount", "24", writer)
            crNode("totalWithheldAmount", "0", writer)
            crNode("totalFeesAmount", "0", writer)
            crNode("totalStampDutyAmount", "0", writer)
            crNode("totalOtherTaxesAmount", "0", writer)
            crNode("totalDeductionsAmount", "0", writer)
            crNode("totalGrossValue", Format(SYN_KAU + SYN_FPA, "#######.#####"), writer)
            writer.WriteEndElement() '  /invoicesummary
            '=========================================================
            writer.WriteEndElement() ' / Invoice

        Next





        writer.WriteEndElement() 'InvoicesDoc








        writer.WriteEndDocument()
        writer.Close()
        MsgBox("ok")


        ListBox2.Items.Clear()
        FileOpen(1, "C:\TXTFILES\CHECKXSD.TXT", OpenMode.Output)



        ' Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim myDocument As New XmlDocument
        myDocument.Load(ff) ' m_filename)  ' "C:\somefile.xml"
        myDocument.Schemas.Add("http://www.aade.gr/myDATA/invoice/v1.0", "c:\txtfiles\InvoicesDoc-v0.5.1.xsd") 'namespace here or empty string
        Dim eventHandler As ValidationEventHandler = New ValidationEventHandler(AddressOf ValidationEventHandler)
        myDocument.Validate(eventHandler)
        MsgBox("ok ελεγχος")

        For n As Integer = 0 To ListBox2.Items.Count - 1
            PrintLine(1, ListBox2.Items(n).ToString)
        Next




        FileClose(1)

        paint_ergasies(DataGridView1, "SELECT ATIM,HME,ENTITY,AADEKAU,AJ1+AJ2+AJ3+AJ4+AJ5+AJ6+AJ7 AS KAUTIM,AADEFPA,FPA1+FPA2+FPA3+FPA4+FPA6+FPA7 AS FPATIM,ENTITYUID,ENTITYMARK FROM TIM WHERE ENTITY>0")

    End Sub



    Private Sub Test()
        Dim objWorkingXML As New System.Xml.XmlDocument
        Dim objValidateXML As System.Xml.XmlValidatingReader
        Dim objSchemasColl As New System.Xml.Schema.XmlSchemaCollection

        objSchemasColl.Add("http://www.aade.gr/myDATA/invoice/v1.0", "c:\txtfiles\InvoicesDoc-v0.5.1.xsd")
        objValidateXML = New System.Xml.XmlValidatingReader(New System.Xml.XmlTextReader("c:\txtfiles\inv2.xml"))

        AddHandler objValidateXML.ValidationEventHandler, AddressOf ValidationCallBack
        objValidateXML.Schemas.Add(objSchemasColl)

        'This is WHERE the validation occurs.. WHEN the XML Document READS through the validating reader

        objWorkingXML.Load(objValidateXML)

        'Close the stream
        objValidateXML.Close()


        'The document is valid
        MsgBox("THE DOCUMENT IS VALID")

    End Sub



    Private Shared Sub ValidationCallBack(ByVal sender As Object, ByVal e As System.Xml.Schema.ValidationEventArgs)
        Dim test As String
        MsgBox(e.Message)
        ' Throw e.Exception

    End Sub







    Private Sub ValidationEventHandler(ByVal sender As Object, ByVal e As ValidationEventArgs)
        Select Case e.Severity
            Case XmlSeverityType.Error
                Debug.WriteLine("Error: {0}", e.Message)
                ListBox2.Items.Add("ERROR " + e.Message)
            Case XmlSeverityType.Warning
                Debug.WriteLine("Warning {0}", e.Message)
                ListBox2.Items.Add("warning " + e.Message)
        End Select
    End Sub




    Private Sub crNode(ByVal pName As String, ByVal cValue As String, ByVal writer As XmlTextWriter)
        writer.WriteStartElement(pName)
        writer.WriteString(cValue)
        writer.WriteEndElement()
    End Sub



    Private Sub EditConnString_Click(sender As Object, e As EventArgs) Handles EditConnString.Click
        If checkServer(1) Then
            MsgBox("OK")
            ExecuteSQLQuery("IF COL_LENGTH('dbo.TIM', 'AADEFPA') IS  NULL  BEGIN; ALTER TABLE TIM ADD AADEFPA FLOAT NULL;END")
            ExecuteSQLQuery("IF COL_LENGTH('dbo.TIM', 'AADEKAU') IS  NULL  BEGIN; ALTER TABLE TIM ADD AADEKAU FLOAT NULL;END")
            ExecuteSQLQuery("IF COL_LENGTH('dbo.TIM', 'ENTITY') IS  NULL  BEGIN;  ALTER TABLE TIM ADD ENTITY  INT NULL;END")
            ExecuteSQLQuery("IF COL_LENGTH('dbo.TIM', 'ENTITYUID') IS  NULL  BEGIN;  ALTER TABLE TIM ADD ENTITYUID VARCHAR(40) NULL;END")
            ExecuteSQLQuery("IF COL_LENGTH('dbo.TIM', 'ENTITYMARK') IS  NULL  BEGIN; ALTER TABLE TIM ADD ENTITYMARK VARCHAR(13) NULL;END")

        End If
    End Sub

    'Δουλεύει κανονικά, αν έχεις κάποια απορία στείλε μου email στο manol14@hotmail.com









    '  <totalNetValue>100</totalNetValue>
    '    <totalVatAmount>24</totalVatAmount>
    '    <totalWithheldAmount>0</totalWithheldAmount>
    '    <totalFeesAmount>0</totalFeesAmount>
    '    <totalStampDutyAmount>0</totalStampDutyAmount>
    '    <totalOtherTaxesAmount>0</totalOtherTaxesAmount>
    '    <totalDeductionsAmount>0</totalDeductionsAmount>
    '    <totalGrossValue>124</totalGrossValue>
    '  </invoiceSummary>
    '</invoice>


    Public Function ExecuteSQLQuery(ByVal SQLQuery As String) As DataTable
        Try
            Dim sqlCon As New OleDbConnection(gConnect)

            Dim sqlDA As New OleDbDataAdapter(SQLQuery, sqlCon)

            Dim sqlCB As New OleDbCommandBuilder(sqlDA)
            sqlDT.Reset() ' refresh 
            sqlDA.Fill(sqlDT)
            'rowsAffected = command.ExecuteNonQuery();
            ' sqlDA.Fill(sqlDaTaSet, "PEL")

        Catch ex As Exception
            MsgBox("Error: " & ex.ToString)
            If Err.Number = 5 Then
                MsgBox("Invalid Database, Configure TCP/IP", MsgBoxStyle.Exclamation, "Sales and Inventory")
            Else
                MsgBox("Error : " & ex.Message)
            End If
            MsgBox("Error No. " & Err.Number & " Invalid database or no database found !! Adjust settings first", MsgBoxStyle.Critical, "Sales And Inventory")
            MsgBox(SQLQuery)
        End Try
        Return sqlDT
    End Function

    Public Sub ExecuteSQLQuery(ByVal SQLQuery As String, ByRef SQLDT As DataTable)
        'αν χρησιμοποιώ  byref  tote prepei να δηλωθεί   
        'Dim DTI As New DataTable


        Try
            Dim sqlCon As New OleDbConnection(gConnect)

            Dim sqlDA As New OleDbDataAdapter(SQLQuery, sqlCon)

            Dim sqlCB As New OleDbCommandBuilder(sqlDA)
            'SQLDT.Reset() ' refresh 
            sqlDA.Fill(SQLDT)
            ' sqlDA.Fill(sqlDaTaSet, "PEL")

        Catch ex As Exception
            MsgBox("Error: " & ex.ToString)
            If Err.Number = 5 Then
                MsgBox("Invalid Database, Configure TCP/IP", MsgBoxStyle.Exclamation, "Sales and Inventory")
            Else
                MsgBox("Error : " & ex.Message)
            End If
            MsgBox("Error No. " & Err.Number & " Invalid database or no database found !! Adjust settings first", MsgBoxStyle.Critical, "Sales And Inventory")
            MsgBox(SQLQuery)
        End Try
        'Return sqlDT
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Test()
    End Sub

    Private Sub UPDATE_TIM_Click(sender As Object, e As EventArgs) Handles UPDATE_TIM.Click
        'ΠΑΙΡΝΩ ΤΟ  "c:\txtfiles\apantSendInv.XML")
        ' ΚΑΙ ΕΝΗΜΕΡΩΝΩ ΤΟ ΤΙΜ ΚΑΙ ΜΕΤΑ
        '' ΕΔΩ ΔΗΜΙΟΥΡΓΩ ΤΟ INCOME ΑΡΧΕΙΟ ΜΕ ΠΡΟΣΔΙΟΡΙΣΜΟ ΤΗΣ ΚΑΘΕ ΕΓΓΡΑΦΗΣ    "c:\txtfiles\inC.xml"





        If checkServer(0) Then
            ' MsgBox("OK")
        End If



        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load("C:\TXTFILES\apantSendInv.xml")
        Dim nodes As XmlNodeList = xmlDoc.DocumentElement.SelectNodes("/ResponseDoc/response")

        'ExecuteSQLQuery("delete from oc_category_filter")
        Me.Text = "Ελεγχος αριθμου εγγραφών"
        Application.DoEvents()

        Dim NF As Long = 0
        For Each node As XmlNode In nodes
            NF = NF + 1
        Next
        Application.DoEvents()
        Me.Text = NF


        ' ΕΔΩ ΔΗΜΙΟΥΡΓΩ ΤΟ INCOME ΑΡΧΕΙΟ ΜΕ ΠΡΟΣΔΙΟΡΙΣΜΟ ΤΗΣ ΚΑΘΕ ΕΓΓΡΑΦΗΣ

        Dim ff As String = "c:\txtfiles\inC.xml"  'c:\mercvb\m" + Format(Now, "yyyyddmmHHMM") + ".export" ' "\\Logisthrio\333\pr.export" '
        Dim writer As New XmlTextWriter(ff, System.Text.Encoding.UTF8)
        writer.WriteStartDocument(True)
        writer.Formatting = Formatting.Indented
        writer.Indentation = 2
        writer.WriteStartElement("IncomeClassificationsDoc")
        writer.WriteAttributeString("xmlns", "https://www.aade.gr/myDATA/incomeClassificaton/v1.0")
        writer.WriteAttributeString("xsi:schemaLocation", "https://www.aade.gr/myDATA/incomeClassificaton/v1.0 schema.xsd")
        writer.WriteAttributeString("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")








        Dim line As Integer
        Dim entityUid As String
        Dim entityMark As String
        For Each node As XmlNode In nodes
            Dim Status As String = node.SelectSingleNode("statusCode").InnerText
            Dim merror As String
            line = node.SelectSingleNode("entitylineNumber").InnerText


            If Status = "Success" Then
                entityUid = node.SelectSingleNode("entityUid").InnerText
                entityMark = node.SelectSingleNode("entityMark").InnerText

                'ΑΝ ΕΧΕΙ ΑΠΟΣΤΑΛΕΙ ΤΟ ΤΙΜΟΛΟΓΙΟ ΜΕ ΕΠΙΤΥΧΙΑ ΠΑΙΡΝΩ ΤΗΝ ΕΥΚΑΙΡΙΑ
                'ΝΑ ΣΤΕΙΛΩ ΚΑΙ ΤΟΝ ΤΥΠΟ ΤΟΥ ΕΣΟΔΟΥ
                Dim temp As New DataTable
                ExecuteSQLQuery("select AADEKAU,ID_NUM,ATIM,HME FROM TIM   WHERE ENTITY=" + Str(line), temp)

                Dim EGGTIM As New DataTable
                ExecuteSQLQuery("select POSO*TIMM*(100-EKPT)/100 AS AJ FROM EGGTIM   WHERE POSO*TIMM<>0 AND ID_NUM=" + Str(temp(0)(1)), EGGTIM)

                writer.WriteComment(temp(0)("ATIM") + " " + Format(temp(0)("HME"), "dd/MM/yyyy"))

                writer.WriteStartElement("incomeInvoiceClassification") '---------------------------
                crNode("mark", entityMark, writer)

                For L As Integer = 0 To EGGTIM.Rows.Count - 1

                    writer.WriteStartElement("invoicesIncomeClassificationDetails") '====
                    crNode("lineNumber", Str(L + 1), writer)
                    writer.WriteStartElement("incomeClassificationDetailData") '***
                    crNode("classificationType", "101", writer)
                    crNode("classificationCategory", "1", writer)
                    crNode("amount", Format(EGGTIM(L)(0), "#####0.#####"), writer)
                    writer.WriteEndElement() 'incomeClassificationDetailData    '****
                    writer.WriteEndElement() ' invoicesIncomeClassificationDetails   ======

                Next
                writer.WriteEndElement() ' incomeInvoiceClassification---------------------

            Else 'ΕΧΕΙ ΛΑΘΟΣ ΟΠΟΤΕ ΑΠΟΘΗΚΕΥΩ ΤΟ ΛΑΘΟΣ ΣΤΟ ΤΙΜ.entityUid
                entityUid = node.SelectSingleNode("errors/error/message").InnerText
                entityMark = "ERROR"
            End If

            ExecuteSQLQuery("update TIM SET ENTITYUID='" + Mid(entityUid, 1, 40) + "' , ENTITYMARK='" + Mid(entityMark, 1, 13) + "' WHERE ENTITY=" + Str(line))


        Next

        writer.WriteEndDocument()
        writer.Close()

    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged

    End Sub

    Public Sub paint_ergasies(GridView1 As DataGridView, sqlQuery As String)



        Dim connectionString As String = gSQLCon '"Data Source=.;Initial Catalog=pubs;Integrated Security=True"
        Dim sql As String = sqlQuery ' "SELECT * FROM TIM"
        Dim connection As New SqlConnection(connectionString)
        Dim dataadapter As New SqlDataAdapter(sql, connection)
        Dim ds As New DataSet()
        connection.Open()
        dataadapter.Fill(ds, "Authors_table")
        connection.Close()
        GridView1.DataSource = ds
        GridView1.DataMember = "Authors_table"










        'Dim con As New OleDb.OleDbConnection
        'con.ConnectionString = gConnect
        'con.Open()
        'Dim ds As DataSet = New DataSet
        'Dim adapter As New OleDb.OleDbDataAdapter
        'Dim sql As String

        'sql = sqlQuery

        'adapter.SelectCommand = New OleDb.OleDbCommand(sql, con)
        'adapter.Fill(ds)
        'GridView1.DataSource = ds.Tables(0)


































        Exit Sub

        ''Create connection
        'Dim conn As SqlConnection

        ''create data adapter
        'Dim da As New SqlDataAdapter

        ''create dataset
        'Dim ds As DataSet = New DataSet

        'conn = New SqlConnection(gSQLCon)
        'Try
        '    ' Open connection
        '    conn.Open()

        '    da = New SqlDataAdapter(sqlQuery, conn)

        '    'create command builder
        '    Dim cb As SqlCommandBuilder = New SqlCommandBuilder(da)
        '    ds.Clear()
        '    'fill dataset
        '    da.Fill(ds, "PEL")
        '    GridView1.ClearSelection()
        '    GridView1.DataSource = ds
        '    GridView1.DataMember = "PEL"


        '    GridView1.Refresh()







        'Catch ex As SqlException
        '    MsgBox(ex.ToString)
        'Finally
        '    ' Close connection
        '    conn.Close()
        'End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If checkServer(0) Then

        End If
        paint_ergasies(DataGridView1, "Select ATIM, HME, ENTITY, AADEKAU, AJ1 + AJ2 + AJ3 + AJ4 + AJ5 + AJ6 + AJ7 As KAUTIM, AADEFPA, FPA1 + FPA2 + FPA3 + FPA4 + FPA6 + FPA7 As FPATIM, ENTITYUID, ENTITYMARK FROM TIM WHERE ENTITY>0")
    End Sub

    ' Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
    Private Async Sub MakeRequest2()
        Dim client = New HttpClient()
        Dim queryString = HttpUtility.ParseQueryString(String.Empty)
        client.DefaultRequestHeaders.Add("aade-user-id", "glagakis2")
        client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", "555bc57c80634243958f62b629316aaa")
        queryString("mark") = "1000000006337" ' "{string}"
        '  queryString("nextPartitionKey") = "{string}"
        '   queryString("nextRowKey") = "{string}"
        Dim uri = "https://mydata-dev.azure-api.net/RequestIssuerInvoices?" & queryString.ToString

        Dim response = Await client.GetAsync(uri)
        Dim result = Await response.Content.ReadAsStringAsync()
        TextBox2.Text = result.ToString

        Dim MF = "c:\txtfiles\apantReqInv" + Format(Now, "yyyyddMMHHmm") + ".xml"
        FileOpen(1, MF, OpenMode.Output)
        PrintLine(1, result.ToString)
        FileClose(1)






    End Sub

    '        Using System;
    'Using System.Net.Http.Headers;
    'Using System.Text;
    'Using System.Net.Http;
    'Using System.Web;

    'Namespace CSHttpClientSample
    '{
    '    Static Class Program
    '    {
    '        Static void Main()
    '        {
    '            MakeRequest();
    '            Console.WriteLine("Hit ENTER to exit...");
    '            Console.ReadLine();
    '        }

    '        Static Async void MakeRequest()
    '        {
    '            var client = New HttpClient();
    '            var queryString = HttpUtility.ParseQueryString(String.Empty);

    '            // Request headers
    '            client.DefaultRequestHeaders.Add("aade-user-id", "");
    '            client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", "{subscription key}");

    '            // Request parameters
    '            queryString["mark"] = "{string}";
    '            queryString["nextPartitionKey"] = "{string}";
    '            queryString["nextRowKey"] = "{string}";
    '            var uri = "https://mydata-dev.azure-api.net/RequestIssuerInvoices?" + queryString;

    '            var response = await client.GetAsync(uri);
    '        }
    '    }
    '}	
    '    End Sub


    '  Imports System
    '  Imports System.Net.Http.Headers
    '  Imports System.Text
    '  Imports System.Net.Http
    '  Imports System.Web'
    'Namespace CSHttpClientSample
    '   Module' Program
    Private Sub Main2()
        MakeRequest()
        Console.WriteLine("Hit ENTER to exit...")
        Console.ReadLine()
    End Sub

    'Private Async Sub MakeRequest3()
    '    Dim client = New HttpClient()
    '    Dim queryString = HttpUtility.ParseQueryString(String.Empty)
    '    client.DefaultRequestHeaders.Add("aade-user-id", "")
    '    client.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", "{subscription key}")
    '    queryString("mark") = "{string}"
    '    Dim uri = "https://mydata-dev.azure-api.net/RequestInvoices?" & queryString
    '    Dim response = Await client.GetAsync(uri)
    'End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        MakeRequest2()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        '======================================================================================
        'διαβαζω το apantIncome.xml για να δω τελικά ποια έχουν πρόβλημα 
        ' και να αποθηκεύσω το αποτέλεσμα στο ΤΙΜ
        Dim cc(1000) As String


        Dim xmlDoc As New XmlDocument()
        xmlDoc.Load("C:\TXTFILES\apantiNCOMe.xml")
        Dim nodes As XmlNodeList = xmlDoc.DocumentElement.SelectNodes("/ResponseDoc/response")

        'ExecuteSQLQuery("delete from oc_category_filter")
        Me.Text = "Ελεγχος αριθμου εγγραφών"
        Application.DoEvents()



        Dim k As Integer = 0
        For Each node As XmlNode In nodes
            k = k + 1

            Dim Status As String = node.SelectSingleNode("statusCode").InnerText
            cc(k) = Status
            Dim line As String = node.SelectSingleNode("entitylineNumber").InnerText
        Next




        ' ΔΕΝ ΔΙΑΒΑΖΩ ΤΟ INC.XML ΓΙΑΤΙ ΤΑ ATTRIBUTES ΕΜΠΟΔΙΖΟΥΝ ΤΟ ΔΙΑΒΑΣΜΑ ΤΩΝ NODES
        ' ΕΝΩ ΑΝ ΣΒΗΣΩ ΤΑ ATTRIBUTES ΔΙΠΛΑ ΑΠΟ ΤΟ IncomeClassificationsDoc ΔΟΥΛΕΥΕΙ ΚΑΝΟΝΙΚΑ




    End Sub
    '   End Module
    'End Namespace
End Class
