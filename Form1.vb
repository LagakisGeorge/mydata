Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.IO
Imports System.Xml
Imports System.Data.OleDb
Imports Microsoft.VisualBasic.Compatibility.VB6


Public Class Form1


    Public openedFileStream As System.IO.Stream
    Public dataBytes() As Byte
    Public gConnect As String
    Public sqlDT As New DataTable
    Public sqlDT2 As New DataTable


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MakeRequest()

    End Sub


    Private Async Sub MakeRequest()
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

                TextBox2.Text = result.ToString
                ' "είναι το textbox πανω στη φόρμα που σου επιστρέφει το response xml"

            End Using


        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try


    End Sub


    Private Async Sub MakeIncomeRequest()
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




            Else
                'local
                'MsgBox(Split(tmpStr, ":")(1))
                gConnect = "Provider=SQLOLEDB;Server=" & Split(tmpStr, ":")(1) &
                           ";Database=" & Split(tmpStr, ":")(5) & "; Trusted_Connection=yes;"

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

        ExecuteSQLQuery("select POL,EIDOS,AJIA_APOU from PARASTAT", SQLDT2)

        pol = " "

        Dim row As Integer
        For row = 0 To SQLDT2.Rows.Count - 1

            If IsDBNull(SQLDT2.Rows(row)("eidos")) Or IsDBNull(SQLDT2.Rows(row)("pol")) Or IsDBNull(SQLDT2.Rows(row)("ajia_apou")) Then

            Else

                If SQLDT2.Rows(row)("pol") = "1" And SQLDT2.Rows(row)("ajia_apou") = "3" Then
                    pol = pol + "'" + SQLDT2.Rows(row)("eidos") + "',"
                End If

                If SQLDT2.Rows(row)("pol") = "1" And SQLDT2.Rows(row)("ajia_apou") = "4" Then
                    polepis = polepis + "'" + SQLDT2.Rows(row)("eidos") + "',"
                End If

                If SQLDT2.Rows(row)("pol") = "2" And SQLDT2.Rows(row)("ajia_apou") = "1" Then
                    ago = ago + "'" + SQLDT2.Rows(row)("eidos") + "',"
                End If

                If SQLDT2.Rows(row)("pol") = "2" And SQLDT2.Rows(row)("ajia_apou") = "2" Then
                    AGOEPIS = AGOEPIS + "'" + SQLDT2.Rows(row)("eidos") + "',"
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
        '====================================================================================
        Dim pol, polepis, ago, AGOEPIS As String


        If checkServer(0) Then
            ' MsgBox("OK")
        End If


        Get_AJ_ASCII(pol, polepis, ago, AGOEPIS)

        Dim PAR, SYNT As String
        PAR = pol + polepis
        Dim SQL As String
        SYNT = ""
        SQL = "SELECT ID_NUM, AJ1  ,AJ2 , AJ3,AJ4,AJ5,AJI,FPA1,FPA2,FPA3,FPA4,ATIM,"
        SQL = SQL + "HME,PEL.EPO,PEL.AFM,KPE,PEL.DIE,PEL.XRVMA"    '"CONVERT(CHAR(10),HME,3) AS HMEP
        SQL = SQL + ",PEL.EPA,PEL.POL,AJ6,FPA6,AJ7,FPA7 "

        SQL = SQL + "   FROM TIM INNER JOIN PEL ON TIM.EIDOS=PEL.EIDOS AND TIM.KPE=PEL.KOD "
        SQL = SQL + " WHERE LEFT(ATIM,1) IN     (  " + PAR + "  )    and HME>='" + Format(APO.Value, "MM/dd/yyyy") + "'  AND HME<='" + Format(EOS.Value, "MM/dd/yyyy") + "'  "
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
            sumNet = sqlDT(i)("aj1") + sqlDT(i)("aj2") + sqlDT(i)("aj3") + sqlDT(i)("aj4") + sqlDT(i)("aj5") + sqlDT(i)("aj6") + sqlDT(i)("aj7")
            sumFpa = sqlDT(i)("fpa1") + sqlDT(i)("fpa2") + sqlDT(i)("fpa3") + sqlDT(i)("fpa4") + sqlDT(i)("fpa6") + sqlDT(i)("fpa7")

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
            crNode("aa", sqlDT(i)("ATIM"), writer)   '  crNode("aa", "15", writer)
            crNode("issueDate", Format(sqlDT(i)("ATIM"), "yyyy-MM-dd"), writer) ' crNode("issueDate", "2019-12-15", writer)
            crNode("invoiceType", "1.1", writer)
            crNode("currency", "EUR", writer)
            crNode("exchangeRate", "1.0", writer)
            writer.WriteEndElement() ' /invoiceHeader


            '-----------------------------------------------  invoiceDetails
            writer.WriteStartElement("invoiceDetails")
            crNode("lineNumber", "1", writer)
            crNode("quantity", "1", writer)
            crNode("measurementUnit", "1", writer)
            crNode("netValue", Format(sumNet, "#######.##"), writer)  ' crNode("netValue", "100", writer)
            crNode("vatCategory", "1", writer)
            writer.WriteEndElement()   ' /invoiceDetails


            '------------------------------------------------ InvoiceSummary 
            writer.WriteStartElement("invoiceSummary")
            crNode("totalNetValue", Format(sumNet, "#######.##"), writer)  ' crNode("totalNetValue", "100", writer)
            crNode("totalVatAmount", "24", writer)
            crNode("totalWithheldAmount", "0", writer)
            crNode("totalFeesAmount", "0", writer)
            crNode("totalStampDutyAmount", "0", writer)
            crNode("totalOtherTaxesAmount", "0", writer)
            crNode("totalDeductionsAmount", "0", writer)
            crNode("totalGrossValue", "124", writer)
            writer.WriteEndElement() '  /invoicesummary
            '=========================================================
            writer.WriteEndElement() ' / Invoice

        Next





        writer.WriteEndElement() 'InvoicesDoc








        writer.WriteEndDocument()
        writer.Close()




    End Sub


    Private Sub crNode(ByVal pName As String, ByVal cValue As String, ByVal writer As XmlTextWriter)
        writer.WriteStartElement(pName)
        writer.WriteString(cValue)
        writer.WriteEndElement()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        MakeIncomeRequest()
    End Sub

    Private Sub EditConnString_Click(sender As Object, e As EventArgs) Handles EditConnString.Click
        If checkServer(1) Then
            MsgBox("OK")
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




End Class
