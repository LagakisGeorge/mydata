Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Xml
Imports System.IO
Imports System.Data.OleDb
Imports System.Xml.Schema
Imports System.Data.SqlClient
Imports System.Web
Imports System.CodeDom.Compiler
Imports HtmlAgilityPack
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq






Public Class nikosg
    Dim PEDIO(150) As String, TIMH(150) As String
    Dim f_f As String
    Dim reader As JsonTextReader
    Public openedFileStream As System.IO.Stream
    Public dataBytes() As Byte
    Public gConnect As String
    Public gSQLCon As String
    Public sqlDT As New DataTable
    Public sqlDT2 As New DataTable
    Public gUserId As String
    Public gSubKey As String
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        ListBox2.Items.Clear()
        SKROUTZ_MakeRequest()

    End Sub

    Private Async Sub SKROUTZ_MakeRequest()

        '        CREATE TABLE [dbo].[orderdetails](
        '	[order.code] [nchar](15) NULL,
        '	[invoice] [nchar](10) NULL,
        '    [courier] [nchar](10) NULL,
        '    [shop_uid] [nchar](10) NULL,
        '    [product_name] [nvarchar](50) NULL,
        '    [quantity] [real] NULL,
        '    [unit_price] [real] NULL,
        '    [total_price] [real] NULL,
        '    [price_includes_vat] [nchar](10) NULL,
        '    [DateTime] [DateTime] NULL,
        '    [id] [Int] IDENTITY(1, 1) Not NULL,
        '	[idorder] [nvarchar](16) NULL
        '     [Line] int
        ') ON [PRIMARY]

        If InStr(order.Text, "-") > 0 Then
        Else
            MsgBox("πληκτρολογηστε παραγγελία")
            Exit Sub
        End If

        Dim client = New HttpClient()

        Try
            client.DefaultRequestHeaders.Add("Accept", "application/vnd.skroutz+json; version=3.0")

            '// client.DefaultRequestHeaders.Add("Authorization", "xobten1524571846")
            client.DefaultRequestHeaders.Add("Authorization", "Bearer BQN-WsvJFdttSjg00rpR1bA6fKquC72hFi3l6DzDYuDFQJb9ksG_yLtcjqeAcgnPNLVJjnWDzdAbeVn1ovB-vQ==")
            ' // client.DefaultRequestHeaders.Add("scope", "public")

            Dim uri = "https://api.skroutz.gr/merchants/ecommerce/orders/" + order.Text ' 210127-9011968"  'hwww.skroutz.gr/oauth2/token\?"
            Dim response As HttpResponseMessage
            '      Dim xl As String = XDocument.Load("c:\txtfiles\inv.xml").ToString ' "--> εκει έχω αποθηκεύσει το xml που εφτιαξα"
            '     Dim byteData As Byte() = Encoding.UTF8.GetBytes(xl)

            ' Dim response As HttpResponseMessage
            'response = Await client.PostAsync(uri, "")


            response = Await client.GetAsync(uri)
            Dim result = Await response.Content.ReadAsStringAsync()
            ' TextBox2.Text = result.ToString

            Dim MF = "c:\txtfiles\sendinv\skr" + Format(Now, "yyyyddMMHHmm") + ".json"
            FileOpen(1, MF, OpenMode.Output)
            PrintLine(1, result.ToString)
            FileClose(1)



            '  {"order"{"code":"210127-5178005",
            '"state""accepted",
            '"customer":
            '  {"id""GQYoapve57","first_name":"Βαγγέλης","last_name":"Ελληνικακης",
            '  "address":{"street_name":"kalomiri","street_number":"22","zip":"83100","city":"samos","region":"Σάμου","pickup_from_collection_point":false}},
            '  "invoice":false,"comments":"","courier":"ACS","courier_voucher":null, "courier_tracking_codes":  [],
            '  "line_items":[
            '  {"id":"MYKOK4gNYn","shop_uid":"1002","product_name":"Motospeed H19 Gaming Headset Grey","quantity":1,"unit_price":20.0,"total_price":20.0,"price_includes_vat":true}],


            '  "created_at":"2021-01-27T13:09:01+02:00","expires_at":"2021-01-28T10:09:01+02:00",
            '  "dispatch_until":"2021-01-29T18:00:00+02:00"}}
            '            order.code = 210127 - 5178005
            '            order.state = accepted
            '            order.customer.id = GQYoapve57
            '            order.customer.first_name = Βαγγέλης
            '            order.customer.last_name = Ελληνικακης
            '            order.customer.address.street_name = kalomiri
            '            order.customer.address.street_number = 22
            '            order.customer.address.zip = 83100
            '            order.customer.address.city = samos
            '            order.customer.address.region = Σάμου
            '            order.customer.address.pickup_from_collection_point = False
            '            order.invoice = False
            '            order.courier = ACS
            '            order.line_items[0].id=MYKOK4gNYn
            'order.line_items[0].shop_uid=1002
            'order.line_items[0].product_name=Motospeed H19 Gaming Headset Grey
            'order.line_items[0].quantity=1
            'order.line_items[0].unit_price=20
            'order.line_items[0].total_price=20
            'order.line_items[0].price_includes_vat=True
            'order.created_at = 27 / 1 / 2021 1:09:01 μμ
            'order.expires_at = 28 / 1 / 2021 10:09:01 πμ
            'order.dispatch_until = 29 / 1 / 2021 6:00:00 μμ

            'Dim PEDIO(150) As String, TIMH(150) As String
            Dim f, l As String
            FileOpen(1, "c:\txtfiles\jasonstr.txt", OpenMode.Output)
            'PrintLine(1, result.ToString)
            'FileClose(1)
            reader = New JsonTextReader(New StringReader(result))
            Dim c As String
            Dim N As Integer = 0
            While reader.Read
                If reader.Value IsNot Nothing Then
                    If InStr(reader.Path, reader.Value) = 0 Then
                        N = N + 1
                        PEDIO(N) = "[" + reader.Path + "]"
                        TIMH(N) = reader.Value

                        ListBox2.Items.Add(reader.Path + "=" + reader.Value.ToString)
                        PrintLine(1, reader.Path + "=" + reader.Value.ToString)
                    End If

                End If
            End While

            FileClose(1)

            Dim K As Integer
            Dim dts As New DataTable
            Try
                ExecuteSQLQuery("INSERT INTO orders ([order.code]) values ('" + TIMH(1) + "')")
            Catch ex As Exception
                MsgBox("εχει περαστεί ήδη η παραγγελια")
                Return
            End Try


            For K = 2 To N
                If InStr(PEDIO(K), "customer") > 0 Or InStr(PEDIO(K), "invoice") > 0 Then
                    Try
                        ListBox4.Items.Add(PEDIO(K) + "  " + TIMH(K))
                        If InStr(PEDIO(K), "order.customer.address.street_number") > 0 Then
                            TIMH(K) = Mid(TIMH(K), 1, 5)
                        End If

                        ExecuteSQLQuery("update orders Set " + PEDIO(K) + "='" + TIMH(K) + "' where [order.code]='" + TIMH(1) + "'")

                    Catch ex As Exception

                    End Try
                End If

            Next

            '
            Dim product_name As String = Replace(FindByName("[order.line_items[0].product_name]"), "'", "`")
            Dim shop_uid As String = FindByName("[order.line_items[0].shop_uid]")
            Dim product_id As String = FindByName("[order.line_items[0].id]")
            Dim quantity As String = FindByName("[order.line_items[0].quantity]")
            Dim unit_price As String = FindByName("[order.line_items[0].unit_price]")
            '  Dim shop_uid As String = FindByName("[order.line_items[0].shop_uid]")
            unit_price = Replace(unit_price, ",", ".")

            Dim SQL As String
            'order.line_items[0].unit_price
            SQL = "insert into orderdetails (shop_uid,[order.code],line,product_name,product_id,quantity,unit_price,date) values ('" + shop_uid + "','" + TIMH(1) + "',1,'" + product_name + "','" + product_id + "'," + quantity + "," + unit_price + ",GETDATE() )"
            ExecuteSQLQuery("INSERT into LOGGING (LOGGING) VALUES ('" + Replace(SQL, "'", "`") + "')")

            ExecuteSQLQuery(SQL)


            'FindByName




            '  Dim result = Await response.Content.ReadAsStringAsync()
            '   response = Await client.GetAsync(uri)
            'Using content = New ByteArrayContent(byteData)
            'content.Headers.ContentType = New MediaTypeHeaderValue("application/xml")
            ' https://api.skroutz.gr/search?q=apple
            ' response = Await client.PostAsync(uri, content)
            '  Dim result = Await response.Content.ReadAsStringAsync()
            'response = Await client.PostAsync(uri, content)
            'Dim result = Await response.Content.ReadAsStringAsync()
            'Dim MF = "c:\txtfiles\sendinv\apantSendInv" + Format(Now, "yyyyddMMHHmm") + ".xml"
            'FileOpen(1, MF, OpenMode.Output)
            'PrintLine(1, result.ToString)
            'FileClose(1)


            ' End Using


        Catch ex As Exception

            MsgBox(ex.ToString)
        End Try


        ExecuteSQLQuery("select [order.code] as kod,shop_uid from orderdetails")
        ListBox3.Items.Clear()
        For k As Integer = 0 To sqlDT.Rows.Count - 1

            ListBox3.Items.Add(sqlDT.Rows(k)("kod") + "  κωδ¨:" + sqlDT.Rows(k)("shop_uid"))

        Next





    End Sub

    Function FindByName(c As String)
        FindByName = "0"
        For k As Integer = 1 To 150
            If PEDIO(k) = c Then
                FindByName = TIMH(k)
                Exit For

            End If
        Next
    End Function

    Private Sub nikosg_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If checkServer(0) Then

        End If

        ChromiumWebBrowser1.Load("https://merchants.skroutz.gr/merchants")

    End Sub


    Function checkServer(ByVal check_path As Integer) As Boolean
        Dim c As String
        Dim tmpStr As String
        c = "Config.ini"


        Dim par As String = ""
        Dim mf As String
        mf = c   ' "c:\mercvb\err3.txt"  If System.IO.File.Exists(SavePath) Then
        If Len(Dir(UCase(mf))) = 0 Then
            par = ":DELLAGAKIS\SQL17:sa:12345678:1:EMP"    '" 'G','g','Ξ','D'  "
            par = InputBox("ΒΑΣΗ ΔΕΔΟΜΕΝΩΝ", , par)
            gUserId = InputBox("Χρήστης", gUserId)
            gSubKey = InputBox("Κλειδί", gSubKey)
        Else
            FileOpen(11, mf, OpenMode.Input)
            '   Input(1, par)
            par = LineInput(11)
            gUserId = LineInput(11)
            gSubKey = LineInput(11)
            FileClose(11)
        End If
        If check_path = 1 Then
            par = InputBox("ΒΑΣΗ ΔΕΔΟΜΕΝΩΝ  (CONFIG.INI ΣΤΟΝ ΤΡΕΧΟΝΤΑ ΦΑΚΕΛΟ) ", ":Π.Χ. (local)\sql2012:sa:12345678:1:EMP", par)
            gUserId = InputBox("Χρήστης", gUserId)
            gSubKey = InputBox("Κλειδί", gSubKey)
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


        FileOpen(7, mf, OpenMode.Output)
        PrintLine(7, par)
        PrintLine(7, gUserId)
        PrintLine(7, gSubKey)

        FileClose(7)

        Dim A As String
        If System.IO.File.Exists("DATES.TXT") Then
            FileOpen(31, "DATES.TXT", OpenMode.Input)
            A = LineInput(31)
            'ListBox2.Items.Add(A)
            A = LineInput(31)
            ' ListBox2.Items.Add(A)
            FileClose(31)

        End If


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
            '  ERROR_WRITE(Format(Now, "dd/MM/yy") + "Error No. " & Err.Number & " Invalid database or no database found !! Adjust settings first")
            '  ERROR_WRITE(SQLQuery)
            'MsgBox(SQLQuery)
            ' End


        End Try
        Return sqlDT
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        navigating()
    End Sub


    Async Sub navigating()



        ' Dim mydriver As New OpenQA.Selenium.ChromeDriver("C:\Users\Marc\Downloads\chromedriver_win32")

        ChromiumWebBrowser1.Load("https://merchants.skroutz.gr/merchants")

        '   Dim a As String = ChromiumWebBrowser1.ToString()


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        ' DIABAZEI THN SELIDA
        '  Dim ts As Task = getSource()



        Dim ff As String = Application.StartupPath + "/currentSource.txt"
        '  Exit Sub
        '  MsgBox("OK")


        ' Thread.sleep


        '   <a href="/merchants/orders/   find
        '   <a href="/merchants/orders/open">
        '   <a href="/merchants/orders/210129-5519019">210129-5519019</a>
        '    Dim f As String = f_f


        'Dim fileReader As System.IO.StreamReader
        'fileReader =
        'My.Computer.FileSystem.OpenTextFileReader(ff)
        '   Dim stringReader As String
        '  StringReader = fileReader.ReadLine()



        '--------------------------------- ΦΟΡΤΩΝΩ ΤΟ HTML SE PINAKA --------------------------------------
        Dim pin(10000) As String
        Dim aa As String = "/merchants/orders/"
        Dim a2 As String = "<a href="
        Dim c As String, nn As Integer
        Dim nc As Integer = 0
        FileOpen(1, ff, OpenMode.Input)
        Do While Not EOF(1)

            nc = nc + 1 : pin(nc) = LineInput(1)
            ' c = LineInput(1)
            'nn = InStr(c, aa)
            'If nn > 0 And InStr(c, a2) > 0 Then
            '    Dim gg As String = Mid(c, nn + 18, 14)
            '    If InStr(gg, "-") > 0 Then
            '        ListBox1.Items.Add(gg)
            '    End If
            'End If



        Loop
        Dim n1, n2, n3, n4 As Integer

        FileClose(1)




        'ΑΠΟ ΤΟΝ ΠΑΡΑΠΑΝΩ ΠΙΝΑΚΑ ΞΕΧΩΡΙΖΨ ΤΙΣ ΠΑΡΑΓΓΕΛΙΕΣ
        Dim ONO As String, POSO As String

        For k As Integer = 1 To nc
            c = pin(k)
            nn = InStr(c, aa)
            If nn > 0 And InStr(c, a2) > 0 Then
                Dim gg As String = Mid(c, nn + 18, 14)
                If InStr(gg, "-") > 0 Then

                    n1 = InStr(pin(k + 4), ">") '   <span class="name">ΑΡΕΤΗ ΚΩΤΣΟΠΟΥΛΟΥ</span>
                    n2 = InStr(pin(k + 4), "</span>")
                    If n2 > n1 Then
                        ONO = Mid(Mid(pin(k + 4), n1 + 1, n2 - n1 - 1) + Space(50), 1, 35)
                    Else
                        ONO = Space(35)
                    End If


                    n3 = InStr(pin(k + 8), ">") '   <span class="name">ΑΡΕΤΗ ΚΩΤΣΟΠΟΥΛΟΥ</span>
                    n4 = InStr(pin(k + 8), "</div>")  '   <div>54,75 €</div>

                    If n4 > n3 Then
                        POSO = Mid(Mid(pin(k + 8), n3 + 1, n4 - n3 - 1) + Space(50), 1, 7)

                    Else
                        POSO = "       "
                    End If





                    ListBox1.Items.Add(gg + " " + ONO + " " + POSO)
                End If
            End If




        Next




    End Sub


    Private Async Function getSource() As Task
        Dim source As String = Await ChromiumWebBrowser1.GetBrowser().MainFrame.GetSourceAsync()

        Dim f As String
        f = Application.StartupPath + "/currentSource.txt"

        Dim wr As StreamWriter = New StreamWriter(f, False, System.Text.Encoding.Default)
        wr.Write(source)
        wr.Close()

        System.Diagnostics.Process.Start(f)


    End Function

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        '        <div Class='titlebar'><h1>Event Log started at 02/06/2015 13:07:30</h1></div>
        '<div Class='Information'><h1>02/06/2015 13:09:30 | Log has opened</h1></div>
        '<div Class='Interest'><h1>02/06/2015 13:13:03 | finished!</h1></div>
        '<div Class='Interest'><h1>02/06/2015 13:17:12 | finished!</h1></div>
        '<div Class='Interest'><h1>02/06/2015 13:21:35 | finished!</h1></div>
        '<div Class='Interest'><h1>02/06/2015 13:24:58 | finished!</h1></div>
        '<div Class='Warning'><h1>02/06/2015 17:04:33 | Failed to stop, retrying...</h1></div>
        '<div Class='Warning'><h1>02/06/2015 17:04:56 | Error during mix
        Dim f As String
        f = Application.StartupPath + "/currentSource.txt"

        Dim html As String = File.ReadAllText(f, Encoding.UTF8)

        Dim doc As New HtmlAgilityPack.HtmlDocument()
        doc.LoadHtml(html)


        Dim nodes As HtmlNode() = doc.DocumentNode.SelectNodes("//a").ToArray()

        For Each item As HtmlNode In nodes
            ' Console.WriteLine(item.InnerHtml)
            '  ListBox1.Items.Add(item.InnerHtml)
        Next
        ListBox1.Items.Add("=================================================")
        ' Dim interestDivs = doc.DocumentNode.SelectNodes("//div[contains(@class,'name')]")  //div[@id='div1']//a
        Dim nodes2 As HtmlNode() = doc.DocumentNode.SelectNodes("//div[@class='contact-card']//span").ToArray()

        If Not nodes2 Is Nothing Then
            For Each item As HtmlNode In nodes2
                Console.WriteLine(item.InnerHtml)
                ListBox1.Items.Add(item.InnerHtml)
            Next
        End If



        ListBox1.Items.Add("=====dt ============================================")

        ' If doc.DocumentNode.SelectNodes("//dd").Count > 0 Then
        Try


            Dim nodes3 As HtmlNode() = doc.DocumentNode.SelectNodes("//dd").ToArray()
            For Each item As HtmlNode In nodes3
                ListBox1.Items.Add(item.InnerHtml)
            Next
            '  End If
        Catch ex As Exception

        End Try


        'class="cell-product">

        ListBox1.Items.Add("=====product ============================================")
        Dim nodes4 As HtmlNode() = doc.DocumentNode.SelectNodes("//td[@class='cell-product']//a").ToArray()
        If Not nodes4 Is Nothing Then
            For Each item As HtmlNode In nodes4
                Console.WriteLine(item.InnerHtml)
                ListBox1.Items.Add(item.InnerHtml)
            Next
        End If



        ListBox1.Items.Add("=====barcode ============================================")
        Dim nodes5 As HtmlNode() = doc.DocumentNode.SelectNodes("//td[@class='cell-product']//div").ToArray()
        If Not nodes5 Is Nothing Then
            For Each item As HtmlNode In nodes5
                ListBox1.Items.Add(item.InnerHtml)
            Next
        End If

        '<td class="cell-quantity">
        '           <span>1 ×</span>
        '      </td>

        ListBox1.Items.Add("=====posothta============================================")
        Dim nodes6 As HtmlNode() = doc.DocumentNode.SelectNodes("//td[@class='cell-quantity']//span").ToArray()
        If Not nodes6 Is Nothing Then
            For Each item As HtmlNode In nodes6
                ListBox1.Items.Add(item.InnerHtml)
            Next
        End If


        ListBox1.Items.Add("=====timh============================================")
        Dim nodes7 As HtmlNode() = doc.DocumentNode.SelectNodes("//td[@class='cell-money']").ToArray()
        If Not nodes7 Is Nothing Then
            For Each item As HtmlNode In nodes7
                ListBox1.Items.Add(item.InnerHtml)
            Next
        End If







        Exit Sub


        '<p class="name">
        Dim interestDivs = doc.DocumentNode.SelectNodes("//div[contains(@class,'name')]")
        Dim warningDivs = doc.DocumentNode.SelectNodes("//div[contains(@class,'address')]")
        Dim informationDivs = doc.DocumentNode.SelectNodes("//div[contains(@class,'communication')]")

        Dim lines = From div In interestDivs Select div.InnerText
        lines = lines.Concat(From div In warningDivs Select div.InnerText)
        lines = lines.Concat(From div In informationDivs Select div.InnerText)
        TextBox2.Lines = lines.ToArray()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim c As String
        '  order.Text = Mid(ListBox1.SelectedItem, 1, 14)



    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ExecuteSQLQuery("delete from orders")
        ExecuteSQLQuery("delete from orderdetails")


        ExecuteSQLQuery("select [order.code] as kod,shop_uid from orderdetails")
        ListBox3.Items.Clear()
        For k As Integer = 0 To sqlDT.Rows.Count - 1

            ListBox3.Items.Add(sqlDT.Rows(k)("kod") + "  κωδ¨:" + sqlDT.Rows(k)("shop_uid"))

        Next




    End Sub

    Private Sub ListBox1_Click(sender As Object, e As EventArgs) Handles ListBox1.Click

        If Mid(ListBox1.SelectedItem, 1, 2) = "**" Then
        Else


            order.Text = Mid(ListBox1.SelectedItem, 1, 14)
            If ListBox1.SelectedIndex >= 0 Then
                ListBox1.Items(ListBox1.SelectedIndex) = "**" + ListBox1.Items(ListBox1.SelectedIndex)
            End If


        End If


    End Sub
End Class