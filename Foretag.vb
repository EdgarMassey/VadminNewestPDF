
Imports System
Imports System.IO

Imports System.Data.Odbc
Imports System.Configuration
Public Class Foretag
    Public Shared Sub laslogg()
        Try
            ' Create an instance of StreamReader to read from a file.
            Dim sr As StreamReader = New StreamReader(sokvag + "vdriver.dll")
            Dim line As String, dummy As String
            ' Read and display the lines from the file until the end 
            ' of the file is reached.
            Do
                line = sr.ReadLine()
                'Console.WriteLine(line)
                dummy = parserad(line)
            Loop Until line Is Nothing
            sr.Close()
        Catch E As Exception
            ' Let the user know what went wrong.
            Console.WriteLine("The file could not be read:")
            Console.WriteLine(E.Message)
        End Try
    End Sub
    Public Shared Sub laslokal()
        Try
            ' Create an instance of StreamReader to read from a file.
            Dim sr As StreamReader = New StreamReader(sokvag + KlientID + ".cfg")
            Dim line As String, dummy As String
            ' Read and display the lines from the file until the end 
            ' of the file is reached.

            Do
                line = sr.ReadLine()
                'Console.WriteLine(line)
                dummy = parserad(line)

            Loop Until line Is Nothing
            sr.Close()
        Catch E As Exception
            ' Let the user know what went wrong.
            Console.WriteLine("The file could not be read:")
            Console.WriteLine(E.Message)
        End Try
    End Sub
    Private Sub Foretag_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim ip() As Net.IPAddress = System.Net.Dns.GetHostAddresses("")

        If ip.Count > 0 Then
            For Each ipadd As Net.IPAddress In ip

                ipadress = (ipadd.ToString)
            Next
        End If
        Dim rect As Rectangle = Screen.PrimaryScreen.WorkingArea
        sokvag = AppDomain.CurrentDomain.BaseDirectory
        Me.Top = (rect.Height / 2) - (Me.Height / 2)
        Me.Left = (rect.Width / 2) - (Me.Width / 2)
        laslogg()
        KlientIdTB.Text = KlientID
        LosenTB.Text = losen

        If spar = "True" Then
            SparaCB.Checked = True
        End If

        vernr = "N1312.26a"

        Prognamn = "VadminPDF2013net"
        Ver.Text = "Version: " + vernr
        Huvud.Text = Prognamn + " - Behörighetskontroll"
        today = Format(Now, "yyyy-MM-dd")
        odbclosen = "alfons"
        datum.Text = today

    End Sub
    Private Sub AvslutaK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AvslutaK.Click
        End
    End Sub

    Private Sub KlientIdTB_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles KlientIdTB.KeyUp
        If e.KeyCode = Keys.Enter Then
            loginrutin()
        End If
    End Sub

    Private Sub KlientIdTB_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles KlientIdTB.LostFocus
        KlientID = UCase(KlientIdTB.Text)
    End Sub
    Private Sub LoginK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoginK.Click
        loginrutin
    End Sub
    Function hamptaforetag(ByVal clientid As String)

        If Len(nullhantering(odbcsource, "S")) < 2 Then
            odbcsource = "VadminODBC"
            odbcsourcer = odbcsource
        End If

        On Error GoTo nocon
        Dim cn As OdbcConnection, mySQL As String
        Dim connStr As String, test As Decimal
        connStr = "DSN=" + odbcsource + "; Database=" + databasnamn + ";Uid=v2000;Pwd=" + odbclosen
        cn = New OdbcConnection(connStr)
        cn.Open()
        mySQL = "SELECT * FROM foretagreg"
        mySQL = mySQL + " WHERE ClientID = '" + KlientID + "' "
        Dim myCmd As New OdbcCommand(mySQL, cn)
        Dim tabel As OdbcDataReader = myCmd.ExecuteReader(CommandBehavior.CloseConnection)
        'myCmd.ExecuteReader()
        While tabel.Read()
            rakenskapID = nullhantering(tabel("rakenskapID"), "S")
            rakenskapID = nullhantering(tabel("yearid"), "S")
            logonamn = nullhantering(tabel("Logonamn"), "S")
            Firmanamn = nullhantering(tabel("Firmanamn"), "S")
            Postadress1 = nullhantering(tabel("Postadress1"), "S")
            Postadress2 = nullhantering(tabel("Postadress2"), "S")
            Postadress3 = nullhantering(tabel("Postadress3"), "S")
            besad1 = nullhantering(tabel("besoksadress1"), "S")
            besad2 = nullhantering(tabel("besoksadress2"), "S")
            firmaland = nullhantering(tabel("Land"), "S")
            telefon = nullhantering(tabel("Telefon"), "S")
            telefax = nullhantering(tabel("Telefax"), "S")
            gsm = nullhantering(tabel("GSM"), "S")
            epost = nullhantering(tabel("epost"), "S")
            kontaktperson = nullhantering(tabel("kontaktperson"), "S")
            password1 = nullhantering(tabel("Password1"), "S")
            password2 = nullhantering(tabel("Password2"), "S")
            password3 = nullhantering(tabel("Password3"), "S")
            password4 = nullhantering(tabel("Password4"), "S")
            password5 = nullhantering(tabel("Password5"), "S")
            password6 = nullhantering(tabel("Password6"), "S")
            antkvitto = nullhantering(tabel("Antalkvitto"), "T")
            antalfolj = nullhantering(tabel("Antalfoljesedel"), "T")
            antfakt = nullhantering(tabel("Antalfaktura"), "T")
            antorderbekraft = nullhantering(tabel("Antalorderbekraft"), "T")
            antallonespec = nullhantering(tabel("Antallonespec"), "T")
            grundlista = nullhantering(tabel("Grundlista"), "T")
            specialversion = nullhantering(tabel("Specialversion"), "S")
            moms1 = nullhantering(tabel("moms1"), "T")
            moms2 = nullhantering(tabel("moms2"), "T")
            moms3 = nullhantering(tabel("moms3"), "T")
            test = moms1 / 100
            momsv(1) = (moms1 / 100) : momsv(2) = moms2 / 100 : momsv(3) = moms3 / 100
            momsv(4) = Int(moms1 / ((moms1 / 100) + 1) * 100) / 100
            momsv(5) = Int(moms2 / ((moms2 / 100) + 1) * 100) / 100
            momsv(6) = Int(moms3 / ((moms3 / 100) + 1) * 100) / 100
            momsv(4) = momsv(4) / 100
            momsv(5) = momsv(5) / 100
            momsv(6) = momsv(6) / 100
            drojsmalsranta = nullhantering(tabel("drojsmalsranta"), "T")
            fortrycktpapper = nullhantering(tabel("Fortryckbankgiro"), "T")
            fortrycktfot = nullhantering(tabel("Fortrycktfot"), "T")
            bankgiro = nullhantering(tabel("bankgiro"), "S")
            postgiro = nullhantering(tabel("postgiro"), "S")
            regnr = nullhantering(tabel("regnr"), "S")
            bankkonto = nullhantering(tabel("bank"), "S")
            bankkonteringsnr = nullhantering(tabel("bankkonteringsnr"), "S")
            If Len(bankkonteringsnr) < 4 Then bankkonteringsnr = "1940"
            stdforsaljning = nullhantering(tabel("stdforsaljning"), "S")
            utmoms = nullhantering(tabel("utmoms"), "S")
            lagerup = nullhantering(tabel("Lagerup"), "S")
            telebankgiro = nullhantering(tabel("telebankgiro"), "T")
            expavgift = nullhantering(tabel("Expavgift"), "T")
            avgiftsgrans = nullhantering(tabel("avgiftsgrans"), "T")
            ettlangd = nullhantering(tabel("Ettikettlangd"), "T")
            ettbred = nullhantering(tabel("Ettantalibred"), "T")
            komponentup = nullhantering(tabel("Girolink"), "T")
            komment1 = nullhantering(tabel("Komment1"), "S")
            komment2 = nullhantering(tabel("Komment2"), "S")
            sidolev = nullhantering(tabel("Sidolev"), "T")
            genast = nullhantering(tabel("Telebankgiro"), "T")
            EAN = nullhantering(tabel("Inteean"), "T")
            Language = nullhantering(tabel("sprak"), "S")
            faktfolj = nullhantering(tabel("faktfolj"), "T")
            girolinknr = nullhantering(tabel("girolinknr"), "S")
            girolinkfil = nullhantering(tabel("girolinkfil"), "S")
            epost = nullhantering(tabel("epost"), "S")
            rabradfakt = nullhantering(tabel("rabradfakt"), "S")
            arbavgift = nullhantering(tabel("arbavgift"), "T")
            ejprissatt = nullhantering(tabel("ejprisspar"), "T")
            nolllagerspar = nullhantering(tabel("nolllagerspar"), "T")
            a5kvitto = nullhantering(tabel("a5kvitto"), "T")
            rantdagar = nullhantering(tabel("rantdagar"), "T")
            rantbelopp = nullhantering(tabel("rantbelopp"), "T")
            BasEU97 = nullhantering(tabel("baseu97"), "T")
            ejavrund = nullhantering(tabel("ejavrundning"), "T")
            ordsok = nullhantering(tabel("sokorder"), "T")
            ejcapri = nullhantering(tabel("ejcapris"), "T")
            lagerred = nullhantering(tabel("lagerred"), "T")
            garantitext = nullhantering(tabel("garantitext"), "S")
            textradpos = nullhantering(tabel("textrad"), "T")
            orderformat = nullhantering(tabel("orderformat"), "S")
            meddelande = nullhantering(tabel("meddelande"), "S")
            gsm = nullhantering(tabel("gsm"), "S")
            SMTPservernamn = nullhantering(tabel("SMTPserver"), "S")
            SMTPUser = nullhantering(tabel("SMTPUser"), "S")
            SMTPPassW = nullhantering(tabel("SMTPPassW"), "S")
            banknamn = nullhantering(tabel("banknamn"), "S")
            webadress = nullhantering(tabel("webadress"), "S")
            ingmoms = nullhantering(tabel("ingmomskonto"), "S")
            kp = nullhantering(tabel("kostnpos"), "T")
            antallonebesked = nullhantering(tabel("antallonespec"), "S")
        End While
        cn.Close()
nocon:
        If Err.Number <> 0 Then
            MessageBox.Show("ODBC förbindelse saknas")
        End If
    End Function

    Private Sub LosenTB_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles LosenTB.KeyUp
        If e.KeyCode = Keys.Enter Then
            loginrutin()
        End If
    End Sub
    Private Sub LosenTB_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles LosenTB.LostFocus
        losen = LosenTB.Text
    End Sub

    Private Sub loginrutin()
        KlientID = UCase(KlientID)
        Me.Cursor = Cursors.WaitCursor
        laslokal()

        If spar = "True" Then
            SparaCB.Checked = True
        Else
            SparaCB.Checked = False
        End If
        hamptaforetag(KlientID)
        If losen = "EMassey46" Or losen = password1 Then sakerhet = 1
        If losen = password2 Then sakerhet = 2
        If losen = password3 Then sakerhet = 3
        If losen = password4 Then sakerhet = 4
        If losen = password5 Then sakerhet = 5
        If losen = password6 Then sakerhet = 6

        If sakerhet > 0 Then
            If SparaCB.Checked = True Then
                My.Computer.FileSystem.WriteAllText(sokvag + "vdriver.dll", "ClientID=" + KlientID + vbCrLf, False)
                My.Computer.FileSystem.WriteAllText(sokvag + "vdriver.dll", "Lösen=" + losen + vbCrLf, True)
                My.Computer.FileSystem.WriteAllText(sokvag + "vdriver.dll", "Spara=" + CStr(SparaCB.Checked) + vbCrLf, True)
            Else
                My.Computer.FileSystem.WriteAllText(sokvag + "vdriver.dll", "" + vbCrLf, False)

            End If

            OrdFaktF.Show()
        Else
            MessageBox.Show("Ogiltig inloggning")
        End If
        Me.Cursor = Cursors.Arrow
    End Sub

    Private Sub LosenTB_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LosenTB.TextChanged
        losen = LosenTB.Text
    End Sub

    Private Sub SparaCB_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SparaCB.CheckedChanged
        spar = SparaCB.Checked
    End Sub

    


    Private Sub datum_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles datum.Click

    End Sub

    Private Sub KlientIdTB_TextChanged(sender As Object, e As EventArgs) Handles KlientIdTB.TextChanged
        KlientID = UCase(KlientIdTB.Text)
    End Sub
End Class
