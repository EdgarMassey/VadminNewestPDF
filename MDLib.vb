Imports System.IO
Imports System.Data.Odbc
Module MDlib
    'Local values
    Public invoicepr As String, standardpr As String, packpr As String, deliverypr As String, billpr As String, postpr As String, paminpr As String
    Public lonbeskpr As String, telefaxpr As String, bestpr As String, timecardpr As String, defaultref As String, defaultDoc As String, stationtype As String
    Public calbok As String, kvittorems As String, fraktsedpr As String, lolliadrpr As String, farliggodspr As String, stdmoms As String, databasnamn As String
    Public kolliadrpr As String, offertpr As String, odbcsource As String, odbcsourcerr As String, stationsid As String, odbcsourcer As String, losen As String
    Public tid As String, spar As String, sokvag As String, pdfpr As String, sakerhet As Integer, ipadress As String, iii As Integer, foljnr(100) As String, docdatum As String
    'Grundinformation
    Public KlientID As String, vernr As String, Prognamn As String, today As String, bibliotek As String, nummer As String, ingmoms As String, doctyp As String
    Public rakenskapID As String, räkid As String, logonamn As String, Firmanamn As String, Postadress1 As String, Postadress2 As String, prodsok As String
    Public Postadress3 As String, besad1 As String, besad2 As String, firmaland As String, telefon As String, telefax As String, gsm As String
    Public epost As String, kontaktperson As String, password1 As String, password2 As String, password3 As String, password4 As String, password5 As String
    Public password6 As String, antkvitto As Integer, antalfolj As Integer, antfakt As Integer, antallonespec As Integer, antorderbekraft As Integer, grundlista As Integer
    Public specialversion As String, momsv(10) As Decimal, moms1 As Integer, moms2 As Integer, moms3 As Integer, drojsmalsranta As Decimal, fortrycktpapper As Integer
    Public fortrycktfot As Integer, bankgiro As String, postgiro As String, regnr As String, bankkonto As String, bankkonteringsnr As String
    Public stdforsaljning As String, utmoms As String, lagerup As String, telebankgiro As Integer, expavgift As Integer, avgiftsgrans As Integer, ettlangd As Integer
    Public ettbred As Integer, komponentup As Integer, komment1 As String, komment2 As String, sidolev As Integer, genast As Integer, EAN As Integer, Language As String
    Public faktfolj As Integer, girolinknr As String, girolinkfil As String, rabradfakt As String, arbavgift As Integer, ejprissatt As Integer, rantdagar As Integer
    Public rantbelopp As Integer, BasEU97 As Integer, ejavrund As Integer, ordsok As Integer, ejcapri As Integer, lagerred As Integer, garantitext As String
    Public a5kvitto As Integer, textradpos As Integer, orderformat As String, meddelande As String, nolllagerspar As Integer, vsp As Integer, levdatum As String
    Public timedummy As String, SMTPservernamn As String, SMTPUser As String, SMTPPassW As String, land As String, webadress As String, kp As Double, antal As Integer
    Public antallonebesked As Integer, banknamn As String, varref As String, ertordernr As String, markning As String, momstyp As Integer, ordertyp As String
    Public orderanteckning As String, extrarab As Double
    'kundregister
    Public prevkundnamn As String, kordning As String, kejaktiv As Double, utskriftstyp As String
    Public kkundnr As String, kkundnamn As String, kadress1 As String, kadress2 As String, docnummer As String
    Public kpostnr As String, kort As String, kland As String, ktelefon As String, ktelefax As String
    Public klev1 As String, klev2 As String, klev3 As String, krabattmall As String, kgrundnom As Integer
    Public kkomment1 As String, kkomment2 As String, kkontaktman As String, kleveranssatt As String
    Public kleveransvilkor As String
    Public kkreditdagar As Integer, krabatt As Double, ksaljare As String, kkreditstatus As String
    Public kkundgrupp As String, kkundtyp As String, kdistrikt As String, kregnr As String
    Public kvaluta As String, kfakturatyp As String, kfakturaperiod As String, kkreditgrans As String
    Public kutnyttjad As Double, klosen As String, kepost As String, ksenfakt As String
    Public kantperar As Integer, ksendebn As Double, kkundinfo As String, kfaktlevtyp As String
    'Produktregister
    Public pantiforpack As String, pEnhet As String, fordatum As String, fordrojd As String, pordning As String
    Public pprodnr As String, pprodnamn As String, pkonto As String, putpr As Double, prantal As Double
    Public pinpris As Double, pgrinpris As Double, pantal As Double, pprisfaktor As Double, prest As Double
    Public putpris1 As Double, putpris2 As Double, putpris3 As Double, putpris4 As Double, putpris As Double
    Public putpris5 As Double, putpris6 As Double, putpris7 As Double, pantiförpack As String
    Public pgrundpris As Double, pgrundrab As Double, prabattfran As String, prabatttill As String
    Public pejrabatt As Integer, pvikt As Double, peankod As String, pkomment1 As String
    Public prabfron As String, prabtill As String, plageradress As String, prabattpris As Double
    Public pleverantor As String, pbestnr As String, pbestpunkt As Double, pbladderaktiv As Integer
    Public pprodukttyp As String, tbestpunkt As Double, psumma As Double, prabatt As Double
    Public pproduktgrupp As String, pbrytpunkt As Double, pbrytlista As Double
    Public pprislista As Integer, pbildid As String, pkomment As String, tillfpris As Double
    Public prablista As Integer, plongnamn As String, pundergrupp As String, pejwebaktiv As Integer
    Public pnyhet As Integer, perbjudande As Integer
    Public innetyper(200) As String, typerklar As String, altlevstatus As String, altlevprodukt As String
    Public rabmall(90, 10) As String
    'kundreskontra
    Public krfakturanummer As String, krFakturadatum As String, krForfallodatum As String, krAntalpaminnelser As Integer
    Public krKundnamn As String, krKundnummer As String, krAdress1 As String, krAdress2 As String, krPostnummer As String, krOrt As String
    Public krLand As String, krKontaktperson As String, krkreditdagar As Integer, krTelefon As String, krTelefax As String
    Public krKomment1 As String, krKomment2 As String, krValuta As String, krFakturabelopp As Double, krDelbetalningar As Integer
    Public krBetDatum As String, krBetid As String, krBetTyp As String, krRantafakt As String, krxmoms As String
    'Standard
    Public fran As String, skattear As String, semesterarstart As String, odbclosen As String, htmhuv As String, rr As Integer
    Public blackPen As New Pen(Color.Black, 1), rmax As Integer, franprog As String
    Public BlackFont As New SolidBrush(Color.Black)
    Public SmallFontB As New Font("Times New Roman", 10, FontStyle.Bold)
    Public SmallFontR As New Font("Times New Roman", 10, FontStyle.Regular)
    Public OrdFontR As New Font("Times New Roman", 12, FontStyle.Regular)
    Public OrdFontB As New Font("Times New Roman", 12, FontStyle.Bold)
    Public LargeFontB As New Font("Times New Roman", 14, FontStyle.Bold)
    Public LargeFontR As New Font("Times New Roman", 14, FontStyle.Regular)
    Function kommatillpunkt(ByVal nummer As String)
        If InStr(nummer, ",") > 0 Then
            Mid$(nummer, InStr(nummer, ","), 1) = "."
            'Stop
        End If
        kommatillpunkt = nummer
    End Function
    Function BLANKBORT(ByVal blankkoll As String)
        On Error Resume Next
        Dim l5 As Integer
        For l5 = 1 To 50
            If Len(blankkoll) > 0 Then
                If Asc(Right$(blankkoll, 1)) < 33 Then
                    blankkoll = Left$(blankkoll, Len(blankkoll) - 1)
                End If
                If Asc(Left$(blankkoll, 1)) < 33 Then
                    blankkoll = Right$(blankkoll, Len(blankkoll) - 1)
                End If
            End If
        Next l5
        BLANKBORT = blankkoll
    End Function
    Public Function parserad(ByVal rad As String)
        If Left$(rad, 7) = "Invoice" Then invoicepr = BLANKBORT(Mid$(rad, 9, 40))
        If Left$(rad, 8) = "Standard" Then standardpr = BLANKBORT(Mid$(rad, 10, 40))
        If Left$(rad, 4) = "Pack" Then packpr = BLANKBORT(Mid$(rad, 6, 40))
        If Left$(rad, 8) = "Delivery" Then deliverypr = BLANKBORT(Mid$(rad, 10, 40))
        If Left$(rad, 4) = "Bill" Then billpr = BLANKBORT(Mid$(rad, 6, 40))
        If Left$(rad, 3) = "PDF" Then pdfpr = BLANKBORT(Mid$(rad, 5, 40))
        If Left$(rad, 7) = "PostPak" Then postpr = BLANKBORT(Mid$(rad, 9, 40))
        If Left$(rad, 8) = "Reminder" Then paminpr = BLANKBORT(Mid$(rad, 10, 40))
        If Left$(rad, 8) = "Wagespec" Then lonbeskpr = BLANKBORT(Mid$(rad, 10, 40))
        If Left$(rad, 7) = "Telefax" Then telefaxpr = BLANKBORT(Mid$(rad, 9, 40))
        If Left$(rad, 7) = "Bestskr" Then bestpr = BLANKBORT(Mid$(rad, 9, 40))
        If Left$(rad, 8) = "Timecard" Then timecardpr = BLANKBORT(Mid$(rad, 10, 40))
        If Left$(rad, 10) = "DefaultRef" Then defaultref = BLANKBORT(Mid$(rad, 12, 40))
        If Left$(rad, 10) = "DefaultDoc" Then defaultDoc = BLANKBORT(Mid$(rad, 12, 40))
        If Left$(rad, 11) = "StationType" Then stationtype = BLANKBORT(Mid$(rad, 13, 40))
        If Left$(rad, 8) = "Calender" Then calbok = BLANKBORT(Mid$(rad, 10, 20))
        If Left$(rad, 10) = "Kvittorems" Then kvittorems = BLANKBORT(Mid$(rad, 12, 40))
        If Left$(rad, 10) = "Fraktsedel" Then fraktsedpr = BLANKBORT(Mid$(rad, 12, 40))
        If Left$(rad, 11) = "Kolliadress" Then kolliadrpr = BLANKBORT(Mid$(rad, 13, 40))
        If Left$(rad, 10) = "Farliggods" Then farliggodspr = BLANKBORT(Mid$(rad, 12, 40))
        If Left$(rad, 7) = "Stdmoms" Then stdmoms = BLANKBORT(Mid$(rad, 9, 1))
        If Left$(rad, 12) = "Databasename" Then databasnamn = BLANKBORT(Mid$(rad, 14, 40))
        If Left$(rad, 6) = "Offert" Then offertpr = BLANKBORT(Mid$(rad, 8, 40))

        If Left$(rad, 10) = "ODBCSource" Then
            odbcsourcer = BLANKBORT(Mid$(rad, 12, 11))
            odbcsourcerr = odbcsourcer
            odbcsource = odbcsourcer
        End If
        If Left$(rad, 10) = "StationsID" Then stationsid = BLANKBORT(Mid$(rad, 12, 20))
        If Left$(rad, Len("ClientID")) = "ClientID" Then KlientID = BLANKBORT(Mid$(rad, Len("ClientID") + 2, 20))
        If Left$(rad, 1) = "L" Then
            losen = BLANKBORT(Mid$(rad, Len("Lösen") + 2, 20))
        End If
        If Left$(rad, Len("Tid")) = "Tid" Then tid = BLANKBORT(Mid$(rad, Len("Tid") + 2, 20))
        If Left$(rad, Len("Spar")) = "Spar" Then spar = BLANKBORT(Mid$(rad, Len("Spar") + 3, 4))
        parserad = "1"

    End Function
    Public Function nullhantering(ByVal varde As Object, ByVal typ As String) As Object
        On Error Resume Next
        Select Case VarType(varde)
            Case 0 : If typ = "S" Then nullhantering = "" Else nullhantering = 0
            Case 1 : If typ = "S" Then nullhantering = "" Else nullhantering = 0
            Case 8
                If typ = "S" Then
                    nullhantering = varde
                Else
                    If Len(varde) = 0 Then varde = "0"
                    If varde = " " Then varde = "0"
                    nullhantering = CDbl(varde)
                End If
            Case Else : nullhantering = varde
        End Select
    End Function
    Public Function rightpad(ByVal siff As String)
        On Error Resume Next
        If CDbl(siff) > 999 Then
            rightpad = siff.PadLeft(21 - siff.Length)
        Else
            rightpad = siff.PadLeft(20 - siff.Length)
        End If
    End Function
    
    
    Function getnummer(ByVal typ As String, ByVal nummerserie As String)
  
    End Function
   
    

    Function hamptakund(ByVal typ As String)
        Dim cn As OdbcConnection, mySQL As String
        Dim connStr As String
        connStr = "DSN=" + odbcsourcer + "; Database=" + databasnamn + ";Uid=v2000;Pwd=" + odbclosen
        cn = New OdbcConnection(connStr)
        cn.Open()
        mySQL = "SELECT Top 2 * FROM kundreg"
        mySQL = mySQL + " WHERE ClientID = '" + KlientID
        If typ = "Start" Then
            mySQL = mySQL + "' ORDER BY kundnummer asc "
        ElseIf typ = "Fram" Then
            If kordning = "Nummer" Then
                mySQL = mySQL + "' And kundnummer > '" + kkundnr + "' ORDER BY kundnummer asc "
            Else
                mySQL = mySQL + "' And kundnamn > '" + kkundnamn + "' ORDER BY kundnamn asc "
            End If

        ElseIf typ = "Bak" Then
            If kordning = "Nummer" Then
                mySQL = mySQL + "' And kundnummer < '" + kkundnr + "' ORDER BY kundnummer desc "
            Else
                mySQL = mySQL + "' And kundnamn < '" + kkundnamn + "' ORDER BY kundnamn desc "
            End If
        ElseIf typ = "LikaNr" Then
            mySQL = mySQL + "' And kundnummer = '" + kkundnr + "' ORDER BY kundnummer asc "
        ElseIf typ = "LikaNamn" Then
            mySQL = mySQL + "' And kundnamn  >= '" + kkundnamn + "' ORDER BY kundnamn asc "

        End If
        Dim myCmd As New OdbcCommand(mySQL, cn)
        Dim tabel As OdbcDataReader = myCmd.ExecuteReader(CommandBehavior.CloseConnection)
        'myCmd.ExecuteNonQuery()
        If tabel.HasRows = False Then
            Dim Message As String = "Posten finns inte"
            Dim Caption As String = "Upplysning"
            Dim Buttons As MessageBoxButtons = MessageBoxButtons.OK
            Dim Result As DialogResult
            Result = MessageBox.Show(Message, Caption, Buttons)
            If Result = 1 Then
                GoTo slut
            End If
        Else
            tabel.Read()
            '   Stop
            kkundnr = nullhantering(tabel("Kundnummer"), "S")
            kkundnamn = nullhantering(tabel("Kundnamn"), "S")
            prevkundnamn = kkundnr
            kadress1 = nullhantering(tabel("Adress1"), "S")
            kadress2 = nullhantering(tabel("Adress2"), "S")
            kpostnr = nullhantering(tabel("Postnummer"), "S")
            kort = nullhantering(tabel("Ort"), "S")

            kland = nullhantering(tabel("Land"), "S")
            ktelefon = nullhantering(tabel("Telefon"), "S")
            ktelefax = nullhantering(tabel("Telefax"), "S")
            klev1 = nullhantering(tabel("Levadress1"), "S")
            klev2 = nullhantering(tabel("Levadress2"), "S")
            klev3 = nullhantering(tabel("Levadress3"), "S")
            krabattmall = nullhantering(tabel("Rabattmall"), "S")
            kgrundnom = nullhantering(tabel("Grundnominering"), "T")
            kkomment1 = nullhantering(tabel("Komment1"), "S")
            kkomment2 = nullhantering(tabel("Komment2"), "S")
            kkontaktman = nullhantering(tabel("Kontaktperson"), "S")
            kleveranssatt = nullhantering(tabel("Leveranssatt"), "S")
            kkreditdagar = nullhantering(tabel("Kreditdagar"), "T")
            krabatt = nullhantering(tabel("Rabatt"), "T")
            ksaljare = nullhantering(tabel("Saljare"), "S")
            kkreditstatus = nullhantering(tabel("Kreditstatus"), "S")
            kkundgrupp = nullhantering(tabel("Kundgrupp"), "S")
            kkundtyp = nullhantering(tabel("Kundtyp"), "S")
            kdistrikt = nullhantering(tabel("Distrikt"), "S")
            kregnr = nullhantering(tabel("Regnr"), "S")
            kvaluta = nullhantering(tabel("Valuta"), "S")
            kfakturatyp = nullhantering(tabel("Fakturatyp"), "S")
            kfakturaperiod = nullhantering(tabel("Fakturaperiod"), "S")
            kkreditgrans = nullhantering(tabel("Kreditgrans"), "T")
            kutnyttjad = nullhantering(tabel("Utnyttjadkredit"), "T")
            klosen = nullhantering(tabel("losen"), "S")
            kepost = nullhantering(tabel("email"), "S")
            kejaktiv = nullhantering(tabel("ejaktiv"), "T")
            ksenfakt = nullhantering(tabel("Senfakt"), "S")
            kantperar = nullhantering(tabel("Serieantal"), "T")
            ksendebn = nullhantering(tabel("Seriesenast"), "T")
            kkundinfo = nullhantering(tabel("kundinfo"), "S")
            kfaktlevtyp = nullhantering(tabel("faktlevtyp"), "S")
            hamptakund = "JA"
        End If
slut:
        cn.Close()
        hamptakund = "Yes"

    End Function
    Function hamptaprodukt(ByVal typ As String)
        hamptaprodukt = "No"
        pinpris = 0
        pgrundpris = 0
        pgrundrab = 0
        pprisfaktor = 0
        pantal = 0
        pEnhet = ""
        pvikt = 0
        pkonto = ""
        peankod = ""
        pbestnr = ""
        pbestpunkt = 0
        plageradress = ""
        pleverantor = ""
        pantiforpack = ""
        pejrabatt = 0
        pnyhet = 0
        perbjudande = 0
        pproduktgrupp = ""
        pprodukttyp = ""
        putpris1 = 0
        putpris2 = 0
        putpris3 = 0
        putpris4 = 0
        putpris5 = 0
        putpris6 = 0
        putpris7 = 0
        prabfron = ""
        prabtill = ""
        prablista = 0
        pbrytpunkt = 0
        pbrytlista = 0
        pprislista = 0
        pejwebaktiv = 0
        pprislista = 0
        pbladderaktiv = 0
        plongnamn = ""
        pleverantor = ""







        Dim cn As OdbcConnection, mySQL As String
        Dim connStr As String
        connStr = "DSN=" + odbcsourcer + "; Database=" + databasnamn + ";Uid=v2000;Pwd=" + odbclosen
        cn = New OdbcConnection(connStr)
        cn.Open()
        mySQL = "SELECT Top 2 * FROM produktreg"
        mySQL = mySQL + " WHERE ClientID = '" + KlientID
        If typ = "Start" Then
            mySQL = mySQL + "' ORDER BY prodnummer asc "
        ElseIf typ = "Fram" Then
            If pordning = "Nummer" Then
                mySQL = mySQL + "' And prodnummer > '" + pprodnr + "' ORDER BY prodnummer asc "
            Else
                mySQL = mySQL + "' And prodnamn > '" + pprodnamn + "' ORDER BY prodnamn asc "
            End If

        ElseIf typ = "Bak" Then
            If pordning = "Nummer" Then
                mySQL = mySQL + "' And prodnummer < '" + pprodnr + "' ORDER BY prodnummer desc "
            Else
                mySQL = mySQL + "' And prodnamn < '" + pprodnamn + "' ORDER BY prodnamn desc "
            End If
        ElseIf typ = "LikaNr" Then
            If Len(pprodnr) = 13 Then
                mySQL = mySQL + "' And eankod = '" + pprodnr + "' ORDER BY prodnummer asc "
            Else
                If Right(pprodnr, 1) = "-" Then
                    mySQL = mySQL + "' And prodnummer > '" + pprodnr + "' ORDER BY prodnummer asc "
                Else
                    mySQL = mySQL + "' And prodnummer = '" + pprodnr + "' ORDER BY prodnummer asc "
                End If
            End If
        ElseIf typ = "LikaNamn" Then
                If Right(pprodnamn, 1) = "-" Then
                    mySQL = mySQL + "' And prodnamn  > '" + pprodnamn + "' ORDER BY prodnamn asc "
                Else
                    mySQL = mySQL + "' And prodnamn  >= '" + pprodnamn + "' ORDER BY prodnamn asc "
                End If
            End If

            Dim myCmd As New OdbcCommand(mySQL, cn)
            Dim tabel As OdbcDataReader = myCmd.ExecuteReader(CommandBehavior.CloseConnection)
            'myCmd.ExecuteNonQuery()
            If tabel.HasRows = False Then
                Dim Message As String = "Posten finns inte"
                Dim Caption As String = "Upplysning"
                Dim Buttons As MessageBoxButtons = MessageBoxButtons.OK
                Dim Result As DialogResult = 1
                'Result = MessageBox.Show(Message, Caption, Buttons)
                If Result = 1 Then
                    pprodnr = ""
                    pprodnamn = ""
                    GoTo slut
                End If
            Else
                tabel.Read()
                'Stop
                pprodnr = Format$(nullhantering(tabel("Prodnummer"), "S"))
                pprodnamn = nullhantering(tabel("Prodnamn"), "S")
                pinpris = nullhantering(tabel("Inpris"), "T")
                pgrundpris = nullhantering(tabel("grundpris"), "T")
                pgrundrab = nullhantering(tabel("grundprisrab"), "T")
                pprisfaktor = nullhantering(tabel("prisfaktor"), "T")

                If pprisfaktor = 0 Then
                    pprisfaktor = 1
                End If

                pprisfaktor = Format(pprisfaktor, "###0.0##")
                pantal = nullhantering(tabel("antal"), "T")
                pEnhet = nullhantering(tabel("enhet"), "S")
                pvikt = nullhantering(tabel("vikt"), "T")
                pkonto = nullhantering(tabel("konto"), "S")
                peankod = nullhantering(tabel("eankod"), "S")
                pbestnr = nullhantering(tabel("bestallningsnummer"), "S")
                pbestpunkt = nullhantering(tabel("bestallningspunkt"), "T")
                plageradress = nullhantering(tabel("lageradress"), "S")
                pleverantor = nullhantering(tabel("leverantor"), "S")
                pantiforpack = nullhantering(tabel("antiforpack"), "S") '
                pejrabatt = nullhantering(tabel("ejrabatt"), "T")
                pnyhet = nullhantering(tabel("nyhet"), "T")
                perbjudande = nullhantering(tabel("erbjudande"), "T")
                pproduktgrupp = nullhantering(tabel("produktgrupp"), "S")
                pprodukttyp = nullhantering(tabel("produkttyp"), "S")
                putpris1 = nullhantering(tabel("utpris1"), "T")
                putpris2 = nullhantering(tabel("utpris2"), "T")
                putpris3 = nullhantering(tabel("utpris3"), "T")
                putpris4 = nullhantering(tabel("utpris4"), "T")
                putpris5 = nullhantering(tabel("utpris5"), "T")
                putpris6 = nullhantering(tabel("utpris6"), "T")
                putpris7 = nullhantering(tabel("utpris7"), "T")
                prabfron = nullhantering(tabel("rabattfran"), "S")
                prabtill = nullhantering(tabel("rabatttill"), "S")
                prablista = nullhantering(tabel("rabattpris"), "T")
                pbrytpunkt = nullhantering(tabel("brytpunkt"), "T")
                pbrytlista = nullhantering(tabel("brytlista"), "T")
                pprislista = nullhantering(tabel("prislista"), "T")
                pejwebaktiv = nullhantering(tabel("ejwebaktiv"), "T")
                pprislista = nullhantering(tabel("prislista"), "T")
                pbladderaktiv = nullhantering((tabel("bladderaktiv")), "T")
                plongnamn = nullhantering(tabel("longnamn"), "S")
                pleverantor = nullhantering(tabel("Leverantor"), "S")
                hamptaprodukt = "JA"
            End If
slut:
            cn.Close()


    End Function
    Function gettimestamp()
        gettimestamp = CStr(DateTime.Now.Ticks)
        While gettimestamp <= timedummy

            gettimestamp = CStr(DateTime.Now.Ticks)
            If gettimestamp > timedummy Then Exit While
        End While
        timedummy = gettimestamp
    End Function

    Function getkundpris(ByVal mall1 As String, ByVal mall2 As String, ByVal artnr As String)
        On Error Resume Next
        Dim n As Integer, nn As Integer, dummy As String, delpris As Double, pris As Double
        Dim cn As OdbcConnection, mySQL As String
        Dim connStr As String, utpris(10) As Double
        Dim myCmd As OdbcCommand
        Dim tabel As OdbcDataReader

        For n = 1 To 90
            For nn = 1 To 8
                rabmall(n, nn) = ""
            Next
        Next n
        connStr = "DSN=" + odbcsourcer + "; Database=" + databasnamn + ";Uid=v2000;Pwd=" + odbclosen
        cn = New OdbcConnection(connStr)
        cn.Open()
        mySQL = "SELECT * FROM rabattmall "
        mySQL = mySQL + "Where rabattmallnr = " & "'" + mall1 + "'"
        mySQL = mySQL + " And clientid = '" + KlientID + "' "
        myCmd = New OdbcCommand(mySQL, cn)
        tabel = myCmd.ExecuteReader(CommandBehavior.CloseConnection)
        n = 1
        While tabel.Read()

            rabmall(n, 1) = nullhantering(tabel("produktgrupp"), "S")
            rabmall(n, 2) = nullhantering(tabel("produktnummer"), "S")
            rabmall(n, 3) = nullhantering(tabel("prislista"), "S")
            rabmall(n, 4) = nullhantering(tabel("rabatt"), "S")
            rabmall(n, 5) = nullhantering(tabel("fastpris"), "S")
            rabmall(n, 6) = nullhantering(tabel("fron"), "S")
            rabmall(n, 7) = nullhantering(tabel("tom"), "S")
            n = n + 1
        End While
        mySQL = "SELECT * FROM rabattmall "
        mySQL = mySQL + "Where rabattmallnr = " & "'" + mall2 + "'"
        mySQL = mySQL + " And clientid = '" + KlientID + "' "
        myCmd = New OdbcCommand(mySQL, cn)
        tabel = myCmd.ExecuteReader(CommandBehavior.CloseConnection)
        While tabel.Read()
            rabmall(n, 1) = nullhantering(tabel("produktgrupp"), "S")
            rabmall(n, 2) = nullhantering(tabel("produktnummer"), "S")
            rabmall(n, 3) = nullhantering(tabel("prislista"), "S")
            rabmall(n, 4) = nullhantering(tabel("rabatt"), "S")
            rabmall(n, 5) = nullhantering(tabel("fastpris"), "S")
            rabmall(n, 6) = nullhantering(tabel("fron"), "S")
            rabmall(n, 7) = nullhantering(tabel("tom"), "S")
            n = n + 1
        End While
        cn.Close()

        pprodnr = artnr

        dummy = hamptaprodukt("LikaNr")
        utpris(1) = putpris1 : utpris(2) = putpris2 : utpris(3) = putpris3 : utpris(4) = putpris4 : utpris(5) = putpris5
        utpris(6) = putpris6 : utpris(7) = putpris7
        If grundlista = 2 Then
            delpris = putpris2
        ElseIf grundlista = 3 Then
            delpris = putpris3
        ElseIf grundlista = 4 Then
            delpris = putpris4
        ElseIf grundlista = 5 Then
            delpris = putpris5
        ElseIf grundlista = 6 Then
            delpris = putpris7
        ElseIf grundlista = 7 Then
            delpris = putpris7
        Else
            delpris = putpris1
        End If


        pris = putpris1

        If kgrundnom = 1 Then delpris = putpris1 : pris = putpris1
        If kgrundnom = 2 Then delpris = putpris2 : pris = putpris2
        If kgrundnom = 3 Then delpris = putpris3 : pris = putpris3
        If kgrundnom = 4 Then delpris = putpris4 : pris = putpris4
        If kgrundnom = 5 Then delpris = putpris5 : pris = putpris5
        If kgrundnom = 6 Then delpris = putpris6 : pris = putpris6
        If kgrundnom = 7 Then delpris = putpris7 : pris = putpris7

        getkundpris = putpris1

        For nn = 1 To n
            If rabmall(nn, 1) = pproduktgrupp Then
                pris = CDbl(utpris(CDbl(rabmall(nn, 3))))

                If CDbl(nullhantering(rabmall(nn, 4), "T")) > 0 Then
                    pris = pris - (pris * (CDbl(rabmall(nn, 4)) / 100))
                End If

                GoTo slutloop
            End If

            If rabmall(nn, 2) = artnr Then
                If CDbl(nullhantering(rabmall(nn, 3), "T")) Then
                    pris = CDbl(utpris(CDbl(rabmall(nn, 3))))
                    If CDbl(nullhantering(rabmall(nn, 4), "T")) > 0 Then
                        pris = pris - (pris * (CDbl(rabmall(nn, 4)) / 100))
                    End If
                End If
                If rabmall(nn, 5) > 0 Then
                    pris = CDbl(rabmall(nn, 5))
                End If

                GoTo slutloop
            End If



slutloop:
            If pris < delpris Then delpris = pris

        Next



        getkundpris = delpris
    End Function
    Function skickatillepostrelay(ByVal emailfrom As String, ByVal emailto As String, ByVal emailsubject As String, ByVal emailtext As String, ByVal emailid As String)

        Dim cn As OdbcConnection, mySQL As String
        Dim connStr As String, falt As String, varden As String
        connStr = "DSN=" + odbcsourcer + "; Database=" + "Emailrelay" + ";Uid=v2000;Pwd=" + odbclosen
        cn = New OdbcConnection(connStr)
        cn.Open()
        emailtext = tabortdubblastrings(emailtext)
        falt = "" : varden = ""
        mySQL = "INSERT INTO inkommandeepost "
        falt = falt + "Mytimestamp,"
        varden = varden + "'" + gettimestamp() + "',"
        falt = falt + "fromsystem,"
        varden = varden + "'" + KlientID + " - " + emailid + "',"
        falt = falt + "Regdatum,"
        varden = varden + "'" + CStr(Now) + "',"
        falt = falt + "epostfrom,"
        varden = varden + "'" + emailfrom + "',"
        falt = falt + "eposttill,"
        varden = varden + "'" + emailto + "',"
        falt = falt + "epostsubject,"
        varden = varden + "'" + emailsubject + "',"
        falt = falt + "epostbody"
        varden = varden + "'" + emailtext + "'"
        mySQL = mySQL & "(" & falt & ") VALUES (" & varden & ");"
        Dim myCmd As New OdbcCommand(mySQL, cn)
        myCmd.ExecuteNonQuery()

        cn.Close()

        skickatillepostrelay = "Yes"
    End Function
    Function tabortdubblastrings(ByVal nummer As String)

        While InStr(nummer, "'") > 0
            Mid$(nummer, InStr(nummer, "'"), 1) = "§"

        End While

        tabortdubblastrings = nummer
    End Function
    Function avrund(ByVal PRIS) As Object
        On Error Resume Next
        'If PRIS > 0 Then Stop
        If Len(kvaluta) > 1 And kvaluta <> "SEK" Then
            avrund = PRIS
        Else
            If ejavrund = False Then
                PRIS = Int((PRIS) + 0.5)
            End If
            avrund = PRIS
        End If
    End Function
    Function tabortblankochbind(pnr As String)

        While InStr(pnr, "-") > 0
            pnr = Strings.Left$(pnr, InStr(pnr, "-") - 1) + Strings.Right$(pnr, Len(pnr) - InStr(pnr, "-"))
        End While
        'Avsändarens bankgiro UTAN bindestreck
        While InStr(pnr, " ") > 0
            pnr = Strings.Left$(pnr, InStr(pnr, " ") - 1) + Strings.Right$(pnr, Len(pnr) - InStr(pnr, " "))
        End While                                 'Avsändarens bankgiro UTAN MELLANSLAG
        tabortblankochbind = pnr
    End Function
    Function skapamapp(mappnamn As String)

        Dim di As DirectoryInfo = New DirectoryInfo(mappnamn)
        skapamapp = "Skapar"
        Try
            If di.Exists Then
                Return "Finns"
            End If

            ' Try to create the directory.
            di.Create()
        Catch e As Exception
            MessageBox.Show("The process failed: {0}", e.ToString())
        End Try
    End Function
    Function avrund100(PRIS) As Object
        On Error Resume Next
        If ejavrund = False Then
            PRIS = Int((PRIS * 100) + 0.499) / 100
        End If
        avrund100 = PRIS
    End Function
   
End Module
