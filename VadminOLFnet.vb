Imports System.Data.Odbc
Imports System.Net.Mail
Imports System.Net.Sockets
Public Class OrdFaktF
    Dim r As Integer, c As Integer, ordernummer As String, utskriftsort As String
    Dim lev1 As String, lev2 As String, lev3 As String, rr As Integer, rmax As Integer, y As Integer

    Private Sub OrdFaktF_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim rect As Rectangle = Screen.PrimaryScreen.WorkingArea
        Me.Top = (rect.Height / 2) - (Me.Height / 2)
        Me.Left = (rect.Width / 2) - (Me.Width / 2)
        Ver.Text = "Version: " + vernr
        Huvud.Text = Prognamn + " - Order/Fakturering"
        datum.Text = today
        KlientIdl.Text = "KlientID:" + KlientID
        odbcsource.Text = "ODBCsource:" + odbcsourcer
        databas.Text = "Databasnamn:" + databasnamn
        Gnamn.Text = Firmanamn
        nolla()
        FakturaRB.Checked = True
        VarrefTB.Select()
        LeveranssattTB.Text = hamptalevsatt(LeveranssattTB.Text)
        LevVilkorTB.Text = hamptalevsvilkor(LevVilkorTB.Text)
    End Sub


    Private Sub KundlistaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KundlistaToolStripMenuItem.Click
        KundlistaF.Show()
        KundlistaF.BringToFront()
    End Sub

    Private Sub KundregToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KundregToolStripMenuItem.Click
        If sakerhet < 4 Then
            KundregF.Show()
            KundregF.BringToFront()
        Else
            MessageBox.Show("Otillräklig behörighet!")
        End If

    End Sub


    Private Sub ProduktlistaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProduktlistaToolStripMenuItem.Click
        ProduktListaF.Show()
        ProduktListaF.BringToFront()
    End Sub

    Private Sub ProduktregToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProduktregToolStripMenuItem.Click
        If sakerhet < 4 Then
            ProdregF.Show()
            ProdregF.BringToFront()
        Else
            MessageBox.Show("Otillräklig behörighet!")
        End If
    End Sub

    Private Function makegrid(ByVal typ As String)
        DataGridView1.Columns.Clear()
        DataGridView1.Rows.Clear()
        If typ = "SimpleKontant" Then
            DataGridView1.Left = 0
            DataGridView1.Top = 56
            DataGridView1.Height = 314
        ElseIf typ = "SimpleKredit" Then
            DataGridView1.Left = 0
            DataGridView1.Top = 214
            DataGridView1.Height = 164
        End If
        With (Me.DataGridView1.Columns)
            .Add("", "ProdNr")
            .Add("", "ProdNamn")
            .Add("", "Enhet")
            .Add("", "PF")
            .Add("", "Antal")
            .Add("", "Apris")
            .Add("", "Summa")
            .Add("", "CaPris")
            .Add("", "Rabatt")
        End With
        With Me.DataGridView1.Rows
            For Me.r = 1 To 50
                .Add("")
            Next r
        End With
        r = 0
        With DataGridView1.ColumnHeadersDefaultCellStyle
            .BackColor = Color.Gray
            .ForeColor = Color.White
            .Font = New Font(DataGridView1.Font, FontStyle.Bold)
        End With
        DataGridView1.RowHeadersWidth = 12
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).Width = 300
        'enhet
        DataGridView1.Columns(2).Width = 50
        'prisfaktor
        DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns(3).Width = 30
        'antal
        DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns(4).Width = 50
        'apris
        DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns(5).Width = 60
        Me.DataGridView1.Columns.Item(5).DefaultCellStyle.Format = "n2"
        'summa
        DataGridView1.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns(6).Width = 70
        Me.DataGridView1.Columns.Item(6).DefaultCellStyle.Format = "n2"
        DataGridView1.Visible = True
        makegrid = ""
    End Function
    Public Function AddDays(ByVal value As Double) As DateTime
        AddDays = Now
    End Function

    Private Sub KontantRB_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KontantRB.CheckedChanged
        If KontantRB.Checked = True Then
            nolla()
            makegrid("SimpleKontant")
            kkundnr = "0000"
            kkundnamn = "Kontantförsäljning"
        Else
            makegrid("SimpleKredit")
        End If
    End Sub

    Private Sub AvslutaK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AvslutaK.Click
        Me.Close()
    End Sub

    Private Sub kundnrTB_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles kundnrTB.KeyUp
        If e.KeyCode = Keys.Enter Then
            Dim dummy As String
            dummy = hamptakund("LikaNr")
            skrivakundtillform()
            BladdertypL.Text = "Kundnummer"
        End If
    End Sub

    Private Sub skrivakundtillform()
        kundnrTB.Text = kkundnr
        KundnamnTB.Text = kkundnamn
        Adress1TB.Text = kadress1
        Adress2TB.Text = kadress2
        PostnrTB.Text = kpostnr
        OrtTB.Text = kort
        LandTB.Text = kland
        KreditDagarTB.Text = kkreditdagar
        KontaktmanTB.Text = kkontaktman
        exrabTB.Text = krabatt
        MarkningTB.Text = ""
        LevVilkorTB.Text = "Fritt"
        LeveranssattTB.Text = ""
        LevdatumTB.Text = today
        EpostTB.Text = kepost
    End Sub

    Private Sub KundnamnTB_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles KundnamnTB.KeyUp
        If e.KeyCode = Keys.Enter Then
            Dim dummy As String
            dummy = hamptakund("LikaNamn")
            skrivakundtillform()
            BladdertypL.Text = "Kundnamn"
            DataGridView1.Focus()
            DataGridView1.CurrentCell = DataGridView1.Item(0, 0)
        End If
    End Sub


    Private Sub kundnrTB_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kundnrTB.TextChanged
        kkundnr = kundnrTB.Text
    End Sub

    Private Sub KundnamnTB_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KundnamnTB.TextChanged
        kkundnamn = KundnamnTB.Text
    End Sub


    Private Sub BakB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BakB.Click
        Dim dummy As String
        If BladdertypL.Text = "Kundnummer" Then
            kordning = "Nummer"
        Else
            kordning = "Namn"
        End If
        dummy = hamptakund("Bak")
        skrivakundtillform()
    End Sub

    Private Sub FramB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FramB.Click
        Dim dummy As String
        If BladdertypL.Text = "Kundnummer" Then
            kordning = "Nummer"
        Else
            kordning = "Namn"
        End If
        dummy = hamptakund("Fram")
        skrivakundtillform()
    End Sub

    Private Sub DataGridView1_CellEnter(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEnter
        radraknare()
    End Sub

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        radraknare()
    End Sub

    Private Sub DataGridView1_CellMouseEnter(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellMouseEnter

    End Sub


    Private Sub DataGridView1_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged

    End Sub



    Private Sub DataGridView1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DataGridView1.KeyUp
        Dim tillfpris As Double
        On Error Resume Next

        c = DataGridView1.CurrentCell.ColumnIndex
        If e.KeyCode = Keys.Enter Then
            If c = 0 Then
                Dim dummy As String
                pprodnr = DataGridView1.Item(0, rr - 1).Value
                dummy = hamptaprodukt("LikaNr")

                DataGridView1.Item(0, rr - 1).Value = pprodnr
                DataGridView1.Item(1, rr - 1).Value = pprodnamn
                DataGridView1.Item(2, rr - 1).Value = pEnhet
                DataGridView1.Item(3, rr - 1).Value = pprisfaktor
                DataGridView1.Item(4, rr - 1).Value = 1
                If MomsTypTB.Text = 4 Then
                    tillfpris = (getkundpris(krabattmall, kkundnr, pprodnr))
                    tillfpris = tillfpris + (tillfpris * momsv(1))
                    DataGridView1.Item(5, rr - 1).Value = tillfpris
                    DataGridView1.Item(7, rr - 1).Value = putpris1 + (putpris1 * momsv(1))
                Else
                    tillfpris = (getkundpris(krabattmall, kkundnr, pprodnr))
                    DataGridView1.Item(5, rr - 1).Value = tillfpris
                    DataGridView1.Item(7, rr - 1).Value = putpris1 + (putpris1 * momsv(1))
                End If


                DataGridView1.Item(6, rr - 1).Value = (DataGridView1.Item(3, rr - 1).Value * DataGridView1.Item(5, rr - 1).Value * DataGridView1.Item(4, rr - 1).Value)
                DataGridView1.CurrentCell = DataGridView1.Item(4, rr - 1)
                BladdertypL.Text = "Produktnummer"
            ElseIf c = 1 Then
                Dim dummy As String
                pprodnamn = DataGridView1.Item(1, r - 1).Value
                dummy = hamptaprodukt("LikaNamn")
                DataGridView1.Item(0, rr).Value = pprodnr
                DataGridView1.Item(1, rr).Value = pprodnamn
                DataGridView1.Item(2, rr).Value = pEnhet
                DataGridView1.Item(3, rr).Value = pprisfaktor
                DataGridView1.Item(4, rr).Value = 1
                If MomsTypTB.Text = 4 Then
                    tillfpris = (getkundpris(krabattmall, kkundnr, pprodnr))
                    tillfpris = tillfpris + (tillfpris * momsv(1))
                    DataGridView1.Item(5, rr).Value = tillfpris
                    DataGridView1.Item(7, rr).Value = putpris1 + (putpris1 * momsv(1))
                Else
                    tillfpris = (getkundpris(krabattmall, kkundnr, pprodnr))
                    DataGridView1.Item(5, rr).Value = tillfpris
                    DataGridView1.Item(7, rr).Value = putpris1
                End If

                DataGridView1.Item(6, rr - 1).Value = (DataGridView1.Item(3, rr - 1).Value * DataGridView1.Item(5, rr - 1).Value * DataGridView1.Item(4, rr - 1).Value)
                DataGridView1.CurrentCell = DataGridView1.Item(4, rr - 1)
                BladdertypL.Text = "Produktnamn"
            ElseIf c = 3 Then
                DataGridView1.Item(6, rr - 1).Value = (DataGridView1.Item(3, rr - 1).Value * DataGridView1.Item(5, rr - 1).Value * DataGridView1.Item(4, rr - 1).Value)
                DataGridView1.CurrentCell = DataGridView1.Item(4, rr - 1)

            ElseIf c = 4 Then
                DataGridView1.Item(6, rr - 1).Value = (DataGridView1.Item(3, rr - 1).Value * DataGridView1.Item(5, rr - 1).Value * DataGridView1.Item(4, rr - 1).Value)
                DataGridView1.CurrentCell = DataGridView1.Item(5, rr - 1)
            ElseIf c = 5 Then
                DataGridView1.Item(6, rr - 1).Value = (DataGridView1.Item(3, rr - 1).Value * DataGridView1.Item(5, rr - 1).Value * DataGridView1.Item(4, rr - 1).Value)
                DataGridView1.CurrentCell = DataGridView1.Item(0, rr)
                summering()
            End If

        End If
    End Sub

    Private Sub DataGridView1_RowLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.RowLeave

        summering()
    End Sub

    Private Sub summering()
        On Error Resume Next
        Dim TOT As Double = 0, exrab As Double, momsbel As Double, attbet As Double

        If rr > 1 Then
            DataGridView1.Item(6, r - 1).Value = Format$((DataGridView1.Item(3, r - 1).Value * DataGridView1.Item(5, r - 1).Value * DataGridView1.Item(4, r - 1).Value), "###,##0.00")
        End If
        On Error Resume Next
        TOT = 0
        For i As Integer = 0 To rmax + 1
            TOT = TOT + DataGridView1.Rows(i).Cells(6).Value
        Next
        OrdsumL.Text = Format$(TOT, "###,##0.00")
        exrab = TOT * (exrabTB.Text / 100)
        ExrabL.Text = Format$(exrab, "###,##0.00")
        If MomsTypTB.Text = 1 Then
            momsbel = (TOT - exrab) * momsv(1)
            momsbelL.Text = Format$(momsbel, "###,##0.00")
            attbet = TOT - exrab + momsbel
            MOMSLabel.Text = "Moms belopp"
        ElseIf MomsTypTB.Text = 4 Then
            momsbel = (TOT - exrab) * momsv(4)
            momsbelL.Text = Format$(momsbel, "###,##0.00")
            attbet = TOT - exrab
            MOMSLabel.Text = "Därav moms"
        ElseIf MomsTypTB.Text = 0 Then

            momsbelL.Text = Format$(0, "###,##0.00")
            attbet = TOT - exrab
            MOMSLabel.Text = "Ingen moms"
        End If
        attbetL.Text = Format$(attbet, "###,##0.00")
    End Sub

    Private Sub KortCB_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KortCB.CheckedChanged
        If KortCB.Checked = True Then
            KortTB.Text = Format$((CDbl(attbetL.Text - CDbl(nullhantering(KontantTB.Text, "T")))), "###,##0.00")
        Else
            KortTB.Text = Format$(CDbl("0"), "###,##0.00")
        End If
    End Sub


    Private Sub KontantCB_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KontantCB.CheckedChanged
        If KontantCB.Checked = True Then
            KontantTB.Text = Format$((CDbl(attbetL.Text - CDbl(nullhantering(KortTB.Text, "T")))), "###,##0.00")
        Else
            KontantTB.Text = Format$(CDbl("0"), "###,##0.00")
        End If
    End Sub

    Private Sub UtskriftB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UtskriftB.Click
        If RakningRB.Checked = True Then
            If kvittorems = True Then
                skapaFHDfil()
            End If
            ProgressBar1.Value = 50
            getnummer("Next", "Nästa räkningsnummer")
            lagerupdateringrutin()
            ProgressBar1.Value = 70
            TillOrderArkiv()

            ProgressBar1.Value = 85
            suddaorder(OrdnrTB.Text)
            nolla()
            ProgressBar1.Value = 100

        ElseIf FakturaRB.Checked = True Then
            If FGCB.CheckState = 1 Then
                PrintPreviewDialog1.Document = PrintDocument1
                PrintPreviewDialog1.ShowDialog()
            Else
                PrintDocument1.PrinterSettings.Copies = antfakt
                PrintDocument1.Print()
            End If

            Dim Message As String = "Godkänns utskriften ?"
            Dim Caption As String = "Kontrollera"
            Dim Buttons As MessageBoxButtons = MessageBoxButtons.OKCancel
            Dim Result As DialogResult
            Result = MessageBox.Show(Message, Caption, Buttons)
            If Result = 1 Then
                Kundreskontrareg()
                ProgressBar1.Value = 50
                getnummer("Next", "Nästa fakturanummer")
                lagerupdateringrutin()
                ProgressBar1.Value = 70
                TillOrderArkiv()

                ProgressBar1.Value = 85
                If OrdertypTB.Text = "Återkommande" Then
                    sparaorder()
                Else
                    suddaorder(OrdnrTB.Text)
                End If
                suddafoljesedel(FoljnrTB.Text)
                nolla()
                ProgressBar1.Value = 100
            End If

        ElseIf FoljesedelRB.Checked = True Then
            If FGCB.CheckState = 1 Then
                PrintPreviewDialog1.Document = PrintDocument1
                PrintPreviewDialog1.ShowDialog()
            Else
                PrintDocument1.PrinterSettings.Copies = antfakt
                PrintDocument1.Print()
            End If
            Dim Message As String = "Godkänns utskriften ?"
            Dim Caption As String = "Kontrollera"
            Dim Buttons As MessageBoxButtons = MessageBoxButtons.OKCancel
            Dim Result As DialogResult
            Result = MessageBox.Show(Message, Caption, Buttons)
            If Result = 1 Then
                ProgressBar1.Value = 50
                getnummer("Next", "Nästa följesedelnummer")
                ProgressBar1.Value = 60
                sparafoljesedel()
                ProgressBar1.Value = 85
                nolla()
                ProgressBar1.Value = 100
            End If

        End If

    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim x As Integer, y As Integer, yy As Integer, point1 As Point, point2 As Point
        If FakturaRB.Checked = True Then
            'väljerskrivare
            PrintDocument1.PrinterSettings.PrinterName = invoicepr
            'header
            x = 40 : y = 8
            e.Graphics.DrawString(Firmanamn, LargeFontB, BlackFont, x, y)
            e.Graphics.DrawString("Faktura", LargeFontB, BlackFont, 580, y) : y = y + 22
            e.Graphics.DrawString(Postadress1, SmallFontR, BlackFont, x, y) : y = y + 20
            e.Graphics.DrawString(Postadress2, SmallFontR, BlackFont, x, y) : y = y + 20
            e.Graphics.DrawString(Postadress3, SmallFontR, BlackFont, x, y) : y = y + 20
            e.Graphics.DrawString("Telefon: " + telefon, SmallFontR, BlackFont, x, y) : y = y + 20
            e.Graphics.DrawString("Epost: " + telefax, SmallFontR, BlackFont, x, y) : y = y + 20
            nummer = getnummer("Kolla", "Nästa fakturanummer")
            y = 30
            e.Graphics.DrawString("Nummer: " + nummer, SmallFontR, BlackFont, 580, y) : y = y + 22
            e.Graphics.DrawString("Kund nr: " + kkundnr, SmallFontR, BlackFont, 580, y) : y = y + 22
            e.Graphics.DrawString("Datum: " + today, SmallFontR, BlackFont, 580, y)
            y = 170
            e.Graphics.DrawString("Leveranssätt: ", SmallFontR, BlackFont, 40, y)
            e.Graphics.DrawString(kleveranssatt, SmallFontR, BlackFont, 160, y) : y = y + 20 : yy = y
            e.Graphics.DrawString("Leveransvilkor: ", SmallFontR, BlackFont, 40, y)
            e.Graphics.DrawString("Fritt Upphärad", SmallFontR, BlackFont, 160, y) : y = y + 20
            e.Graphics.DrawString("Betalningsvilkor ", SmallFontR, BlackFont, 40, y)
            e.Graphics.DrawString(CStr(kkreditdagar) + " dagar", SmallFontR, BlackFont, 160, y) : y = y + 20
            e.Graphics.DrawString("Förfallodadum: ", SmallFontR, BlackFont, 40, y)
            AddDays(3D)
            e.Graphics.DrawString(CDate(today).AddDays(kkreditdagar), SmallFontR, BlackFont, 160, y) : y = y + 20
            e.Graphics.DrawString("Ordernummer: ", SmallFontR, BlackFont, 40, y)
            e.Graphics.DrawString(OrdnrTB.Text, SmallFontR, BlackFont, 160, y) : y = y + 20
            e.Graphics.DrawString("Märkning: ", SmallFontR, BlackFont, 40, y)
            e.Graphics.DrawString(MarkningTB.Text, SmallFontR, BlackFont, 160, y) : y = y + 20
            e.Graphics.DrawString("Regnummer: ", SmallFontR, BlackFont, 40, y)
            e.Graphics.DrawString(kregnr, SmallFontR, BlackFont, 160, y) : y = y + 20
            e.Graphics.DrawString("Leveransdatum: ", SmallFontR, BlackFont, 40, y)
            e.Graphics.DrawString(LevdatumTB.Text, SmallFontR, BlackFont, 160, y) : y = y + 20
            If specialversion = "MD" Then
                y = 220
            Else
                y = 140
            End If
            e.Graphics.DrawString(kkundnamn, OrdFontB, BlackFont, 440, y) : y = y + 20
            e.Graphics.DrawString(kadress1, SmallFontR, BlackFont, 440, y) : y = y + 20
            e.Graphics.DrawString(kadress2, SmallFontR, BlackFont, 440, y) : y = y + 20
            e.Graphics.DrawString(kpostnr + " " + kort, SmallFontR, BlackFont, 440, y) : y = y + 20
            e.Graphics.DrawString(kland, SmallFontR, BlackFont, 440, y) : y = y + 20
            If specialversion = "MD" Then
                y = 140
            Else
                y = 240
            End If
            e.Graphics.DrawString("Epost " + kepost, SmallFontR, BlackFont, 440, y) : y = y + 20
            e.Graphics.DrawString("Telefonnr. " + ktelefon, SmallFontR, BlackFont, 440, y) : y = y + 20
            'rubriker
            y = 340
            e.Graphics.DrawString("Prodnr:", SmallFontB, BlackFont, 32, y)
            e.Graphics.DrawString("Prodnamn", SmallFontB, BlackFont, 146, y)
            e.Graphics.DrawString("Enhet", SmallFontB, BlackFont, 460, y)
            e.Graphics.DrawString("Antal", SmallFontB, BlackFont, 540, y)
            e.Graphics.DrawString("ápris", SmallFontB, BlackFont, 620, y)
            e.Graphics.DrawString("Belopp", SmallFontB, BlackFont, 722, y)
            y = y + 20
            point1 = New Point(30, y) : point2 = New Point(780, y)
            e.Graphics.DrawLine(blackPen, point1, point2) : y = y - 10
            For i As Integer = 0 To rmax
                pprodnr = DataGridView1.Rows(i).Cells(0).Value
                pprodnamn = DataGridView1.Rows(i).Cells(1).Value
                pEnhet = DataGridView1.Rows(i).Cells(2).Value
                prantal = DataGridView1.Rows(i).Cells(4).Value
                putpris = DataGridView1.Rows(i).Cells(5).Value
                psumma = DataGridView1.Rows(i).Cells(6).Value
                y = y + 20
                e.Graphics.DrawString(pprodnr, SmallFontR, BlackFont, 32, y)
                e.Graphics.DrawString(pprodnamn, SmallFontR, BlackFont, 150, y)
                e.Graphics.DrawString(pEnhet, SmallFontR, BlackFont, 480, y)
                If prantal <> 0 Then
                    e.Graphics.DrawString(rightpad(Format$(prantal, "###,##0.0")), SmallFontR, BlackFont, 490, y)
                End If
                If putpris <> 0 Then
                    If putpris < 0 Then
                        e.Graphics.DrawString(rightpad(Format$(putpris, "###,##0.00")), SmallFontR, BlackFont, 587, y)
                    Else
                        e.Graphics.DrawString(rightpad(Format$(putpris, "###,##0.00")), SmallFontR, BlackFont, 580, y)
                    End If
                End If
                If psumma <> 0 Then
                    If psumma < 0 Then
                        e.Graphics.DrawString(rightpad(Format$(psumma, "###,##0.00")), SmallFontR, BlackFont, 702, y)
                    Else
                        e.Graphics.DrawString(rightpad(Format$(psumma, "###,##0.00")), SmallFontR, BlackFont, 700, y)
                    End If
                End If
            Next
            y = y + 40
            If CDbl(nullhantering((ExrabL.Text), "T")) <> 0 Then
                e.Graphics.DrawString("Ordersumma: ", SmallFontR, BlackFont, 570, y)
                e.Graphics.DrawString(rightpad(Format$(CDbl(OrdsumL.Text), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20

                e.Graphics.DrawString("Extra rabatt: ", SmallFontR, BlackFont, 570, y)
                e.Graphics.DrawString(rightpad(Format$(CDbl(ExrabL.Text), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20
            End If
            If CDbl(momsbelL.Text) <> 0 Then
                e.Graphics.DrawString("Momspliktig belopp: ", SmallFontR, BlackFont, 570, y)
                e.Graphics.DrawString(rightpad(Format$(CDbl(OrdsumL.Text) - CDbl(ExrabL.Text), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20

                e.Graphics.DrawString(MOMSLabel.Text, SmallFontR, BlackFont, 570, y)
                e.Graphics.DrawString(rightpad(Format$(CDbl(momsbelL.Text), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20
            End If
            e.Graphics.DrawString("Att Betala: ", SmallFontR, BlackFont, 570, y)
            e.Graphics.DrawString(rightpad(Format$(CDbl(attbetL.Text), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20
            'fot
            y = 1070
            e.Graphics.DrawString("Sida 1 ", SmallFontR, BlackFont, 40, y)
            e.Graphics.DrawString(komment1, SmallFontB, BlackFont, 320, y)
            e.Graphics.DrawString("Handläggare " + VarrefTB.Text, SmallFontR, BlackFont, 580, y)
            y = y + 20
            point1 = New Point(30, y) : point2 = New Point(780, y)
            e.Graphics.DrawLine(blackPen, point1, point2) : y = y - 12
            y = y + 20
            e.Graphics.DrawString("Bankgiro ", SmallFontB, BlackFont, 40, y)
            e.Graphics.DrawString("Reg.nummer", SmallFontB, BlackFont, 320, y)
            e.Graphics.DrawString("Plusgiro", SmallFontB, BlackFont, 660, y)
            y = y + 20
            e.Graphics.DrawString(bankgiro, SmallFontR, BlackFont, 40, y)
            e.Graphics.DrawString(regnr, SmallFontR, BlackFont, 320, y)
            e.Graphics.DrawString(postgiro, SmallFontR, BlackFont, 660, y)
        ElseIf FoljesedelRB.Checked = True Then

            'följesedlar
            'väljerskrivare
            PrintDocument1.PrinterSettings.PrinterName = deliverypr
            'header
            x = 40 : y = 8
            e.Graphics.DrawString(Firmanamn, LargeFontB, BlackFont, x, y)
            e.Graphics.DrawString("Följesedel", LargeFontB, BlackFont, 580, y) : y = y + 22
            e.Graphics.DrawString(Postadress1, SmallFontR, BlackFont, x, y) : y = y + 20
            e.Graphics.DrawString(Postadress2, SmallFontR, BlackFont, x, y) : y = y + 20
            e.Graphics.DrawString(Postadress3, SmallFontR, BlackFont, x, y) : y = y + 20
            e.Graphics.DrawString("Telefon: " + telefon, SmallFontR, BlackFont, x, y) : y = y + 20
            e.Graphics.DrawString("Epost: " + telefax, SmallFontR, BlackFont, x, y) : y = y + 20
            nummer = getnummer("Kolla", "Nästa följesedelnummer")
            FoljnrTB.Text = nummer
            y = 30
            e.Graphics.DrawString("Nummer: " + nummer, SmallFontR, BlackFont, 580, y) : y = y + 22
            e.Graphics.DrawString("Kund nr: " + kkundnr, SmallFontR, BlackFont, 580, y) : y = y + 22
            e.Graphics.DrawString("Datum: " + today, SmallFontR, BlackFont, 580, y)
            y = 170
            e.Graphics.DrawString("Leveranssätt: ", SmallFontR, BlackFont, 40, y)
            e.Graphics.DrawString(kleveranssatt, SmallFontR, BlackFont, 160, y) : y = y + 20 : yy = y
            e.Graphics.DrawString("Leveransvilkor: ", SmallFontR, BlackFont, 40, y)
            e.Graphics.DrawString("Fritt Upphärad", SmallFontR, BlackFont, 160, y) : y = y + 20
            'e.Graphics.DrawString("Betalningsvilkor ", SmallFontR, BlackFont, 40, y)
            'e.Graphics.DrawString(CStr(kkreditdagar) + " dagar", SmallFontR, BlackFont, 160, y) : y = y + 20
            e.Graphics.DrawString("Förfallodadum: ", SmallFontR, BlackFont, 40, y)
            AddDays(3D)
            e.Graphics.DrawString(CDate(today).AddDays(kkreditdagar), SmallFontR, BlackFont, 160, y) : y = y + 20
            e.Graphics.DrawString("Ordernummer: ", SmallFontR, BlackFont, 40, y)
            e.Graphics.DrawString(OrdnrTB.Text, SmallFontR, BlackFont, 160, y) : y = y + 20
            e.Graphics.DrawString("Märkning: ", SmallFontR, BlackFont, 40, y)
            e.Graphics.DrawString(MarkningTB.Text, SmallFontR, BlackFont, 160, y) : y = y + 20
            e.Graphics.DrawString("Regnummer: ", SmallFontR, BlackFont, 40, y)
            e.Graphics.DrawString(kregnr, SmallFontR, BlackFont, 160, y) : y = y + 20
            e.Graphics.DrawString("Leveransdatum: ", SmallFontR, BlackFont, 40, y)
            e.Graphics.DrawString(LevdatumTB.Text, SmallFontR, BlackFont, 160, y) : y = y + 20
            If specialversion = "MD" Then
                y = 220
            Else
                y = 140
            End If
            e.Graphics.DrawString(kkundnamn, OrdFontB, BlackFont, 440, y) : y = y + 20
            e.Graphics.DrawString(kadress1, SmallFontR, BlackFont, 440, y) : y = y + 20
            e.Graphics.DrawString(kadress2, SmallFontR, BlackFont, 440, y) : y = y + 20
            e.Graphics.DrawString(kpostnr + " " + kort, SmallFontR, BlackFont, 440, y) : y = y + 20
            e.Graphics.DrawString(kland, SmallFontR, BlackFont, 440, y) : y = y + 20
            If specialversion = "MD" Then
                y = 140
            Else
                y = 240
            End If
            e.Graphics.DrawString("Epost " + kepost, SmallFontR, BlackFont, 440, y) : y = y + 20
            e.Graphics.DrawString("Telefonnr. " + ktelefon, SmallFontR, BlackFont, 440, y) : y = y + 20
            'rubriker
            y = 340
            e.Graphics.DrawString("Prodnr:", SmallFontB, BlackFont, 32, y)
            e.Graphics.DrawString("Prodnamn", SmallFontB, BlackFont, 146, y)
            e.Graphics.DrawString("Enhet", SmallFontB, BlackFont, 460, y)
            e.Graphics.DrawString("Antal", SmallFontB, BlackFont, 540, y)
            'e.Graphics.DrawString("ápris", SmallFontB, BlackFont, 620, y)
            'e.Graphics.DrawString("Belopp", SmallFontB, BlackFont, 722, y)
            y = y + 20
            point1 = New Point(30, y) : point2 = New Point(780, y)
            e.Graphics.DrawLine(blackPen, point1, point2) : y = y - 10
            For i As Integer = 0 To rr
                pprodnr = DataGridView1.Rows(i).Cells(0).Value
                pprodnamn = DataGridView1.Rows(i).Cells(1).Value
                pEnhet = DataGridView1.Rows(i).Cells(2).Value
                prantal = DataGridView1.Rows(i).Cells(4).Value
                putpris = DataGridView1.Rows(i).Cells(5).Value
                psumma = DataGridView1.Rows(i).Cells(6).Value
                y = y + 20
                e.Graphics.DrawString(pprodnr, SmallFontR, BlackFont, 32, y)
                e.Graphics.DrawString(pprodnamn, SmallFontR, BlackFont, 150, y)
                e.Graphics.DrawString(pEnhet, SmallFontR, BlackFont, 480, y)
                If prantal <> 0 Then
                    e.Graphics.DrawString(rightpad(Format$(prantal, "###,##0.0")), SmallFontR, BlackFont, 490, y)
                End If
                If putpris <> 0 Then
                    If putpris < 0 Then
                        'e.Graphics.DrawString(rightpad(Format$(putpris, "###,##0.00")), SmallFontR, BlackFont, 587, y)
                    Else
                        'e.Graphics.DrawString(rightpad(Format$(putpris, "###,##0.00")), SmallFontR, BlackFont, 580, y)
                    End If
                End If
                If psumma <> 0 Then
                    If psumma < 0 Then
                        'e.Graphics.DrawString(rightpad(Format$(psumma, "###,##0.00")), SmallFontR, BlackFont, 702, y)
                    Else
                        'e.Graphics.DrawString(rightpad(Format$(psumma, "###,##0.00")), SmallFontR, BlackFont, 700, y)
                    End If
                End If
            Next
            y = y + 40
            
            'fot
            y = 1070
            e.Graphics.DrawString("Sida 1 ", SmallFontR, BlackFont, 40, y)
            e.Graphics.DrawString(komment1, SmallFontB, BlackFont, 320, y)
            e.Graphics.DrawString("Handläggare " + VarrefTB.Text, SmallFontR, BlackFont, 580, y)
            y = y + 20
            point1 = New Point(30, y) : point2 = New Point(780, y)
            e.Graphics.DrawLine(blackPen, point1, point2) : y = y - 12
            y = y + 20
            e.Graphics.DrawString("Bankgiro ", SmallFontB, BlackFont, 40, y)
            e.Graphics.DrawString("Reg.nummer", SmallFontB, BlackFont, 320, y)
            e.Graphics.DrawString("Plusgiro", SmallFontB, BlackFont, 660, y)
            y = y + 20
            e.Graphics.DrawString(bankgiro, SmallFontR, BlackFont, 40, y)
            e.Graphics.DrawString(regnr, SmallFontR, BlackFont, 320, y)
            e.Graphics.DrawString(postgiro, SmallFontR, BlackFont, 660, y)

        End If

    End Sub

    Private Sub KreditDagarTB_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KreditDagarTB.TextChanged
        kkreditdagar = KreditDagarTB.Text
    End Sub
    Private Sub Kundreskontrareg()
        Dim cn As OdbcConnection, mySQL As String, forfallodatum As String, momspro As Double
        Dim connStr As String, falt As String, varden As String
        connStr = "DSN=" + odbcsourcer + "; Database=" + databasnamn + ";Uid=v2000;Pwd=" + odbclosen
        cn = New OdbcConnection(connStr)
        cn.Open()
        mySQL = "DELETE FROM kundreg "
        mySQL = mySQL + " Where kundnummer = '" + kkundnr + "' "
        mySQL = mySQL + " And ClientID = '" + KlientID + "' "
        Dim myCmd As New OdbcCommand(mySQL, cn)
        'myCmd.ExecuteNonQuery()
        falt = "" : varden = ""
        mySQL = "INSERT INTO kundres "
        falt = falt + "clientid,"
        varden = varden + "'" + KlientID + "',"
        falt = falt + "Fakturanummer,"
        varden = varden + "'" + CStr(nummer) + "',"
        falt = falt + "Fakturadatum,"
        varden = varden + "'" + today + "',"
        AddDays(3D)
        forfallodatum = CDate(today).AddDays(kkreditdagar)
        falt = falt + "forfallodatum,"
        varden = varden + "'" + forfallodatum + "',"
        falt = falt + "Kundnamn,"
        varden = varden + "'" + KundnamnTB.Text + "',"
        falt = falt + "kundnummer,"
        varden = varden + "'" + kundnrTB.Text + "',"
        falt = falt + "adress1,"
        varden = varden + "'" + Adress1TB.Text + "',"
        falt = falt + "adress2,"
        varden = varden + "'" + Adress2TB.Text + "',"
        falt = falt + "postnummer,"
        varden = varden + "'" + PostnrTB.Text + "',"
        falt = falt + "ort,"
        varden = varden + "'" + OrtTB.Text + "',"
        falt = falt + "land,"
        varden = varden + "'" + LandTB.Text + "',"
        If specialversion = "Malmgren" Then
            falt = falt + "xmoms,"
            If CDbl(MomsTypTB.Text) <> 0 Then momspro = 25 Else momspro = 0
            varden = varden + "" + CStr(momspro) + ","
        End If
        falt = falt + "Kreditdagar,"
        varden = varden + "" + KreditDagarTB.Text + ","
        falt = falt + "telefon,"
        varden = varden + "'" + ktelefon + "',"
        falt = falt + "telefax,"
        varden = varden + "'" + kepost + "',"

        falt = falt + "Fakturabelopp"
        If Len(kvaluta) > 1 And kvaluta <> "SEK" Then
            'varden = varden + "" + kommatillpunkt(CStr(CDbl(attbetL.Text) * getvalutakurs(kvaluta))) + ""
        Else
            varden = varden + "" + kommatillpunkt(CStr(CDbl(attbetL.Text))) + ""
        End If
        mySQL = mySQL & "(" & falt & ") VALUES (" & varden & ");"
        myCmd = New OdbcCommand(mySQL, cn)
        myCmd.ExecuteNonQuery()
        'Stop
        cn.Close()
    End Sub

    Private Sub TillOrderArkiv()

        If FakturaRB.Checked = True Then
            Dim radkost As Double
            Dim cn As OdbcConnection, mySQL As String, dummy As String
            Dim connStr As String, falt As String, varden As String
            connStr = "DSN=" + odbcsourcer + "; Database=" + databasnamn + ";Uid=v2000;Pwd=" + odbclosen
            cn = New OdbcConnection(connStr)
            cn.Open()
            mySQL = "DELETE FROM kundreg "
            mySQL = mySQL + " Where kundnummer = '" + kkundnr + "' "
            mySQL = mySQL + " And ClientID = '" + KlientID + "' "
            Dim myCmd As New OdbcCommand(mySQL, cn)
            'myCmd.ExecuteNonQuery()
            For i As Integer = 0 To DataGridView1.RowCount - 1
                If Len((DataGridView1.Rows(i).Cells(0).Value)) < 1 And Len((DataGridView1.Rows(i).Cells(1).Value)) < 1 Then
                    GoTo slut2
                End If
                falt = "" : varden = ""
                mySQL = "INSERT INTO orderarkiv "
                falt = falt + "clientid,"
                varden = varden + "'" + KlientID + "',"
                falt = falt + "Mytimestamp,"
                varden = varden + "'" + gettimestamp() + "',"
                falt = falt + "Debittyp,"
                If KreditDagarTB.Text = 0 Then
                    varden = varden + "'" + "Kvitto" + "',"
                Else
                    varden = varden + "'" + "Faktura" + "',"
                End If
                falt = falt + "DebitID,"
                varden = varden + "'" + nullhantering((CStr(nummer)), "S") + "',"
                falt = falt + "Ordernummer,"
                varden = varden + "'" + nullhantering((OrdnrTB.Text), "S") + "',"
                falt = falt + "Foljesedelnummer,"
                varden = varden + "'" + nullhantering((FoljnrTB.Text), "S") + "',"
                falt = falt + "Regdatum,"
                varden = varden + "'" + today + "',"
                falt = falt + "regtid,"
                varden = varden + "'" + "1" + "',"
                falt = falt + "Erreferens,"
                varden = varden + "'" + KontaktmanTB.Text + "',"
                falt = falt + "Kundnamn,"
                varden = varden + "'" + kkundnamn + "',"
                falt = falt + "Kundnummer,"
                varden = varden + "'" + kkundnr + "',"
                falt = falt + "Adress1,"
                varden = varden + "'" + Adress1TB.Text + "',"
                falt = falt + "Adress2,"
                varden = varden + "'" + Adress2TB.Text + "',"
                falt = falt + "Postnummer,"
                varden = varden + "'" + PostnrTB.Text + "',"
                falt = falt + "Ort,"
                varden = varden + "'" + OrtTB.Text + "',"
                falt = falt + "Land,"
                varden = varden + "'" + LandTB.Text + "',"
                falt = falt + "Kreditdagar,"
                varden = varden + "" + KreditDagarTB.Text + ","
                falt = falt + "Telefon,"
                varden = varden + "'" + ktelefon + "',"
                falt = falt + "Telefax,"
                varden = varden + "'" + ktelefax + "',"
                falt = falt + "Valuta,"
                varden = varden + "'" + "SEK" + "',"
                If Len(DataGridView1.Rows(i).Cells(6).Value) < 1 Then

                    falt = falt + "Produktnamn"
                    varden = varden + "'" + nullhantering(DataGridView1.Rows(i).Cells(1).Value, "S") + "'"
                    GoTo slut
                End If
                falt = falt + "Produktnummer,"
                varden = varden + "'" + DataGridView1.Rows(i).Cells(0).Value + "',"
                falt = falt + "Produktnamn,"
                varden = varden + "'" + nullhantering(DataGridView1.Rows(i).Cells(1).Value, "S") + "',"
                falt = falt + "antal,"
                varden = varden + "" + kommatillpunkt(CStr(nullhantering(DataGridView1.Rows(i).Cells(4).Value, "T"))) + ","
                falt = falt + "apris,"
                If Len(kvaluta) > 1 And kvaluta <> "SEK" Then
                    'varden = varden + "" + kommatillpunkt(CStr(pri0) * getvalutakurs(kvaluta)) + ","
                Else
                    varden = varden + "" + kommatillpunkt(CStr(nullhantering(DataGridView1.Rows(i).Cells(5).Value, "T"))) + ","
                End If
                falt = falt + "prisfaktor,"
                varden = varden + "" + kommatillpunkt(CStr(nullhantering(DataGridView1.Rows(i).Cells(3).Value, "T"))) + ","
                falt = falt + "Antiforpackning,"
                varden = varden + "'" + nullhantering(pantiforpack, "S") + "',"
                falt = falt + "Kontonummer,"
                varden = varden + "'" + nullhantering(pkonto, "S") + "',"
                falt = falt + "Varreferens,"
                varden = varden + "'" + VarrefTB.Text + "',"
                falt = falt + "Leveranssatt,"
                varden = varden + "'" + nullhantering(LeveranssattTB.Text, "S") + "',"
                falt = falt + "Erordernummer,"
                varden = varden + "'" + ErOrdernrTB.Text + "',"
                falt = falt + "Markning,"
                varden = varden + "'" + MarkningTB.Text + "',"
                pprodnr = DataGridView1.Rows(i).Cells(0).Value
                dummy = hamptaprodukt("LikaNr")
                falt = falt + "inpris,"
                varden = varden + "" + kommatillpunkt(CStr(nullhantering(pinpris, "T"))) + ","
                falt = falt + "pgrupp,"
                varden = varden + "'" + CStr(nullhantering((pproduktgrupp), "S")) + "',"
                falt = falt + "caenhetspris,"
                varden = varden + "" + kommatillpunkt(CStr(putpris1)) + ","
                falt = falt + "Enhet,"
                varden = varden + "'" + nullhantering(DataGridView1.Rows(i).Cells(2).Value, "S") + "',"
                falt = falt + "rabatt,"
                varden = varden + "" + kommatillpunkt(CStr(nullhantering(ExrabL.Text, "T"))) + ","
                falt = falt + "Extrarab,"
                varden = varden + "" + kommatillpunkt(CStr(nullhantering((exrabTB.Text), "T"))) + ","
                falt = falt + "Levvilkor,"
                varden = varden + "'" + nullhantering(LevVilkorTB.Text, "S") + "',"
                falt = falt + "Ordertyp,"
                varden = varden + "'" + "Standard" + "',"
                falt = falt + "Momstyp,"
                varden = varden + "" + CStr(MomsTypTB.Text) + ","
                falt = falt + "Levdat,"
                varden = varden + "'" + nullhantering(LevdatumTB.Text, "S") + "',"
                falt = falt + "Radsumma,"
                varden = varden + "" + kommatillpunkt(CStr(nullhantering(DataGridView1.Rows(i).Cells(6).Value, "T"))) + ","
                falt = falt + "Radkostnad,"
                radkost = pinpris * nullhantering(DataGridView1.Rows(i).Cells(3).Value, "T") * nullhantering(DataGridView1.Rows(i).Cells(4).Value, "T")
                varden = varden + "" + kommatillpunkt(CStr(radkost)) + ","
                falt = falt + "Rest"
                varden = varden + "" + kommatillpunkt(CStr(nullhantering(0, "T"))) + ""
slut:
                mySQL = mySQL & "(" & falt & ") VALUES (" & varden & ");"
                myCmd = New OdbcCommand(mySQL, cn)
                myCmd.ExecuteNonQuery()

slut2:

            Next
        End If

    End Sub
    Private Sub nolla()
        kundnrTB.Text = ""
        KundnamnTB.Text = ""
        Adress1TB.Text = ""
        Adress2TB.Text = ""
        PostnrTB.Text = ""
        OrtTB.Text = ""
        LandTB.Text = ""
        KreditDagarTB.Text = 0
        KontaktmanTB.Text = ""
        exrabTB.Text = 0
        MarkningTB.Text = ""
        LevVilkorTB.Text = ""
        LeveranssattTB.Text = ""
        LevdatumTB.Text = today
        EpostTB.Text = ""
        OrdsumL.Text = ""
        momsbelL.Text = ""
        ExrabL.Text = ""
        attbetL.Text = ""
        makegrid("SimpleKredit")
        OrdnrTB.Text = getnummer("Next", "Nästa ordernummer")
        FoljnrTB.Text = ""
        OrdsumL.Text = ""
        momsbelL.Text = ""
        ExrabL.Text = ""
        attbetL.Text = ""
        kkundnr = ""
        kkundnamn = ""
        kadress1 = ""
        kadress2 = ""
        kpostnr = ""
        krOrt = ""
        LandTB.Text = ""
        ktelefon = ""
        ktelefax = ""
        kepost = ""
        kregnr = ""
        klev1 = ""
        klev2 = ""
        klev3 = ""
        kkontaktman = ""

        kgrundnom = 0
        krabattmall = ""
        krabatt = 0
        kkreditdagar = 0
        kleveranssatt = ""
        ksaljare = ""
        kkundgrupp = ""
        kkundtyp = ""
        kdistrikt = ""
        kvaluta = ""
        kkreditstatus = ""
        kkreditgrans = 0
        kfakturaperiod = ""
        klosen = ""
        OrdertypTB.Text = "Standard"
        VarrefTB.Text = defaultref
        MomsTypTB.Text = stdmoms

    End Sub

    Private Sub FakturaRB_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FakturaRB.CheckedChanged
        If FakturaRB.Checked = True Then
            BetSattGB.Visible = False
        End If
    End Sub

   
    Private Sub RakningRB_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RakningRB.CheckedChanged
        If RakningRB.Checked = True Then
            BetSattGB.Visible = True
        End If
    End Sub
    Public Sub getkund()
        Dim dummy As String
        kkundnr = kundnrTB.Text
        dummy = hamptakund("LikaNr")
        skrivakundtillform()
    End Sub
    Public Sub getprodukt()
        Dim utpris As Double
        pprodnr = DataGridView1.Item(0, rr).Value
        UTPRIS = getkundpris(krabattmall, kkundnr, pprodnr)
        DataGridView1.Item(0, rr).Value = pprodnr
        DataGridView1.Item(1, rr).Value = pprodnamn
        DataGridView1.Item(2, rr).Value = pEnhet
        DataGridView1.Item(3, rr).Value = pprisfaktor
        DataGridView1.Item(4, rr).Value = 1
        DataGridView1.Item(5, rr).Value = UTPRIS
        DataGridView1.Item(6, rr).Value = (DataGridView1.Item(3, rr).Value * DataGridView1.Item(5, rr).Value * DataGridView1.Item(4, rr).Value)
        DataGridView1.CurrentCell = DataGridView1.Item(4, rr)
    End Sub

   
    Private Sub NyB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NyB.Click
        nolla()
    End Sub

    Function hamptavarref(ByVal kod As String)
        Dim cn As OdbcConnection, mySQL As String
        Dim connStr As String
        connStr = "DSN=" + odbcsourcer + "; Database=" + databasnamn + ";Uid=v2000;Pwd=" + odbclosen
        cn = New OdbcConnection(connStr)
        cn.Open()
        mySQL = "SELECT * FROM varref"
        mySQL = mySQL + " WHERE ClientID = '" + KlientID
        mySQL = mySQL + "' And refkod = '" + kod + "'  "
        hamptavarref = "Finns ej"
        Dim myCmd As New OdbcCommand(mySQL, cn)
        Dim tabel As OdbcDataReader = myCmd.ExecuteReader(CommandBehavior.CloseConnection)
        'myCmd.ExecuteNonQuery()
        If tabel.HasRows = False Then
            GoTo slut
        Else
            tabel.Read()
            '   Stop
            hamptavarref = nullhantering(tabel("varref"), "S")
        End If
slut:
        cn.Close()
    End Function

    Private Sub VarrefTB_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles VarrefTB.KeyUp
        If e.KeyCode = Keys.Enter Then

            VarrefTB.Text = hamptavarref(VarrefTB.Text)
            If VarrefTB.Text <> "Finns ej" Then KundnamnTB.Select()
        End If
    End Sub
    Function hamptalevsatt(ByVal kod As String)
        Dim cn As OdbcConnection, mySQL As String
        Dim connStr As String
        connStr = "DSN=" + odbcsourcer + "; Database=" + databasnamn + ";Uid=v2000;Pwd=" + odbclosen
        cn = New OdbcConnection(connStr)
        cn.Open()
        mySQL = "SELECT * FROM levsatt"
        mySQL = mySQL + " WHERE ClientID = '" + KlientID
        mySQL = mySQL + "' And levsatttext > '" + kod + "'  "
        mySQL = mySQL + " order by levsatttext"
        hamptalevsatt = " "
        Dim myCmd As New OdbcCommand(mySQL, cn)
        Dim tabel As OdbcDataReader = myCmd.ExecuteReader(CommandBehavior.CloseConnection)
        'myCmd.ExecuteNonQuery()
        If tabel.HasRows = False Then
            GoTo slut
        Else
            tabel.Read()
            hamptalevsatt = nullhantering(tabel("levsatttext"), "S")
        End If
slut:
        cn.Close()
    End Function

    Private Sub LeveranssattTB_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LeveranssattTB.Click
        LeveranssattTB.Text = hamptalevsatt(LeveranssattTB.Text)
    End Sub

    Function hamptalevsvilkor(ByVal kod As String)
        Dim cn As OdbcConnection, mySQL As String
        Dim connStr As String
        connStr = "DSN=" + odbcsourcer + "; Database=" + databasnamn + ";Uid=v2000;Pwd=" + odbclosen
        cn = New OdbcConnection(connStr)
        cn.Open()
        mySQL = "SELECT * FROM levvilkor"
        mySQL = mySQL + " WHERE ClientID = '" + KlientID
        mySQL = mySQL + "' And levvilktext > '" + kod + "'  "
        mySQL = mySQL + " order by levvilktext"
        hamptalevsvilkor = " "
        Dim myCmd As New OdbcCommand(mySQL, cn)
        Dim tabel As OdbcDataReader = myCmd.ExecuteReader(CommandBehavior.CloseConnection)
        'myCmd.ExecuteNonQuery()
        If tabel.HasRows = False Then
            GoTo slut
        Else
            tabel.Read()
            hamptalevsvilkor = nullhantering(tabel("levvilktext"), "S")
        End If
slut:
        cn.Close()
    End Function

    Private Sub LevVilkorTB_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LevVilkorTB.Click
        LevVilkorTB.Text = hamptalevsvilkor(LevVilkorTB.Text)
    End Sub
    Private Sub sparafoljesedel()

        Dim cn As OdbcConnection, mySQL As String
        Dim connStr As String, falt As String, varden As String
        connStr = "DSN=" + odbcsourcer + "; Database=" + databasnamn + ";Uid=v2000;Pwd=" + odbclosen
        cn = New OdbcConnection(connStr)
        cn.Open()
        mySQL = "DELETE FROM foljesedlar "
        mySQL = mySQL + " Where foljesedelnummer = '" + nullhantering((FoljnrTB.Text), "S") + "' "
        mySQL = mySQL + " And ClientID = '" + KlientID + "' "
        Dim myCmd As New OdbcCommand(mySQL, cn)
        myCmd.ExecuteNonQuery()

        For i As Integer = 0 To rr
            nollaprodukt()
            pprodnr = nullhantering(DataGridView1.Rows(i).Cells(0).Value, "S")
            pprodnamn = DataGridView1.Rows(i).Cells(1).Value
            pEnhet = DataGridView1.Rows(i).Cells(2).Value
            pprisfaktor = DataGridView1.Rows(i).Cells(3).Value
            pantal = DataGridView1.Rows(i).Cells(4).Value
            putpris = DataGridView1.Rows(i).Cells(5).Value
            psumma = DataGridView1.Rows(i).Cells(6).Value
            falt = "" : varden = ""
            mySQL = "INSERT INTO foljesedlar "
            falt = falt + "clientid,"
            varden = varden + "'" + KlientID + "',"
            falt = falt + "MyTimeStamp,"
            varden = varden + "'" + gettimestamp() + "',"
            falt = falt + "foljesedelnummer,"
            varden = varden + "'" + nullhantering((CStr(FoljnrTB.Text)), "S") + "',"
            falt = falt + "ordernummer,"
            varden = varden + "'" + nullhantering((OrdnrTB.Text), "S") + "',"
            falt = falt + "foljesedeldatum,"
            varden = varden + "'" + today + "',"
            falt = falt + "foljesedeltid,"
            varden = varden + "'" + Mid$(Now, 12, 5) + "',"
            falt = falt + "Erreferens,"
            varden = varden + "'" + KontaktmanTB.Text + "',"
            falt = falt + "Kundnamn,"
            varden = varden + "'" + kkundnamn + "',"
            falt = falt + "Kundnummer,"
            varden = varden + "'" + kkundnr + "',"
            falt = falt + "Adress1,"
            varden = varden + "'" + kadress1 + "',"
            falt = falt + "Adress2,"
            varden = varden + "'" + kadress2 + "',"
            falt = falt + "Postnummer,"
            varden = varden + "'" + kpostnr + "',"
            falt = falt + "Ort,"
            varden = varden + "'" + kort + "',"
            falt = falt + "Land,"
            varden = varden + "'" + kland + "',"
            falt = falt + "Kreditdagar,"
            varden = varden + "" + KreditDagarTB.Text + ","
            falt = falt + "Telefon,"
            varden = varden + "'" + ktelefon + "',"
            falt = falt + "Telefax,"
            varden = varden + "'" + ktelefax + "',"
            falt = falt + "epost,"
            varden = varden + "'" + kepost + "',"
            falt = falt + "Valuta,"
            varden = varden + "'" + "SEK" + "',"
            falt = falt + "Produktnummer,"
            varden = varden + "'" + nullhantering(pprodnr, "S") + "',"
            falt = falt + "Produktnamn,"
            varden = varden + "'" + nullhantering(pprodnamn, "S") + "',"
            falt = falt + "antal,"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(pantal, "T"))) + ","
            falt = falt + "apris,"
            If MomsTypTB.Text = 4 Then
                putpris = putpris - (putpris * momsv(4))
                psumma = psumma - (psumma * momsv(4))
            End If
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(putpris, "T"))) + ","
            falt = falt + "prisfaktor,"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(pprisfaktor, "T"))) + ","
            falt = falt + "radsumma,"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(psumma, "T"))) + ","
            If Len(pprodnr) > 0 Then
                hamptaprodukt("LikaNr")
            End If
            falt = falt + "Antiforpackning,"
            varden = varden + "'" + nullhantering(pantiforpack, "S") + "',"
            If Len(pkonto) < 4 Then pkonto = stdforsaljning
            falt = falt + "Kontonummer,"
            varden = varden + "'" + nullhantering(pkonto, "S") + "',"
            falt = falt + "Varreferens,"
            varden = varden + "'" + VarrefTB.Text + "',"
            falt = falt + "Leveranssatt,"
            varden = varden + "'" + nullhantering(LeveranssattTB.Text, "S") + "',"
            falt = falt + "Erordernummer,"
            varden = varden + "'" + ErOrdernrTB.Text + "',"
            falt = falt + "Markning,"
            varden = varden + "'" + MarkningTB.Text + "',"
            falt = falt + "inpris,"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(pinpris, "T"))) + ","
            falt = falt + "inprisradsumma,"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(pinpris * pantal * pprisfaktor, "T"))) + ","
            falt = falt + "pgrupp,"
            varden = varden + "'" + CStr(nullhantering(pproduktgrupp, "S")) + "',"
            falt = falt + "caenhetspris,"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(putpris1, "T"))) + ","
            falt = falt + "Enhet,"
            varden = varden + "'" + nullhantering(pEnhet, "S") + "',"
            falt = falt + "rabatt,"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(prabatt, "T"))) + ","
            falt = falt + "Extrarab,"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering((exrabTB.Text), "T"))) + ","
            falt = falt + "Levvilkor,"
            varden = varden + "'" + nullhantering(LevVilkorTB.Text, "S") + "',"
            falt = falt + "Ordning,"
            varden = varden + "" + CStr(i) + ","
            falt = falt + "Ordertyp,"
            varden = varden + "'" + OrdertypTB.Text + "',"
            falt = falt + "Momstyp,"
            varden = varden + "" + CStr(MomsTypTB.Text) + ","
            falt = falt + "Levdat,"
            varden = varden + "'" + nullhantering(LevdatumTB.Text, "S") + "',"
            falt = falt + "Anteckning,"
            varden = varden + "'" + nullhantering("", "S") + "',"
            falt = falt + "Status,"
            varden = varden + "'" + "O" + "',"
            falt = falt + "Rest"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(prest, "T"))) + ""
            mySQL = mySQL & "(" & falt & ") VALUES (" & varden & ");"
            myCmd = New OdbcCommand(mySQL, cn)
            myCmd.ExecuteNonQuery()
        Next i
        'Stop
        cn.Close()
    End Sub
    Private Sub sparaorder()
        Dim cn As OdbcConnection, mySQL As String
        Dim connStr As String, falt As String, varden As String
        connStr = "DSN=" + odbcsourcer + "; Database=" + databasnamn + ";Uid=v2000;Pwd=" + odbclosen
        cn = New OdbcConnection(connStr)
        cn.Open()
        mySQL = "DELETE FROM ordrar "
        mySQL = mySQL + " Where ordernummer = '" + nullhantering((OrdnrTB.Text), "S") + "' "
        mySQL = mySQL + " And ClientID = '" + KlientID + "' "
        Dim myCmd As New OdbcCommand(mySQL, cn)
        myCmd.ExecuteNonQuery()

        For i As Integer = 0 To rr
            nollaprodukt()
            pprodnr = nullhantering(DataGridView1.Rows(i).Cells(0).Value, "S")
            pprodnamn = DataGridView1.Rows(i).Cells(1).Value
            pEnhet = DataGridView1.Rows(i).Cells(2).Value
            pprisfaktor = DataGridView1.Rows(i).Cells(3).Value
            pantal = DataGridView1.Rows(i).Cells(4).Value
            putpris = DataGridView1.Rows(i).Cells(5).Value
            psumma = DataGridView1.Rows(i).Cells(6).Value
            falt = "" : varden = ""
            mySQL = "INSERT INTO Ordrar "
            falt = falt + "clientid,"
            varden = varden + "'" + KlientID + "',"
            falt = falt + "MyTimeStamp,"
            varden = varden + "'" + gettimestamp() + "',"
            falt = falt + "ordernummer,"
            varden = varden + "'" + nullhantering((OrdnrTB.Text), "S") + "',"
            falt = falt + "Orderdatum,"
            varden = varden + "'" + today + "',"
            falt = falt + "Ordertid,"
            varden = varden + "'" + Mid$(Now, 12, 5) + "',"
            falt = falt + "Erreferens,"
            varden = varden + "'" + KontaktmanTB.Text + "',"
            falt = falt + "Kundnamn,"
            varden = varden + "'" + KundnamnTB.Text + "',"
            falt = falt + "Kundnummer,"
            varden = varden + "'" + kundnrTB.Text + "',"
            falt = falt + "Adress1,"
            varden = varden + "'" + Adress1TB.Text + "',"
            falt = falt + "Adress2,"
            varden = varden + "'" + Adress2TB.Text + "',"
            falt = falt + "Postnummer,"
            varden = varden + "'" + PostnrTB.Text + "',"
            falt = falt + "Ort,"
            varden = varden + "'" + OrtTB.Text + "',"
            falt = falt + "Land,"
            varden = varden + "'" + LandTB.Text + "',"
            falt = falt + "Kreditdagar,"
            varden = varden + "" + KreditDagarTB.Text + ","
            falt = falt + "Telefon,"
            varden = varden + "'" + ktelefon + "',"
            falt = falt + "Telefax,"
            varden = varden + "'" + ktelefax + "',"
            falt = falt + "epost,"
            varden = varden + "'" + kepost + "',"
            falt = falt + "Valuta,"
            varden = varden + "'" + "SEK" + "',"
            falt = falt + "Produktnummer,"
            varden = varden + "'" + nullhantering(pprodnr, "S") + "',"
            falt = falt + "Produktnamn,"
            varden = varden + "'" + nullhantering(pprodnamn, "S") + "',"
            falt = falt + "antal,"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(pantal, "T"))) + ","
            falt = falt + "apris,"
            If MomsTypTB.Text = 4 Then
                putpris = putpris - (putpris * momsv(4))
                psumma = psumma - (psumma * momsv(4))
            End If
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(putpris, "T"))) + ","
            falt = falt + "prisfaktor,"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(pprisfaktor, "T"))) + ","
            falt = falt + "radsumma,"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(psumma, "T"))) + ","
            If Len(pprodnr) > 0 Then
                hamptaprodukt("LikaNr")
            End If
            falt = falt + "Antiforpackning,"
            varden = varden + "'" + nullhantering(pantiforpack, "S") + "',"
            If Len(pkonto) < 4 Then pkonto = stdforsaljning
            falt = falt + "Kontonummer,"
            varden = varden + "'" + nullhantering(pkonto, "S") + "',"
            falt = falt + "Varreferens,"
            varden = varden + "'" + VarrefTB.Text + "',"
            falt = falt + "Leveranssatt,"
            varden = varden + "'" + nullhantering(LeveranssattTB.Text, "S") + "',"
            falt = falt + "Erordernummer,"
            varden = varden + "'" + ErOrdernrTB.Text + "',"
            falt = falt + "Markning,"
            varden = varden + "'" + MarkningTB.Text + "',"
            falt = falt + "inpris,"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(pinpris, "T"))) + ","
            falt = falt + "inprisradsumma,"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(pinpris * pantal * pprisfaktor, "T"))) + ","
            falt = falt + "pgrupp,"
            varden = varden + "'" + CStr(nullhantering(pproduktgrupp, "S")) + "',"
            falt = falt + "caenhetspris,"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(putpris1, "T"))) + ","
            falt = falt + "Enhet,"
            varden = varden + "'" + nullhantering(pEnhet, "S") + "',"
            falt = falt + "rabatt,"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(prabatt, "T"))) + ","
            falt = falt + "Extrarab,"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering((exrabTB.Text), "T"))) + ","
            falt = falt + "Levvilkor,"
            varden = varden + "'" + nullhantering(LevVilkorTB.Text, "S") + "',"
            falt = falt + "Ordning,"
            varden = varden + "" + CStr(i) + ","
            falt = falt + "Ordertyp,"
            varden = varden + "'" + OrdertypTB.Text + "',"
            falt = falt + "Momstyp,"
            varden = varden + "" + CStr(MomsTypTB.Text) + ","
            falt = falt + "Levdat,"
            varden = varden + "'" + nullhantering(LevdatumTB.Text, "S") + "',"
            falt = falt + "Anteckning,"
            varden = varden + "'" + nullhantering("", "S") + "',"
            falt = falt + "Status,"
            varden = varden + "'" + "O" + "',"
            falt = falt + "Rest"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(prest, "T"))) + ""
            mySQL = mySQL & "(" & falt & ") VALUES (" & varden & ");"

            myCmd = New OdbcCommand(mySQL, cn)
            myCmd.ExecuteNonQuery()
        Next i
        cn.Close()
    End Sub
    Private Sub nollaprodukt()
        pprodnr = ""
        pprodnamn = ""
        pinpris = 0
        pgrundpris = 0
        pgrundrab = 0
        pprisfaktor = 0
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
        pantiforpack = "" '
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
    End Sub

    Private Sub sparatillarkiv()
       Dim cn As OdbcConnection, mySQL As String
        Dim connStr As String, falt As String, varden As String
        connStr = "DSN=" + odbcsourcer + "; Database=" + databasnamn + ";Uid=v2000;Pwd=" + odbclosen
        cn = New OdbcConnection(connStr)
        cn.Open()
        For i As Integer = 0 To rr
            nollaprodukt()
            pprodnr = nullhantering(DataGridView1.Rows(i).Cells(0).Value, "S")
            pprodnamn = DataGridView1.Rows(i).Cells(1).Value
            pEnhet = DataGridView1.Rows(i).Cells(2).Value
            pprisfaktor = DataGridView1.Rows(i).Cells(3).Value
            pantal = DataGridView1.Rows(i).Cells(4).Value
            putpris = DataGridView1.Rows(i).Cells(5).Value
            psumma = DataGridView1.Rows(i).Cells(6).Value
            falt = "" : varden = ""
            mySQL = "INSERT INTO Orderarkiv "
            falt = falt + "clientid,"
            varden = varden + "'" + KlientID + "',"
            falt = falt + "MyTimeStamp,"
            varden = varden + "'" + gettimestamp() + "',"
            falt = falt + "Debittyp,"
            varden = varden + "'" + "Faktura" + "',"
            falt = falt + "DebitID,"
            varden = varden + "'" + "12345" + "',"
            falt = falt + "ordernummer,"
            varden = varden + "'" + nullhantering((OrdnrTB.Text), "S") + "',"
            falt = falt + "Arkivdatum,"
            varden = varden + "'" + today + "',"
            falt = falt + "Arkivid,"
            varden = varden + "'" + Mid$(Now, 12, 5) + "',"
            falt = falt + "Erreferens,"
            varden = varden + "'" + KontaktmanTB.Text + "',"
            falt = falt + "Kundnamn,"
            varden = varden + "'" + kkundnamn + "',"
            falt = falt + "Kundnummer,"
            varden = varden + "'" + kkundnr + "',"
            falt = falt + "Adress1,"
            varden = varden + "'" + kadress1 + "',"
            falt = falt + "Adress2,"
            varden = varden + "'" + kadress2 + "',"
            falt = falt + "Postnummer,"
            varden = varden + "'" + kpostnr + "',"
            falt = falt + "Ort,"
            varden = varden + "'" + kort + "',"
            falt = falt + "Land,"
            varden = varden + "'" + kland + "',"
            falt = falt + "Kreditdagar,"
            varden = varden + "" + KreditDagarTB.Text + ","
            falt = falt + "Telefon,"
            varden = varden + "'" + ktelefon + "',"
            falt = falt + "Telefax,"
            varden = varden + "'" + ktelefax + "',"
            falt = falt + "epost,"
            varden = varden + "'" + kepost + "',"
            falt = falt + "Valuta,"
            varden = varden + "'" + "SEK" + "',"
            falt = falt + "Produktnummer,"
            varden = varden + "'" + nullhantering(pprodnr, "S") + "',"
            falt = falt + "Produktnamn,"
            varden = varden + "'" + nullhantering(pprodnamn, "S") + "',"
            falt = falt + "antal,"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(pantal, "T"))) + ","
            falt = falt + "apris,"
            If MomsTypTB.Text = 4 Then
                putpris = putpris - (putpris * momsv(4))
                psumma = psumma - (psumma * momsv(4))
            End If
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(putpris, "T"))) + ","
            falt = falt + "prisfaktor,"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(pprisfaktor, "T"))) + ","
            falt = falt + "radsumma,"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(psumma, "T"))) + ","
            If Len(pprodnr) > 0 Then
                hamptaprodukt("LikaNr")
            End If
            falt = falt + "Antiforpackning,"
            varden = varden + "'" + nullhantering(pantiforpack, "S") + "',"
            If Len(pkonto) < 4 Then pkonto = stdforsaljning
            falt = falt + "Kontonummer,"
            varden = varden + "'" + nullhantering(pkonto, "S") + "',"
            falt = falt + "Varreferens,"
            varden = varden + "'" + VarrefTB.Text + "',"
            falt = falt + "Leveranssatt,"
            varden = varden + "'" + nullhantering(LeveranssattTB.Text, "S") + "',"
            falt = falt + "Erordernummer,"
            varden = varden + "'" + ErOrdernrTB.Text + "',"
            falt = falt + "Markning,"
            varden = varden + "'" + MarkningTB.Text + "',"
            falt = falt + "inpris,"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(pinpris, "T"))) + ","
            falt = falt + "inprisradsumma,"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(pinpris * pantal * pprisfaktor, "T"))) + ","
            falt = falt + "pgrupp,"
            varden = varden + "'" + CStr(nullhantering(pproduktgrupp, "S")) + "',"
            falt = falt + "caenhetspris,"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(putpris1, "T"))) + ","
            falt = falt + "Enhet,"
            varden = varden + "'" + nullhantering(pEnhet, "S") + "',"
            falt = falt + "rabatt,"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(prabatt, "T"))) + ","
            falt = falt + "Extrarab,"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering((exrabTB.Text), "T"))) + ","
            falt = falt + "Levvilkor,"
            varden = varden + "'" + nullhantering(LevVilkorTB.Text, "S") + "',"
            falt = falt + "Ordning,"
            varden = varden + "" + CStr(i) + ","
            falt = falt + "Ordertyp,"
            varden = varden + "'" + OrdertypTB.Text + "',"
            falt = falt + "Momstyp,"
            varden = varden + "" + CStr(MomsTypTB.Text) + ","
            falt = falt + "Levdat,"
            varden = varden + "'" + nullhantering(LevdatumTB.Text, "S") + "',"
            falt = falt + "Anteckning,"
            varden = varden + "'" + nullhantering("", "S") + "',"
            falt = falt + "Bokfort,"
            varden = varden + "'" + nullhantering("", "S") + "',"
            falt = falt + "Status,"
            varden = varden + "'" + "O" + "',"
            falt = falt + "Rest"
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(prest, "T"))) + ""
            mySQL = mySQL & "(" & falt & ") VALUES (" & varden & ");"
            Dim myCmd As New OdbcCommand(mySQL, cn)
            myCmd = New OdbcCommand(mySQL, cn)
            myCmd.ExecuteNonQuery()
        Next i
        cn.Close()
    End Sub


    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Spara_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Spara.Click
        ProgressBar1.Value = 10
        sparaorder()
        ProgressBar1.Value = 100
    End Sub

    Function getorder(ByVal ordnr As String)
        'On Error Resume Next
        Dim n As Integer
        Dim cn As OdbcConnection, mySQL As String
        Dim connStr As String, utpris(10) As Double
        Dim myCmd As OdbcCommand
        Dim tabel As OdbcDataReader


        connStr = "DSN=" + odbcsourcer + "; Database=" + databasnamn + ";Uid=v2000;Pwd=" + odbclosen
        cn = New OdbcConnection(connStr)
        cn.Open()
        mySQL = "SELECT * FROM ordrar "
        mySQL = mySQL + "Where ordernummer = " & "'" + ordnr + "'"
        mySQL = mySQL + " And clientid = '" + KlientID + "' "
        mySQL = mySQL + " order by mytimestamp "
        myCmd = New OdbcCommand(mySQL, cn)
        tabel = myCmd.ExecuteReader(CommandBehavior.CloseConnection)
        n = 1
        getorder = "Yes"

        While tabel.Read()

            OrdnrTB.Text = nullhantering(tabel("ordernummer"), "S")

            kundnrTB.Text = nullhantering(tabel("Kundnummer"), "S") : kkundnr = kundnrTB.Text
            KundnamnTB.Text = nullhantering(tabel("Kundnamn"), "S") : kkundnamn = KundnamnTB.Text
            Adress1TB.Text = nullhantering(tabel("adress1"), "S") : kadress1 = Adress1TB.Text
            Adress2TB.Text = nullhantering(tabel("adress2"), "S") : kadress2 = Adress2TB.Text
            PostnrTB.Text = nullhantering(tabel("Postnummer"), "S") : kpostnr = PostnrTB.Text
            OrtTB.Text = nullhantering(tabel("ort"), "S") : kort = OrtTB.Text
            LandTB.Text = nullhantering(tabel("land"), "S") : kland = LandTB.Text
            MomsTypTB.Text = nullhantering(tabel("momstyp"), "T")
            KontaktmanTB.Text = nullhantering(tabel("erreferens"), "S") : kkontaktman = KontaktmanTB.Text
            EpostTB.Text = nullhantering(tabel("epost"), "S") : kepost = EpostTB.Text
            ktelefon = nullhantering(tabel("telefon"), "S")
            OrdertypTB.Text = nullhantering(tabel("ordertyp"), "S")
            ErOrdernrTB.Text = nullhantering(tabel("erordernummer"), "S")
            MarkningTB.Text = nullhantering(tabel("markning"), "S")
            LevVilkorTB.Text = nullhantering(tabel("levvilkor"), "S")
            LeveranssattTB.Text = nullhantering(tabel("leveranssatt"), "S")
            LevdatumTB.Text = nullhantering(tabel("levdat"), "S")
            'EpostTB.Text = nullhantering(tabel("epost"), "S")
            VarrefTB.Text = nullhantering(tabel("varreferens"), "S")
            KreditDagarTB.Text = nullhantering(tabel("kreditdagar"), "T")
            DataGridView1.Item(0, n - 1).Value = nullhantering(tabel("Produktnummer"), "S")
            DataGridView1.Item(1, n - 1).Value = nullhantering(tabel("Produktnamn"), "S")

            DataGridView1.Item(2, n - 1).Value = nullhantering(tabel("Enhet"), "S")
            If nullhantering(tabel("Prisfaktor"), "T") = 0 Then GoTo slutloop
            DataGridView1.Item(3, n - 1).Value = nullhantering(tabel("Prisfaktor"), "T")
            DataGridView1.Item(4, n - 1).Value = nullhantering(tabel("antal"), "T")
            If MomsTypTB.Text = 4 Then
                tillfpris = nullhantering(tabel("apris"), "T")
                tillfpris = tillfpris + (tillfpris * momsv(1))
                DataGridView1.Item(5, n - 1).Value = tillfpris
                tillfpris = nullhantering(tabel("radsumma"), "T")
                tillfpris = tillfpris + (tillfpris * momsv(1))
                DataGridView1.Item(6, n - 1).Value = tillfpris
                tillfpris = nullhantering(tabel("caenhetspris"), "T")
                tillfpris = tillfpris + (tillfpris * momsv(1))
                DataGridView1.Item(7, n - 1).Value = tillfpris
            Else
                DataGridView1.Item(5, n - 1).Value = nullhantering(tabel("apris"), "T")
                DataGridView1.Item(6, n - 1).Value = nullhantering(tabel("radsumma"), "T")
                DataGridView1.Item(7, n - 1).Value = nullhantering(tabel("caenhetspris"), "T")
            End If
slutloop:
            n = n + 1
        End While
        summering()
        cn.Close()
    End Function


    Private Sub OrdnrTB_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles OrdnrTB.KeyUp
        If e.KeyCode = Keys.Enter Then
            getorder(OrdnrTB.Text)
        End If
    End Sub

    Private Sub OrdnrTB_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OrdnrTB.TextChanged
        ordernummer = OrdnrTB.Text
    End Sub

    Private Sub SammanställningarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SammanställningarToolStripMenuItem.Click
        OrderSamF.Show()
        OrderSamF.BringToFront()
    End Sub

    Function getfoljesedel(ByVal ordnr As String)
        'On Error Resume Next
        Dim n As Integer
        Dim cn As OdbcConnection, mySQL As String
        Dim connStr As String, utpris(10) As Double
        Dim myCmd As OdbcCommand
        Dim tabel As OdbcDataReader


        connStr = "DSN=" + odbcsourcer + "; Database=" + databasnamn + ";Uid=v2000;Pwd=" + odbclosen
        cn = New OdbcConnection(connStr)
        cn.Open()
        mySQL = "SELECT * FROM foljesedlar "
        mySQL = mySQL + "Where foljesedelnummer = " & "'" + ordnr + "'"
        mySQL = mySQL + " And clientid = '" + KlientID + "' "
        mySQL = mySQL + " order by mytimestamp "
        myCmd = New OdbcCommand(mySQL, cn)
        tabel = myCmd.ExecuteReader(CommandBehavior.CloseConnection)
        n = 1
        getfoljesedel = "Yes"

        While tabel.Read()
            FoljnrTB.Text = nullhantering(tabel("Foljesedelnummer"), "S")
            OrdnrTB.Text = nullhantering(tabel("ordernummer"), "S")
            kundnrTB.Text = nullhantering(tabel("Kundnummer"), "S") : kkundnr = kundnrTB.Text
            KundnamnTB.Text = nullhantering(tabel("Kundnamn"), "S") : kkundnamn = KundnamnTB.Text
            Adress1TB.Text = nullhantering(tabel("adress1"), "S") : kadress1 = Adress1TB.Text
            Adress2TB.Text = nullhantering(tabel("adress2"), "S") : kadress2 = Adress2TB.Text
            PostnrTB.Text = nullhantering(tabel("Postnummer"), "S") : kpostnr = PostnrTB.Text
            OrtTB.Text = nullhantering(tabel("ort"), "S") : kort = OrtTB.Text
            LandTB.Text = nullhantering(tabel("land"), "S") : kland = LandTB.Text
            MomsTypTB.Text = nullhantering(tabel("momstyp"), "T")
            KontaktmanTB.Text = nullhantering(tabel("erreferens"), "S") : kkontaktman = KontaktmanTB.Text
            EpostTB.Text = nullhantering(tabel("epost"), "S") : kepost = EpostTB.Text
            ktelefon = nullhantering(tabel("telefon"), "S")
            ErOrdernrTB.Text = nullhantering(tabel("erordernummer"), "S")
            MarkningTB.Text = nullhantering(tabel("markning"), "S")
            LevVilkorTB.Text = nullhantering(tabel("levvilkor"), "S")
            LeveranssattTB.Text = nullhantering(tabel("leveranssatt"), "S")
            LevdatumTB.Text = nullhantering(tabel("levdat"), "S")
            'EpostTB.Text = nullhantering(tabel("epost"), "S")
            VarrefTB.Text = nullhantering(tabel("varreferens"), "S")
            KreditDagarTB.Text = nullhantering(tabel("kreditdagar"), "T")
            DataGridView1.Item(0, n - 1).Value = nullhantering(tabel("Produktnummer"), "S")
            DataGridView1.Item(1, n - 1).Value = nullhantering(tabel("Produktnamn"), "S")

            DataGridView1.Item(2, n - 1).Value = nullhantering(tabel("Enhet"), "S")
            If nullhantering(tabel("Prisfaktor"), "T") = 0 Then GoTo slutloop
            DataGridView1.Item(3, n - 1).Value = nullhantering(tabel("Prisfaktor"), "T")
            DataGridView1.Item(4, n - 1).Value = nullhantering(tabel("antal"), "T")
            If MomsTypTB.Text = 4 Then
                tillfpris = nullhantering(tabel("apris"), "T")
                tillfpris = tillfpris + (tillfpris * momsv(1))
                DataGridView1.Item(5, n - 1).Value = tillfpris
                tillfpris = nullhantering(tabel("radsumma"), "T")
                tillfpris = tillfpris + (tillfpris * momsv(1))
                DataGridView1.Item(6, n - 1).Value = tillfpris
                tillfpris = nullhantering(tabel("caenhetspris"), "T")
                tillfpris = tillfpris + (tillfpris * momsv(1))
                DataGridView1.Item(7, n - 1).Value = tillfpris
            Else
                DataGridView1.Item(5, n - 1).Value = nullhantering(tabel("apris"), "T")
                DataGridView1.Item(6, n - 1).Value = nullhantering(tabel("radsumma"), "T")
                DataGridView1.Item(7, n - 1).Value = nullhantering(tabel("caenhetspris"), "T")
            End If
slutloop:
            n = n + 1
        End While
        summering()
        cn.Close()
    End Function

    Private Sub FoljnrTB_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles FoljnrTB.KeyUp
        If e.KeyCode = Keys.Enter Then
            getfoljesedel(FoljnrTB.Text)
        End If
    End Sub

    Private Sub FoljnrTB_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FoljnrTB.TextChanged

    End Sub
    Function suddaorder(ByVal ordnr As String)
        suddaorder = "Yes"
        Dim cn As OdbcConnection, mySQL As String
        Dim connStr As String
        connStr = "DSN=" + odbcsourcer + "; Database=" + databasnamn + ";Uid=v2000;Pwd=" + odbclosen
        cn = New OdbcConnection(connStr)
        cn.Open()
        mySQL = "DELETE FROM ordrar "
        mySQL = mySQL + " Where ordernummer = '" + nullhantering((ordnr), "S") + "' "
        mySQL = mySQL + " And ClientID = '" + KlientID + "' "
        Dim myCmd As New OdbcCommand(mySQL, cn)
        myCmd.ExecuteNonQuery()
        cn.Close()
    End Function

    Function suddafoljesedel(ByVal ordnr As String)
        suddafoljesedel = "Yes"
        Dim cn As OdbcConnection, mySQL As String
        Dim connStr As String
        connStr = "DSN=" + odbcsourcer + "; Database=" + databasnamn + ";Uid=v2000;Pwd=" + odbclosen
        cn = New OdbcConnection(connStr)
        cn.Open()
        mySQL = "DELETE FROM foljesedlar "
        mySQL = mySQL + " Where foljesedelnummer = '" + nullhantering((ordnr), "S") + "' "
        mySQL = mySQL + " And ClientID = '" + KlientID + "' "
        Dim myCmd As New OdbcCommand(mySQL, cn)
        myCmd.ExecuteNonQuery()
        cn.Close()
    End Function

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub OrdertypTB_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles OrdertypTB.Click
        If OrdertypTB.Text = "Standard" Then
            OrdertypTB.Text = "Återkommande"
        ElseIf OrdertypTB.Text = "Återkommande" Then
            OrdertypTB.Text = "Offert"
        Else
            OrdertypTB.Text = "Standard"

        End If



    End Sub
    Private Function HTMLfakturahuvud()
        htmhuv = ""
        htmhuv = htmhuv + "       <html>"
        htmhuv = htmhuv + "<head>"
        htmhuv = htmhuv + "<meta http-equiv='Content-Language' content='sv'>"
        htmhuv = htmhuv + "<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'>"
        htmhuv = htmhuv + "<title>Dokument</title>"
        htmhuv = htmhuv + "</head>"
        htmhuv = htmhuv + "<body>"
        htmhuv = htmhuv + "<div align='center'>"
        htmhuv = htmhuv + "	<div align='left'>"
        htmhuv = htmhuv + "	<table border='0' width='800'>"
        htmhuv = htmhuv + "		<tr>"
        htmhuv = htmhuv + "			<td width='505' colspan='2' bgcolor='#D8D5A9' height='31'><b><font size='4'>" + Firmanamn + "</font></b></td>"
        htmhuv = htmhuv + "			<td width='285' bgcolor='#D8D5A9' height='31'><b><font size='4'>" + utskriftsort + "</font></b></td>"
        htmhuv = htmhuv + "		</tr>"
        htmhuv = htmhuv + "		<tr>"
        htmhuv = htmhuv + "			<td width='505' colspan='2'><font size='2'>" + Postadress1 + "</font></td>"
        htmhuv = htmhuv + "			<td width='285'><font size='2'>Nummer:" + CStr(getnummer("Kolla", "Nästa fakturanummer")) + "</font></td>"
        htmhuv = htmhuv + "		</tr>"
        htmhuv = htmhuv + "		<tr>"
        htmhuv = htmhuv + "			<td width='505' colspan='2'><font size='2'>" + Postadress2 + "</font></td>"
        htmhuv = htmhuv + "			<td width='285'>&nbsp;</td>"
        htmhuv = htmhuv + "		</tr>"
        htmhuv = htmhuv + "		<tr>"
        htmhuv = htmhuv + "			<td width='505' colspan='2'><font size='2'>" + Postadress3 + "</font></td>"
        htmhuv = htmhuv + "			<td width='285'><font size='2'>Datum: " + today + "</font></td>"
        htmhuv = htmhuv + "		</tr>"
        htmhuv = htmhuv + "		<tr>"
        htmhuv = htmhuv + "			<td width='505' colspan='2'><font size='2'>" + land + "</font></td>"
        htmhuv = htmhuv + "			<td width='285'><font size='2'>Kundnr: " + kkundnr + "</font></td>"
        htmhuv = htmhuv + "		</tr>"
        htmhuv = htmhuv + "		<tr>"
        htmhuv = htmhuv + "			<td width='505' colspan='2'><font size='2'>Tel: " + telefon + "</font></td>"
        htmhuv = htmhuv + "			<td width='285'>&nbsp;</td>"
        htmhuv = htmhuv + "		</tr>"
        htmhuv = htmhuv + "		<tr>"
        htmhuv = htmhuv + "			<td width='505' colspan='2'><font size='2'>Epost: " + epost + "</font></td>"
        htmhuv = htmhuv + "			<td width='285'>&nbsp;</td>"
        htmhuv = htmhuv + "		</tr>"
        htmhuv = htmhuv + "		<tr>"
        htmhuv = htmhuv + "			<td width='505' colspan='2'>&nbsp;</td>"
        htmhuv = htmhuv + "			<td width='285'><b>" + kkundnamn + "</b></td>"
        htmhuv = htmhuv + "		</tr>"
        htmhuv = htmhuv + "		<tr>"
        htmhuv = htmhuv + "			<td width='505' colspan='2'>&nbsp;</td>"
        htmhuv = htmhuv + "			<td width='285'><font size='2'>" + kadress1 + "</font></td>"
        htmhuv = htmhuv + "		</tr>"
        htmhuv = htmhuv + "		<tr>"
        htmhuv = htmhuv + "			<td width='118'><font size='2'>Leveranssätt:</font></td>"
        htmhuv = htmhuv + "			<td width='382'><font size='2'>" + LeveranssattTB.Text + "</font></td>"
        htmhuv = htmhuv + "			<td width='285'><font size='2'>" + kadress2 + "</font></td>"
        htmhuv = htmhuv + "		</tr>"
        htmhuv = htmhuv + "		<tr>"
        htmhuv = htmhuv + "			<td width='118'><font size='2'>Leveransvilkor</font></td>"
        htmhuv = htmhuv + "			<td width='382'><font size='2'>" + LevVilkorTB.Text + "</font></td>"
        htmhuv = htmhuv + "			<td width='285'><font size='2'>" + kpostnr + " " + kort + "</font></td>"
        htmhuv = htmhuv + "		</tr>"
        htmhuv = htmhuv + "		<tr>"
        htmhuv = htmhuv + "			<td width='118'><font size='2'>Betalningsvilkor</font></td>"
        htmhuv = htmhuv + "			<td width='382'><font size='2'>" + KreditDagarTB.Text + " dagar" + "</font></td>"
        htmhuv = htmhuv + "			<td width='285'><font size='2'>" + kland + "</font></td>"
        htmhuv = htmhuv + "		</tr>"
        AddDays(3D)
        htmhuv = htmhuv + "		<tr>"
        If FakturaRB.Checked = True Or RakningRB.Checked = True Then
            htmhuv = htmhuv + "			<td width='118'><font size='2'>Förfallodatum</font></td>"
            htmhuv = htmhuv + "			<td width='382'><font size='2'>" + CStr(CDate(today).AddDays(CDbl(KreditDagarTB.Text))) + "</font></td>"
        End If
        htmhuv = htmhuv + "			<td width='285'>&nbsp;</td>"
        htmhuv = htmhuv + "		</tr>"
        htmhuv = htmhuv + "		<tr>"
        htmhuv = htmhuv + "			<td width='118'><font size='2'>Ordernummer</font></td>"
        htmhuv = htmhuv + "			<td width='382'><font size='2'>" + OrdnrTB.Text + "</font></td>"
        htmhuv = htmhuv + "			<td width='285'><font size='2'>Tel: " + ktelefon + "</font></td>"
        htmhuv = htmhuv + "		</tr>"
        htmhuv = htmhuv + "		<tr>"
        htmhuv = htmhuv + "			<td width='118'><font size='2'>Ertordernummer:</font></td>"
        htmhuv = htmhuv + "			<td width='382'><font size='2'>" + ErOrdernrTB.Text + "</font></td>"
        htmhuv = htmhuv + "			<td width='285'><font size='2'>Epost: " + kepost + "</font></td>"
        htmhuv = htmhuv + "		</tr>"
        htmhuv = htmhuv + "		<tr>"
        htmhuv = htmhuv + "			<td width='118'><font size='2'>Leveransdatum</font></td>"
        htmhuv = htmhuv + "			<td width='382'><font size='2'>" + LevdatumTB.Text + "</font></td>"
        htmhuv = htmhuv + "			<td width='285'>&nbsp;</td>"
        htmhuv = htmhuv + "		</tr>"
        htmhuv = htmhuv + "		<tr>"
        htmhuv = htmhuv + "			<td width='118'><font size='2'>Märkning</font></td>"
        htmhuv = htmhuv + "			<td width='382'><font size='2'>" + MarkningTB.Text + "</font></td>"
        htmhuv = htmhuv + "			<td width='285'>&nbsp;</td>"
        htmhuv = htmhuv + "		</tr>"
        htmhuv = htmhuv + "	</table>"
        htmhuv = htmhuv + "	</div>"
        HTMLfakturahuvud = htmhuv
    End Function

    Private Sub EpostB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EpostB.Click
        Dim epostbody As String

        utskriftsort = ""
        If FakturaRB.Checked = True Then
            utskriftsort = "Faktura"
        ElseIf FoljesedelRB.Checked = True Then
            utskriftsort = "Följesedel"
        ElseIf OrderbekraftRB.Checked = True Then
            utskriftsort = "Orderbekräftelse"
        Else
            utskriftsort = "Offert"
        End If

        epostbody = ""
        epostbody = HTMLfakturahuvud() + HTMLrader() + htmlfot()

        skickatillepostrelay(epost, (EpostTB.Text), utskriftsort, epostbody, kkundnamn)
        skickatillepostrelay(epost, (epost), utskriftsort, epostbody, kkundnamn)

        MsgBox("Epost skickat!!")




    End Sub
    Function HTMLrader()
        htmhuv = ""
        htmhuv = htmhuv + "       	<div align='center'>"
        htmhuv = htmhuv + "	<div align='left'>"
        htmhuv = htmhuv + "	<table border='0' width='800'>"
        htmhuv = htmhuv + "		<tr>"
        htmhuv = htmhuv + "		<td width='120' bgcolor='#D8D5A9'><b><font size='2'>Produktnr</font></b></td>"
        htmhuv = htmhuv + "		<td width='314' bgcolor='#D8D5A9'><b><font size='2'>Beskrivning</font></b></td>"
        htmhuv = htmhuv + "		<td align='center' bgcolor='#D8D5A9'><b><font size='2'>Enhet</font></b></td>"
        htmhuv = htmhuv + "		<td align='right' width='64' bgcolor='#D8D5A9'><b><font size='2'>Antal</font></b></td>"
        htmhuv = htmhuv + "		<td align='right' width='94' bgcolor='#D8D5A9'><b><font size='2'>apris</font></b></td>"
        htmhuv = htmhuv + "		<td align='right' width='101' bgcolor='#D8D5A9'><b><font size='2'>Summa</font></b></td>"
        htmhuv = htmhuv + "	</tr>"
        For i As Integer = 0 To rr
            pprodnr = DataGridView1.Rows(i).Cells(0).Value
            pprodnamn = DataGridView1.Rows(i).Cells(1).Value
            pEnhet = DataGridView1.Rows(i).Cells(2).Value
            prantal = DataGridView1.Rows(i).Cells(4).Value
            putpris = DataGridView1.Rows(i).Cells(5).Value
            psumma = DataGridView1.Rows(i).Cells(6).Value
            htmhuv = htmhuv + "	<tr>"
            htmhuv = htmhuv + "		<td width='120'><font size='2'>" + pprodnr + "</font></td>"
            htmhuv = htmhuv + "		<td width='314'><font size='2'>" + pprodnamn + "</font></td>"
            htmhuv = htmhuv + "		<td align='center'><font size='2'>" + pEnhet + "</font></td>"
            If prantal <> 0 Then
                htmhuv = htmhuv + "		<td align='right' width='64'><font size='2'>" + CStr(prantal) + "</font></td>"
                htmhuv = htmhuv + "		<td align='right' width='94'><font size='2'>" + CStr(putpris) + "</font></td>"
                htmhuv = htmhuv + "		<td align='right' width='101'><font size='2'>" + CStr(psumma) + "</font></td>"
            End If
            htmhuv = htmhuv + "	</tr>"
        Next
        htmhuv = htmhuv + "	<tr>"
        htmhuv = htmhuv + "		<td width='120'>&nbsp;</td>"
        htmhuv = htmhuv + "		<td width='314'>&nbsp;</td>"
        htmhuv = htmhuv + "		<td align='center'>&nbsp;</td>"
        htmhuv = htmhuv + "		<td align='right' width='64'>&nbsp;</td>"
        htmhuv = htmhuv + "		<td align='right' width='94'>&nbsp;</td>"
        htmhuv = htmhuv + "		<td align='right' width='101'>&nbsp;</td>"
        htmhuv = htmhuv + "	</tr>"
        If FakturaRB.Checked = True Or OrderbekraftRB.Checked = True Or RakningRB.Checked Then
            htmhuv = htmhuv + "	<tr>"
            htmhuv = htmhuv + "		<td width='120'>&nbsp;</td>"
            htmhuv = htmhuv + "		<td width='314'>&nbsp;</td>"
            htmhuv = htmhuv + "		<td align='center'>&nbsp;</td>"
            htmhuv = htmhuv + "		<td align='right' width='64'>&nbsp;</td>"
            htmhuv = htmhuv + "		<td align='right' width='94'><font size='2'>Summa</font></td>"
            htmhuv = htmhuv + "		<td align='right' width='101'><font size='2'>" + CStr(OrdsumL.Text) + "</font></td>"
            htmhuv = htmhuv + "	</tr>"
            htmhuv = htmhuv + "	<tr>"
            htmhuv = htmhuv + "		<td width='120'>&nbsp;</td>"
            htmhuv = htmhuv + "		<td width='314'>&nbsp;</td>"
            htmhuv = htmhuv + "		<td align='center'>&nbsp;</td>"
            htmhuv = htmhuv + "		<td align='right' width='64'>&nbsp;</td>"
            htmhuv = htmhuv + "		<td align='right' width='94'><font size='2'>Moms 25%</font></td>"
            htmhuv = htmhuv + "		<td align='right' width='101'><font size='2'>" + CStr(momsbelL.Text) + "</font></td>"
            htmhuv = htmhuv + "	</tr>"
            htmhuv = htmhuv + "	<tr>"
            htmhuv = htmhuv + "		<td width='120'>&nbsp;</td>"
            htmhuv = htmhuv + "		<td width='314'>&nbsp;</td>"
            htmhuv = htmhuv + "		<td align='center'>&nbsp;</td>"
            htmhuv = htmhuv + "		<td align='right' width='64'>&nbsp;</td>"
            htmhuv = htmhuv + "		<td align='right' width='94'><font size='2'>Totalt SEK</font></td>"
            htmhuv = htmhuv + "		<td align='right' width='101'><font size='2'>" + CStr(attbetL.Text) + "</font></td>"
            htmhuv = htmhuv + "	</tr>"
        End If
        htmhuv = htmhuv + "</table>"
        htmhuv = htmhuv + "</div>"
        htmhuv = htmhuv + "<p>&nbsp;</div>"


        HTMLrader = htmhuv



    End Function
    Function htmlfot()
        htmhuv = ""
        htmhuv = htmhuv + "<p>&nbsp;</div>"
        htmhuv = htmhuv + "<div align='left'>"
        htmhuv = htmhuv + "<table border='0' width='800'>"
        htmhuv = htmhuv + "<tr>"
        htmhuv = htmhuv + "<td width='123'>&nbsp;</td>"
        htmhuv = htmhuv + "<td colspan='2' align='center'><b><font size='2'>" + komment1 + "  " + komment2 + "</font></b></td>"
        htmhuv = htmhuv + "<td width='190'><font size='2'>Handledare: " + VarrefTB.Text + "</font></td>"
        htmhuv = htmhuv + "</tr>"
        htmhuv = htmhuv + "<tr>"
        htmhuv = htmhuv + "<td width='123' align='center' bgcolor='#D8D5A9'><b>"
        htmhuv = htmhuv + "<font size='2'>Bankgiro</font></b></td>"
        htmhuv = htmhuv + "<td width='144' align='center' bgcolor='#D8D5A9'><b>"
        htmhuv = htmhuv + "<font size='2'>Postgiro</font></b></td>"
        htmhuv = htmhuv + "<td width='325' align='center' bgcolor='#D8D5A9'><b>"
        htmhuv = htmhuv + "<font size='2'>" + banknamn + "</font></b></td>"
        htmhuv = htmhuv + "<td width='190' align='center' bgcolor='#D8D5A9'><b>"
        htmhuv = htmhuv + "<font size='2'>Regnummer</font></b></td>"
        htmhuv = htmhuv + "</tr>"
        htmhuv = htmhuv + "<tr>"
        htmhuv = htmhuv + "<td width='123' align='center'><font size='2'>" + bankgiro + "</font></td>"
        htmhuv = htmhuv + "<td width='144' align='center'><font size='2'>" + postgiro + "</font></td>"
        htmhuv = htmhuv + "<td width='325' align='center'><font size='2'>" + bankkonto + "</font></td>"
        htmhuv = htmhuv + "<td width='190' align='center'><font size='2'>" + regnr + "</font></td>"
        htmhuv = htmhuv + "</tr>"
        htmhuv = htmhuv + "<tr>"
        htmhuv = htmhuv + "<td width='123'>&nbsp;</td>"
        htmhuv = htmhuv + "<td width='144'>&nbsp;</td>"
        htmhuv = htmhuv + "<td width='325'>&nbsp;</td>"
        htmhuv = htmhuv + "<td width='190'>&nbsp;</td>"
        htmhuv = htmhuv + "</tr>"
        htmhuv = htmhuv + "</table>"
        htmhuv = htmhuv + "</div>"
        htmhuv = htmhuv + "</div>"
        htmlfot = htmhuv
    End Function
    Sub lagerupdateringrutin()
        Dim dummy As String
        For i As Integer = 0 To rr
            pprodnr = DataGridView1.Rows(i).Cells(0).Value
            prantal = DataGridView1.Rows(i).Cells(4).Value
            If Len(pprodnr) > 0 Then
                dummy = hamptaprodukt("LikaNr")
                dummy = productsaldouppdatering(pprodnr, prantal)
            End If

        Next

    End Sub
    Function productsaldouppdatering(ByVal pnr As String, ByVal redlager As Double)
        Dim cn As OdbcConnection, mySQL As String
        Dim connStr As String
        connStr = "DSN=" + odbcsourcer + "; Database=" + databasnamn + ";Uid=v2000;Pwd=" + odbclosen
        cn = New OdbcConnection(connStr)
        cn.Open()
        mySQL = "UPDATE produktreg "
        mySQL = mySQL + "SET antal = " + kommatillpunkt(CStr(pantal + (redlager) * -1)) + " "
        mySQL = mySQL + " Where prodnummer = '" + pnr + "' "

        Dim myCmd As New OdbcCommand(mySQL, cn)
        myCmd.ExecuteNonQuery()
        cn.Close()
        productsaldouppdatering = "Yes"
    End Function
    Sub skapaFHDfil()
        Dim pltotal As Double
        Dim filnamn As String

        Dim panth As String, rabbel As String
        Dim lrabbel As Integer
        Dim bcounter As Integer, fcounter As Integer
        Dim dummy As String, pnum As String, lepnum As Integer
        Dim pant As String, lepant As Integer, xmltext As String, antalart As Integer
        Dim pltot As String, lpltot As Integer
        'On Error Resume Next
        Dim capris As String, lcapris As Integer, pantr As Double
        Dim apris As String, lapris As Integer



        REM-------------------------------------------Filtillverkning
        filnamn = "TillFHD" + ".XML"
        fcounter = 0
        xmltext = "<?xml version='1.0' encoding='iso-8859-1' ?>"
        xmltext = xmltext + "<sale>"
        xmltext = xmltext + "<cashier>" + VarrefTB.Text + "   ID:" + CStr(nummer) + "</cashier>"
        antalart = 0
        If Len(Lev1) > 3 Then
            xmltext = xmltext + "<text>" + "Köpare:" + "</text>"
            xmltext = xmltext + "<text>" + nullhantering(xmltecken(Lev1), "S") + "</text>"
            If Len(xmltecken(Lev2)) > 1 Then xmltext = xmltext + "<text>" + nullhantering(xmltecken(Lev2), "S") + "</text>"
            If Len(xmltecken(lev3)) > 1 Then xmltext = xmltext + "<text>" + nullhantering(xmltecken(lev3), "S") + "</text>"
            xmltext = xmltext + "<text>" + "-" + "</text> "
        End If

        For i As Integer = 0 To rr
            pprodnr = DataGridView1.Rows(i).Cells(0).Value
            If Len(pprodnr) > 0 Then
                dummy = hamptaprodukt("LikaNr")
            End If
            pprodnamn = DataGridView1.Rows(i).Cells(1).Value
            pEnhet = DataGridView1.Rows(i).Cells(2).Value
            prantal = DataGridView1.Rows(i).Cells(4).Value
            putpris = DataGridView1.Rows(i).Cells(5).Value
            psumma = DataGridView1.Rows(i).Cells(6).Value
            capris = DataGridView1.Rows(i).Cells(7).Value
            xmltext = xmltext + "<text>" + xmltecken(nullhantering(pprodnamn, "S"))
            xmltext = xmltext + "  " + nullhantering(pantiforpack, "S")
            xmltext = xmltext + "  " + nullhantering(pEnhet, "S") + "</text>"
            pnum = CStr(nullhantering(pprodnr, "S")) : lepnum = Len(pnum)
            pant = CStr(nullhantering(prantal, "T")) : lepant = Len(pant)
            If nullhantering(prantal, "T") < 0 Then
                antalart = nullhantering(prantal, "T") * -1 + antalart
            Else
                antalart = antalart + nullhantering(prantal, "T")
            End If
            xmltext = xmltext + "<text>" + Space(3) + pant
            capris = Format$((avrund(capris * 1.25)), "######.00")
            lcapris = Len(Format$((avrund(capris * 1.25)), "######.00"))
            xmltext = xmltext + Space(12 - lcapris) + capris
            apris = Format$(putpris, "######.00")
            lapris = Len(Format$(putpris, "######.00"))
            xmltext = xmltext + Space(12 - lapris) + apris
            pltot = CStr(Format$(psumma, "######.00"))
            lpltot = Len(pltot)
            xmltext = xmltext + Space(14 - lpltot) + pltot + "</text>"

            If pant < 0 Then
                pltotal = pltotal + CDbl(pltot)
                 pantr = pantr + CDbl(pant)
            End If

            bcounter = bcounter + 1
        Next

        If Val(ExrabL.Text) <> 0 Then
            rabbel = CStr(Format$((ExrabL.Text), "######.00"))
            lrabbel = Len(rabbel)

            xmltext = xmltext + "<text>Extra rabatt" + exrabTB.Text + "%" + Space(20 - lrabbel) + "-" + rabbel + "</text>"
        End If
        xmltext = xmltext + "<topay>" + CStr(Format$(CDbl(attbetL.Text), "######.00")) + "</topay>"
        xmltext = xmltext + "<moms1>" + CStr(Format$(CDbl(momsbelL.Text), "######.00")) + "</moms1>"
        xmltext = xmltext + "<antalart>" + CStr(antalart) + "</antalart>"

        If pantr <> 0 Then
            panth = CStr(CDbl(pantr) * -1)

            xmltext = xmltext + "<retur>" + CStr(Format$(pltotal, "######.00")) + "</retur>"
            xmltext = xmltext + "<antalretur>" + panth + "</antalretur>"
        End If

        xmltext = xmltext + "<payment>" + "Kontant" + "</payment>"
        xmltext = xmltext + "<paymentsum>" + CStr(Format$(CDbl(nullhantering(KontantTB.Text, "T")), "######.00")) + "</paymentsum>"
        xmltext = xmltext + "<payment>" + "Kreditkort" + "</payment>"
        xmltext = xmltext + "<paymentsum>" + CStr(Format$(CDbl(nullhantering(KortTB.Text, "T")), "######.00")) + "</paymentsum>"

        xmltext = xmltext + "<signr>1</signr><sigrow1></sigrow1><sigrow2></sigrow2>"
        xmltext = xmltext + "</sale>"



        dummy = SendToFHD(xmltext)
    End Sub
    Function xmltecken(ByVal trad As String)
        Dim lside As String, rside As String

        If InStr(trad, "&") > 0 Then
            lside = Strings.Left$(trad, InStr(trad, "&") - 1)
            rside = Strings.Right(trad, Len(trad) - InStr(trad, "&"))
            trad = lside + "&amp;" + rside
        End If
        While InStr(trad, ">") > 0
            lside = Strings.Left$(trad, InStr(trad, ">") - 1)
            rside = Strings.Right$(trad, Len(trad) - InStr(trad, ">"))
            trad = lside + "&gt;" + rside
        End While
        While InStr(trad, "<") > 0
            lside = Strings.Left$(trad, InStr(trad, "<") - 1)
            rside = Strings.Right$(trad, Len(trad) - InStr(trad, "<"))
            trad = lside + "&lt;" + rside
        End While
        While InStr(trad, Strings.Chr(34)) > 0
            lside = Strings.Left$(trad, InStr(trad, Strings.Chr(34)) - 1)
            rside = Strings.Right$(trad, Len(trad) - InStr(trad, Strings.Chr(34)))
            trad = lside + "&quot;" + rside
        End While

        While InStr(trad, "'") > 0
            lside = Strings.Left$(trad, InStr(trad, "'") - 1)
            rside = Strings.Right$(trad, Len(trad) - InStr(trad, "'"))
            trad = lside + "&apos;" + rside
        End While


        xmltecken = trad
    End Function
    Private Function SendToFHD(ByVal fil As String)
        On Error GoTo noconn
        Dim ip As String
        ip = "127.0.0.1" ' local host
        ip = ipadress 'hämtad från datorn
        ip = "213.50.35.205" 'fhd simulatorn


        Dim clientSocket As New System.Net.Sockets.TcpClient()
        clientSocket.Connect(ip, 1981)
        Dim serverStream As NetworkStream = clientSocket.GetStream()
        Dim outStream As Byte() = System.Text.Encoding.Default.GetBytes(fil)
        serverStream.Write(outStream, 0, outStream.Length)
        serverStream.Flush()
        serverStream.Close()
        clientSocket.Close()
        SendToFHD = "Yes"
noconn:
        If Err.Number <> 0 Then
            MessageBox.Show("IP förbindelse med FHD programet saknas")
        End If
        Stop
    End Function
    Sub radraknare()
        rr = 0
        y = DataGridView1.CurrentCellAddress.Y
        rr = y
        rrl.Text = rr
        rrl.Refresh()
        If rr > rmax Then rmax = rr
        rmaxl.Text = rmax
        rmaxl.Refresh()
    End Sub

    Private Sub DataGridView1_RowEnter(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.RowEnter

    End Sub

    Private Sub DataGridView1_RowStateChanged(sender As Object, e As System.Windows.Forms.DataGridViewRowStateChangedEventArgs) Handles DataGridView1.RowStateChanged

    End Sub
End Class