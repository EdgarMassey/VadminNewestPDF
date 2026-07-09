Imports System.Data.Odbc
Imports System.Net.Mail
Imports System.Net.Sockets
Imports System.IO
Imports VB = Microsoft.VisualBasic
Imports System.Net


Public Class OrdFaktF
    Dim r As Integer, c As Integer, ordernummer As String, utskriftsort As String, dummy As String, l As Integer, pdfval As String
    Dim lev1 As String, lev2 As String, lev3 As String, y As Integer, levsatt(10) As String, levvilk(10) As String, ticks As Integer
    Dim pagenr As Integer, HistFaktUT As String, faktdatum As String, antpdf As Integer, tiden As String = "Not yet"
    Dim epostfrom As String, eposttill As String, epostsubject As String, epostbody As String, fromsystem As String, mytimestamp As String
    Dim regdatum As String, handpmoms As Double = 0, oresut As Double = 0, pdfprintername As String
    Dim Copytyp As String = "", FardigStatus As String = ""
    Dim startatex As Integer = 0, fortext As Integer = 0

    Private Sub OrdFaktF_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.KeyPreview = True
        Dim rect As Rectangle = Screen.PrimaryScreen.WorkingArea
        Me.Top = (rect.Height / 2) - (Me.Height / 2)
        Me.Left = (rect.Width / 2) - (Me.Width / 2)
        Ver.Text = "Version: " + "20260709a"
        ' pdfprintername = "PDF Creator"
        'pdfprintername = "PDF Architect 5"
        pdfprintername = "WIN2PDF"
        Huvud.Text = Prognamn + " - Order/Fakturering"
        'viktiga rutiner-------------------------------------
        ' getpdfinfo
        'skickapdf
        'skickaepost
        today = Format(Now, "yyyy-MM-dd")
        odbcsourcer = "VadminODBC"
        odbclosen = "alfons"
       
        nolla()
        FakturaRB.Checked = True
        VarrefTB.Select()
        'LeveranssattTB.Text = hamptalevsatt(LeveranssattTB.Text)





        PDFCB.Enabled = True
        If TimerB.Text = "Start timern" Then

            Timer1.Interval = 120000  '2 minute
            Timer1.Start()
            TimerB.Text = "Timer enabled"
        Else


            Timer1.Stop()
            TimerB.Text = "Start timern"
        End If

    End Sub


    Private Sub KundregToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If sakerhet < 4 Then
            franprog = "OrderFakt"

        Else
            MessageBox.Show("Otillräklig behörighet!")
        End If

    End Sub


    Private Sub ProduktlistaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        franprog = "OrderFakt"

    End Sub

    Private Sub ProduktregToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If sakerhet < 4 Then

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
        Else
            DataGridView1.Left = 0
            DataGridView1.Top = 214
            DataGridView1.Height = 164
        End If

        With (Me.DataGridView1.Columns)
            .Add("", "ProdNr")
            .Add("", "ProdNamn")
            .Add("", "Enhet")
            .Add("", "PF")
            .Add("", "Ant")
            .Add("", "Apris")
            .Add("", "Summa")
            .Add("", "CaPris")
            .Add("", "Rab%")
            .Add("", "Enhetspris")
        End With
        With Me.DataGridView1.Rows
            For Me.r = 1 To 2000
                .Add("")
            Next r
        End With
        DataGridView1.CurrentCell = DataGridView1.Item(0, 0)
        DataGridView1.CurrentCell.Style.SelectionBackColor = Color.White
        r = 0
        With DataGridView1.ColumnHeadersDefaultCellStyle
            .BackColor = Color.Gray
            .ForeColor = Color.White
            .Font = New Font(DataGridView1.Font, FontStyle.Bold)
        End With
        DataGridView1.RowHeadersWidth = 22
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).Width = 280
        'enhet
        DataGridView1.Columns(2).Width = 50
        'prisfaktor
        DataGridView1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight

        If orderformat = "Färg" Or orderformat = "MedPrisFaktor" Then
            DataGridView1.Columns(3).Width = 40
        Else
            DataGridView1.Columns(3).Width = 0
        End If


        'antal
        DataGridView1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns(4).Width = 40
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
        DataGridView1.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns(7).Width = 70
        Me.DataGridView1.Columns.Item(7).DefaultCellStyle.Format = "n2"
        DataGridView1.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns(8).Width = 53
        Me.DataGridView1.Columns.Item(8).DefaultCellStyle.Format = "n1"
        DataGridView1.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns(9).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns(9).Width = 70
        Me.DataGridView1.Columns.Item(9).DefaultCellStyle.Format = "n2"

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
            RakningRB.Checked = True
            UtTyp.Enabled = False

            MomsTypTB.Text = 4
            DataGridView1.Focus()
            DataGridView1.CurrentCell = DataGridView1.Item(0, 0)
            DataGridView1.CurrentCell.Style.SelectionBackColor = Color.LightBlue
        Else

            UtTyp.Enabled = True
            makegrid("SimpleKredit")
            FakturaRB.Checked = True
            MomsTypTB.Text = stdmoms
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

            KontaktmanTB.Select()
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
        If Strings.InStr(kepost, "@") > 0 Then
            PDFCB.Checked = True
        Else
            PDFCB.Checked = False
        End If
        kfaktbettypL.Text = kfaktlevtyp
    End Sub

    Private Sub KundnamnTB_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles KundnamnTB.KeyUp
        If e.KeyCode = Keys.Enter Then
            Dim dummy As String
            dummy = hamptakund("LikaNamn")
            skrivakundtillform()

            KontaktmanTB.Select()
        End If
    End Sub


    Private Sub kundnrTB_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kundnrTB.TextChanged
        kkundnr = kundnrTB.Text
    End Sub

    Private Sub KundnamnTB_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KundnamnTB.TextChanged
        kkundnamn = KundnamnTB.Text
    End Sub


    Private Sub BakB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
      
    End Sub

    Private Sub FramB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

       


    End Sub

    Private Sub DataGridView1_CellEnter(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEnter
       
        radraknare()
        
    End Sub

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
       
        radraknare()
       

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


                If dummy = "No" Then
                    DataGridView1.CurrentCell = DataGridView1.Item(0, rr - 1)
                    Exit Sub
                End If
                DataGridView1.Item(0, rr - 1).Value = pprodnr
                DataGridView1.Item(1, rr - 1).Value = pprodnamn
                DataGridView1.Item(2, rr - 1).Value = pEnhet
                DataGridView1.Item(3, rr - 1).Value = pprisfaktor
                DataGridView1.Item(4, rr - 1).Value = 1
                If MomsTypTB.Text = 4 Then

                    tillfpris = (getkundpris(krabattmall, kkundnr, pprodnr))

                    tillfpris = tillfpris + (tillfpris * momsv(1))
                    If orderformat = "Färg" Then

                        DataGridView1.Item(0, rr - 1).Value = pprodnr
                        DataGridView1.Item(5, rr - 1).Value = tillfpris * DataGridView1.Item(3, rr - 1).Value
                        DataGridView1.Item(9, rr - 1).Value = tillfpris
                        DataGridView1.Item(7, rr - 1).Value = (putpris1 + (putpris1 * momsv(1))) * DataGridView1.Item(4, rr - 1).Value * DataGridView1.Item(3, rr - 1).Value
                        DataGridView1.Item(8, rr - 1).Value = 100 - ((DataGridView1.Item(5, rr - 1).Value / DataGridView1.Item(7, rr - 1).Value) * 100)
                        DataGridView1.Item(6, rr - 1).Value = (DataGridView1.Item(5, rr - 1).Value * DataGridView1.Item(4, rr - 1).Value)
                    Else
                        DataGridView1.Item(5, rr - 1).Value = tillfpris
                        DataGridView1.Item(9, rr - 1).Value = tillfpris
                        DataGridView1.Item(7, rr - 1).Value = (putpris1 + (putpris1 * momsv(1))) * DataGridView1.Item(4, rr - 1).Value
                        DataGridView1.Item(8, rr - 1).Value = 100 - ((DataGridView1.Item(5, rr - 1).Value / DataGridView1.Item(7, rr - 1).Value) * 100)
                        DataGridView1.Item(6, rr - 1).Value = (DataGridView1.Item(3, rr - 1).Value * DataGridView1.Item(5, rr - 1).Value * DataGridView1.Item(4, rr - 1).Value)
                    End If
                Else
                    tillfpris = (getkundpris(krabattmall, kkundnr, pprodnr))

                    tillfpris = tillfpris
                    If orderformat = "Färg" Then

                        DataGridView1.Item(5, rr - 1).Value = tillfpris * DataGridView1.Item(3, rr - 1).Value
                        DataGridView1.Item(9, rr - 1).Value = tillfpris
                        DataGridView1.Item(7, rr - 1).Value = (putpris1 * DataGridView1.Item(3, rr - 1).Value)
                        DataGridView1.Item(8, rr - 1).Value = 100 - ((DataGridView1.Item(5, rr - 1).Value / DataGridView1.Item(7, rr - 1).Value) * 100)
                        DataGridView1.Item(6, rr - 1).Value = (DataGridView1.Item(5, rr - 1).Value * DataGridView1.Item(4, rr - 1).Value)
                    Else
                        DataGridView1.Item(5, rr - 1).Value = tillfpris
                        DataGridView1.Item(9, rr - 1).Value = tillfpris
                        DataGridView1.Item(7, rr - 1).Value = (putpris1 * DataGridView1.Item(3, rr - 1).Value)
                        DataGridView1.Item(8, rr - 1).Value = 100 - ((DataGridView1.Item(5, rr - 1).Value / DataGridView1.Item(7, rr - 1).Value) * 100)
                        DataGridView1.Item(6, rr - 1).Value = (DataGridView1.Item(3, rr - 1).Value * DataGridView1.Item(5, rr - 1).Value * DataGridView1.Item(4, rr - 1).Value)
                    End If
                End If
                If Len(pprodnamn) < 2 Then
                    DataGridView1.CurrentCell = DataGridView1.Item(1, rr - 1)
                End If
                'Prodnr = 0
                'Prodnamn = 1
                'Enhet = 2
                'Prisfaktor = 3
                'Antal = 4
                'Apris = 5
                'Summa = 6
                'CaPris = 7
                'Rabatt = 8
                'enhetspris=9


                If tillfpris <> 0 Then
                    DataGridView1.CurrentCell = DataGridView1.Item(4, rr - 1)

                Else
                    DataGridView1.CurrentCell = DataGridView1.Item(1, rr - 1)

                End If
            ElseIf c = 1 Then
                Dim dummy As String

                If Len(pprodnr) > 1 Then
                    If Len(pprodnamn) < 2 Then
                    End If
                Else
                    pprodnamn = DataGridView1.Item(1, rr - 1).Value
                    dummy = hamptaprodukt("LikaNamn")

                    DataGridView1.Item(0, rr - 1).Value = pprodnr
                    DataGridView1.Item(1, rr - 1).Value = pprodnamn
                    DataGridView1.Item(2, rr - 1).Value = pEnhet
                    DataGridView1.Item(3, rr - 1).Value = pprisfaktor
                    DataGridView1.Item(4, rr - 1).Value = 1
                End If
            
            If MomsTypTB.Text = 4 Then
                tillfpris = (getkundpris(krabattmall, kkundnr, pprodnr))
                tillfpris = tillfpris + (tillfpris * momsv(1))
                If orderformat = "Färg" Then
                    DataGridView1.Item(5, rr - 1).Value = tillfpris * DataGridView1.Item(3, rr - 1).Value
                    DataGridView1.Item(9, rr - 1).Value = tillfpris
                    DataGridView1.Item(7, rr - 1).Value = (putpris1 + (putpris1 * momsv(1))) * DataGridView1.Item(3, rr - 1).Value
                    DataGridView1.Item(8, rr - 1).Value = 100 - ((DataGridView1.Item(5, rr - 1).Value / DataGridView1.Item(7, rr - 1).Value) * 100)
                Else
                    DataGridView1.Item(5, rr - 1).Value = tillfpris
                    DataGridView1.Item(9, rr - 1).Value = tillfpris
                    DataGridView1.Item(7, rr - 1).Value = (putpris1 + (putpris1 * momsv(1))) * DataGridView1.Item(3, rr - 1).Value
                    DataGridView1.Item(8, rr - 1).Value = 100 - ((DataGridView1.Item(5, rr - 1).Value / DataGridView1.Item(7, rr - 1).Value) * 100)
                End If


            Else
                tillfpris = (getkundpris(krabattmall, kkundnr, pprodnr))
                DataGridView1.Item(5, rr - 1).Value = tillfpris
                DataGridView1.Item(9, rr - 1).Value = tillfpris
                DataGridView1.Item(7, rr - 1).Value = putpris1
                'DataGridView1.Item(8, rr - 1).Value = (putpris1 * pprisfaktor * 1) - (tillfpris * pprisfaktor * 1)
                DataGridView1.Item(8, rr - 1).Value = 100 - ((DataGridView1.Item(5, rr - 1).Value / DataGridView1.Item(7, rr - 1).Value) * 100)
            End If
            If orderformat = "Färg" Then
                DataGridView1.Item(6, rr - 1).Value = (DataGridView1.Item(5, rr - 1).Value * DataGridView1.Item(4, rr - 1).Value)
            Else
                DataGridView1.Item(6, rr - 1).Value = (DataGridView1.Item(3, rr - 1).Value * DataGridView1.Item(5, rr - 1).Value * DataGridView1.Item(4, rr - 1).Value)
            End If


            If tillfpris <> 0 Then
                DataGridView1.CurrentCell = DataGridView1.Item(4, rr - 1)
            Else
                DataGridView1.CurrentCell = DataGridView1.Item(4, rr - 1)
            End If

        ElseIf c = 3 Then
            ''PRISFAKTOR
            DataGridView1.Item(6, rr - 1).Value = (DataGridView1.Item(3, rr - 1).Value * DataGridView1.Item(5, rr - 1).Value * DataGridView1.Item(4, rr - 1).Value)
            'DataGridView1.Item(8, rr - 1).Value = (DataGridView1.Item(7, rr - 1).Value * DataGridView1.Item(3, rr - 1).Value * DataGridView1.Item(4, rr - 1).Value) - (DataGridView1.Item(5, rr - 1).Value * DataGridView1.Item(3, rr - 1).Value * DataGridView1.Item(4, rr - 1).Value)
            DataGridView1.CurrentCell = DataGridView1.Item(4, rr - 1)
            If MomsTypTB.Text = 4 Then
                DataGridView1.Item(7, rr - 1).Value = putpris1 + (putpris1 * momsv(1))
            ElseIf MomsTypTB.Text = 1 Then
                DataGridView1.Item(7, rr - 1).Value = putpris1
            End If
            DataGridView1.CurrentCell = DataGridView1.Item(4, rr)
        ElseIf c = 4 Then
            If orderformat = "Färg" Then
                DataGridView1.Item(6, rr - 1).Value = (DataGridView1.Item(5, rr - 1).Value * DataGridView1.Item(4, rr - 1).Value)
            Else
                DataGridView1.Item(6, rr - 1).Value = (DataGridView1.Item(3, rr - 1).Value * DataGridView1.Item(5, rr - 1).Value * DataGridView1.Item(4, rr - 1).Value)
            End If
            If DataGridView1.Item(5, rr - 1).Value = 0 Then
                DataGridView1.CurrentCell = DataGridView1.Item(5, rr - 1)
            Else
                DataGridView1.CurrentCell = DataGridView1.Item(0, rr)
            End If

            summering()

        ElseIf c = 5 Then
            'apris
            DataGridView1.Item(6, rr - 1).Value = (DataGridView1.Item(3, rr - 1).Value * DataGridView1.Item(5, rr - 1).Value * DataGridView1.Item(4, rr - 1).Value)
            DataGridView1.Item(9, rr - 1).Value = DataGridView1.Item(5, rr - 1).Value / DataGridView1.Item(3, rr - 1).Value
            DataGridView1.Item(8, rr - 1).Value = 100 - ((DataGridView1.Item(5, rr - 1).Value / DataGridView1.Item(7, rr - 1).Value) * 100)
            summering()
            DataGridView1.CurrentCell = DataGridView1.Item(0, rr)
        ElseIf c = 8 Then
            'rabbatt %
            DataGridView1.Item(5, rr - 1).Value = (1 - ((DataGridView1.Item(8, rr - 1).Value) / 100)) * DataGridView1.Item(7, rr - 1).Value
            If orderformat = "Färg" Then
                DataGridView1.Item(6, rr - 1).Value = (DataGridView1.Item(5, rr - 1).Value * DataGridView1.Item(4, rr - 1).Value)
            Else
                DataGridView1.Item(6, rr - 1).Value = (DataGridView1.Item(3, rr - 1).Value * DataGridView1.Item(5, rr - 1).Value * DataGridView1.Item(4, rr - 1).Value)
            End If
            DataGridView1.Item(9, rr - 1).Value = DataGridView1.Item(5, rr - 1).Value / DataGridView1.Item(3, rr - 1).Value
            DataGridView1.Item(8, rr - 1).Value = 100 - ((DataGridView1.Item(5, rr - 1).Value / DataGridView1.Item(7, rr - 1).Value) * 100)
            summering()
            DataGridView1.CurrentCell = DataGridView1.Item(0, rr)
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
            If orderformat = "Färg" And momstyp = 4 Then
                DataGridView1.Item(6, r - 1).Value = Format$((DataGridView1.Item(5, r - 1).Value * DataGridView1.Item(4, r - 1).Value), "###,##0.00")
            Else

                DataGridView1.Item(6, r - 1).Value = Format$((DataGridView1.Item(3, r - 1).Value * DataGridView1.Item(5, r - 1).Value * DataGridView1.Item(4, r - 1).Value), "###,##0.00")

            End If

        End If
        'On Error Resume Next
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
            MOMSLabel.Text = "Moms " + CStr(moms1) + "%"
        ElseIf MomsTypTB.Text = 2 Then
            momsbel = (TOT - exrab) * momsv(2)
            momsbelL.Text = Format$(momsbel, "###,##0.00")
            attbet = TOT - exrab + momsbel
            MOMSLabel.Text = "Moms " + CStr(moms2) + "%"
        ElseIf MomsTypTB.Text = 3 Then
            momsbel = (TOT - exrab) * momsv(3)
            momsbelL.Text = Format$(momsbel, "###,##0.00")
            attbet = TOT - exrab + momsbel
            MOMSLabel.Text = "Moms " + CStr(moms3) + "%"

        ElseIf MomsTypTB.Text = 4 Then
            momsbel = (TOT - exrab) * momsv(4)
            momsbelL.Text = Format$(momsbel, "###,##0.00")
            attbet = TOT - exrab
            MOMSLabel.Text = "Därav moms " + CStr(moms1) + "%"
        ElseIf MomsTypTB.Text = 0 Then

            momsbelL.Text = Format$(0, "###,##0.00")
            attbet = TOT - exrab
            MOMSLabel.Text = "Ingen moms"
        End If
        attbet = avrund(attbet)
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
        Dim kortt As Double, kontantt As Double, attbetala As Double
        kortt = avrund100(Val(KortTB.Text))
        kontantt = avrund100(Val(KontantTB.Text))
        attbetala = avrund100(Val(attbetL.Text))

        '----------------------------------------Räkning
        If RakningRB.Checked = True Then

            If attbetala <> kontantt + kortt Then
                MessageBox.Show("Välj betalningsätt")

                Return
            End If

            If Val(CInt(nullhantering(kvittorems, "T"))) = 1 Then
                skapaFHDfil()
            Else
                If FGCB.CheckState = 1 Then
                    PrintDocument1.PrinterSettings.PrinterName = invoicepr
                    PrintPreviewDialog1.Document = PrintDocument1
                    PrintPreviewDialog1.ShowDialog()
                Else
                    PrintDocument1.PrinterSettings.Copies = antfakt
                    PrintDocument1.PrinterSettings.PrinterName = invoicepr
                    PrintDocument1.Print()
                End If
            End If
            ProgressBar1.Value = 50
            getnummer("Next", "Nästa räkningsnummer")
            lagerupdateringrutin()
            ProgressBar1.Value = 70
            TillOrderArkiv()
            TillKassaBoken()
            ProgressBar1.Value = 85
            suddaorder(OrdnrTB.Text)
            nolla()
            ProgressBar1.Value = 100
            '----------------------------------Faktura----------------------------------------------------------------------------
        ElseIf FakturaRB.Checked = True Then
            If FGCB.CheckState = 1 Then
                PrintDocument1.PrinterSettings.PrinterName = invoicepr
                PrintPreviewDialog1.Document = PrintDocument1
                PrintPreviewDialog1.ShowDialog()
            Else
                PrintDocument1.PrinterSettings.Copies = antfakt
                PrintDocument1.PrinterSettings.PrinterName = invoicepr
                PrintDocument1.Print()
            End If

            If pdfval = "YES" Then

                If PDFCB.Checked = True Then

                    PrintDocument1.PrinterSettings.Copies = antfakt
                    PrintDocument1.PrinterSettings.PrinterName = pdfpr
                    PrintDocument1.Print()
                End If
            End If

            If HistFaktUT = "YES" Then
                nolla()
                GoTo slut20
            End If
            Dim Message As String = "Godkänns utskriften ?"
            Dim Caption As String = "Kontrollera"
            Dim Buttons As MessageBoxButtons = MessageBoxButtons.OKCancel
            Dim Result As DialogResult

            Result = MessageBox.Show(Message, Caption, Buttons)

            If Result = 1 Then
                If PDFCB.Checked = True Then
                    tillpdfrutin()
                End If
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
                If utskriftstyp = "Multi" Then
                    For i = 0 To iii
                        suddafoljesedel(foljnr(i))
                    Next
                Else
                    suddafoljesedel(FoljnrTB.Text)
                End If
                nolla()
                ProgressBar1.Value = 100
            End If
slut20:
            '-----------------------------------------------Följesedel
        ElseIf FoljesedelRB.Checked = True Then
            If FGCB.CheckState = 1 Then
                PrintDocument2.PrinterSettings.PrinterName = deliverypr
                PrintPreviewDialog1.Document = PrintDocument2
                PrintPreviewDialog1.ShowDialog()
                Else
                PrintDocument2.PrinterSettings.Copies = antfakt
                PrintDocument2.PrinterSettings.PrinterName = deliverypr
                PrintDocument2.Print()
            End If
                If PDFCB.Checked = True Then
                    PrintDocument1.PrinterSettings.Copies = antfakt
                    PrintDocument1.PrinterSettings.PrinterName = pdfpr
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
            '-----------------------------offert



        End If

    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim x As Integer, y As Integer, yy As Integer, point1 As Point, point2 As Point, penhetb As String, prantalb As Double, kkilo As Double, genomutpris As Double
        Dim dblprod As String, dblprod1 As String, dblprod2 As String, pprisfaktorb As Double, putprisb As Double, pprodnrb As String, pprodnamnb As String
        Dim radavstand As Integer = 0, toppos As Integer, lll As Integer = 0
        If specialversion = "Malmgren" Then
            radavstand = 22
        Else
            radavstand = 16
        End If

        'väljerskrivare
        'PrintDocument1.PrinterSettings.PrinterName = invoicepr
        If HistFaktUT <> "YES" Then

            'getnummer("Kolla", "Nästa fakturanummer")
        End If

        'header
        x = 40 : y = 8
        If specialversion = "Målarboden" Then
            toppos = 630
        Else
            toppos = 580
        End If
        If FakturaRB.Checked = True Then
            If Val(attbetL.Text) < 0 Then
                e.Graphics.DrawString("Kreditfaktura", LargeFontB, BlackFont, toppos, y) : y = y + 22
            Else
                If ProformaCB.Checked = True Then
                    e.Graphics.DrawString("ProformaFaktura", LargeFontB, BlackFont, toppos, y) : y = y + 22
                Else

                    e.Graphics.DrawString("Faktura", LargeFontB, BlackFont, toppos, y) : y = y + 22
                End If
            End If

            'If OBSL.Text = "Historisk faktura" Then
            nummer = GetFaktTB.Text
            'Else
            If ProformaCB.Checked = True Then

                'nummer = getnummer("Kolla", "Nästa profakturanummer")
            Else
                    If OmbudCB.Checked = True Then
                        nummer = "000"
                    Else
                    'nummer = getnummer("Kolla", "Nästa fakturanummer")
                End If
                End If

                'End If

                ElseIf FoljesedelRB.Checked = True Then
            e.Graphics.DrawString("Följesedel", LargeFontB, BlackFont, 580, y) : y = y + 22
            If Len(FoljnrTB.Text) < 1 Then
                'nummer = getnummer("Kolla", "Nästa följesedelnummer")
                FoljnrTB.Text = nummer
            Else
                nummer = FoljnrTB.Text
            End If
        ElseIf OrderbekraftRB.Checked = True Then
            e.Graphics.DrawString("Orderbekräftelse", LargeFontB, BlackFont, toppos, y) : y = y + 22
            nummer = OrdnrTB.Text
        ElseIf OffertRB.Checked = True Then
            e.Graphics.DrawString("Offert", LargeFontB, BlackFont, toppos, y) : y = y + 22
            nummer = OrdnrTB.Text
        ElseIf RakningRB.Checked = True Then
            e.Graphics.DrawString("KVITTO", LargeFontB, BlackFont, toppos, y) : y = y + 22
            nummer = getnummer("Kolla", "Nästa räkningsnummer")
        End If

        y = 30
        e.Graphics.DrawString("Nummer: " + nummer, SmallFontR, BlackFont, toppos, y) : y = y + 22
        e.Graphics.DrawString("Kund nr: " + kkundnr, SmallFontR, BlackFont, toppos, y) : y = y + 22
        If HistFaktUT = "YES" Then
            e.Graphics.DrawString("Datum: " + faktdatum, SmallFontR, BlackFont, toppos, y)
        Else

            e.Graphics.DrawString("Datum: " + today, SmallFontR, BlackFont, toppos, y)
        End If

        'If fortrycktpapper <> True Then
        If KlientID = "MD" Then
            Dim img As Image = Image.FromFile(sokvag + "MDlogo.jpg")
            e.Graphics.DrawImage(img, 10, 0)
            x = 40 : y = 70
        ElseIf KlientID = "LS" Then
            Dim img As Image = Image.FromFile(sokvag + "LSlogo.jpg")
            e.Graphics.DrawImage(img, 10, 0)
            x = 40 : y = 70
        ElseIf KlientID = "MB" Then
            Dim img As Image = Image.FromFile(sokvag + "MBlogo.jpg")
            e.Graphics.DrawImage(img, 40, 0)
            x = 40 : y = 70
        ElseIf KlientID = "MC" Then
            Dim img As Image = Image.FromFile(sokvag + "MClogo.jpg")
            e.Graphics.DrawImage(img, 10, 0)
            x = 40 : y = 70

        Else
                x = 40 : y = 8
                e.Graphics.DrawString(Firmanamn, LargeFontB, BlackFont, x, y)
                y = y + 22

            End If


            e.Graphics.DrawString(Postadress1, SmallFontR, BlackFont, x, y) : y = y + 20
            e.Graphics.DrawString(Postadress2, SmallFontR, BlackFont, x, y) : y = y + 20
            e.Graphics.DrawString(Postadress3, SmallFontR, BlackFont, x, y) : y = y + 20
            e.Graphics.DrawString("Telefon: " + telefon, SmallFontR, BlackFont, x, y) : y = y + 20
            e.Graphics.DrawString("Epost: " + epost, SmallFontR, BlackFont, x, y) : y = y + 20
            point1 = New Point(40, y) : point2 = New Point(250, y)
            e.Graphics.DrawLine(blackPen, point1, point2) : y = y - 10
        'End If

        If HistFaktUT = "YES" Then
        Else
            nummer = getnummer("Kolla", "Nästa fakturanummer")
        End If

        y = 30

        y = 240
        e.Graphics.DrawString("Leveranssätt  ", SmallFontR, BlackFont, 44, y)
        e.Graphics.DrawString(":" + LeveranssattTB.Text, SmallFontR, BlackFont, 162, y) : yy = y
        e.Graphics.DrawString("Telefonnr.", SmallFontR, BlackFont, 400, y)
        e.Graphics.DrawString(":" + ktelefon, SmallFontR, BlackFont, 498, y) : y = y + 16

        e.Graphics.DrawString("Leveransvilkor  ", SmallFontR, BlackFont, 44, y)
        e.Graphics.DrawString(":" + LevvilkorTB.Text, SmallFontR, BlackFont, 162, y)
        e.Graphics.DrawString("Er referens  ", SmallFontR, BlackFont, 400, y)
        e.Graphics.DrawString(":" + KontaktmanTB.Text, SmallFontR, BlackFont, 498, y) : yy = y : y = y + 16

        e.Graphics.DrawString("Betalningsvilkor ", SmallFontR, BlackFont, 44, y)
        e.Graphics.DrawString(":" + CStr(kkreditdagar) + " dagar", SmallFontR, BlackFont, 162, y)

        e.Graphics.DrawString("Ert Ordernr  ", SmallFontR, BlackFont, 400, y)
        e.Graphics.DrawString(":" + ErOrdernrTB.Text, SmallFontR, BlackFont, 498, y) : y = y + 16

        If FakturaRB.Checked = True Or RakningRB.Checked = True Then
            e.Graphics.DrawString("Förfallodadum  ", SmallFontR, BlackFont, 44, y)
            AddDays(3D)
            If HistFaktUT = "YES" Then
                e.Graphics.DrawString(CDate(faktdatum).AddDays(kkreditdagar), SmallFontR, BlackFont, 162, y)
            Else
                e.Graphics.DrawString(":" + CDate(today).AddDays(kkreditdagar), SmallFontR, BlackFont, 162, y)
            End If
        End If
        e.Graphics.DrawString("Märkning", SmallFontR, BlackFont, 400, y)
        e.Graphics.DrawString(":" + MarkningTB.Text, SmallFontR, BlackFont, 498, y) : y = y + 16

        If OffertRB.Checked = False Then
            e.Graphics.DrawString("Leveransdatum ", SmallFontR, BlackFont, 44, y)
            e.Graphics.DrawString(":" + LevdatumTB.Text, SmallFontR, BlackFont, 162, y)
        End If

        e.Graphics.DrawString("Ert Org nr ", SmallFontR, BlackFont, 400, y)
        e.Graphics.DrawString(":" + kregnr, SmallFontR, BlackFont, 498, y) : y = y + 16
        If specialversion <> "Malmgren" Then
            e.Graphics.DrawString("Ordernummer  ", SmallFontR, BlackFont, 44, y)
            e.Graphics.DrawString(":" + OrdnrTB.Text, SmallFontR, BlackFont, 162, y)
        End If
        e.Graphics.DrawString("Epost ", SmallFontR, BlackFont, 400, y)
        e.Graphics.DrawString(":" + kepost, SmallFontR, BlackFont, 498, y) : y = y + 16

        y = 140
        e.Graphics.DrawString(kkundnamn, OrdFontB, BlackFont, 400, y) : y = y + 20
        e.Graphics.DrawString(kadress1, SmallFontR, BlackFont, 400, y) : y = y + 16
        e.Graphics.DrawString(kadress2, SmallFontR, BlackFont, 400, y) : y = y + 16
        e.Graphics.DrawString(kpostnr + " " + kort, SmallFontR, BlackFont, 400, y) : y = y + 16
        e.Graphics.DrawString(kland, SmallFontR, BlackFont, 400, y) : y = y + 16



        'rubriker
        y = 340
        e.Graphics.DrawString("Prodnr:", SmallFontB, BlackFont, 44, y)

        If specialversion = "Malmgren" Then
            e.Graphics.DrawString("Prodnamn", SmallFontB, BlackFont, 164, y)
        Else
            e.Graphics.DrawString("Prodnamn", SmallFontB, BlackFont, 146, y)
        End If

        If orderformat = "MedPrisFaktor" Then
            If specialversion = "Malmgren" Then
                e.Graphics.DrawString("Förp", SmallFontB, BlackFont, 400, y)
            Else
                e.Graphics.DrawString("PF", SmallFontB, BlackFont, 400, y)
            End If
        End If
        e.Graphics.DrawString("Enhet", SmallFontB, BlackFont, 440, y)
        e.Graphics.DrawString("Antal", SmallFontB, BlackFont, 520, y)
        If specialversion = "Malmgren" Then
            e.Graphics.DrawString("Pris/Enh", SmallFontB, BlackFont, 650, y)
        Else
            e.Graphics.DrawString("ápris", SmallFontB, BlackFont, 650, y)
        End If
        If specialversion = "Malmgren" And FoljesedelRB.Checked = True Then
        Else
            If specialversion = "Malmgren" Then
                e.Graphics.DrawString("Summa", SmallFontB, BlackFont, 722, y)
            Else
                e.Graphics.DrawString("Belopp", SmallFontB, BlackFont, 722, y)
            End If
        End If

        y = y + 20
        If specialversion = "Malmgren" Then
        Else
            point1 = New Point(30, y) : point2 = New Point(780, y)
            e.Graphics.DrawLine(blackPen, point1, point2) : y = y - 10
        End If
        If TextPositB.Checked = True Then
            If Len(orderanteckning) > 2 And fortext = 0 Then
                y = y + 10
                e.Graphics.DrawString(orderanteckning, SmallFontR, BlackFont, 44, y)
                y = y + 20
                dummy = CountCharacter(orderanteckning, Chr(10))
                For lll = 0 To dummy
                    y = y + 20
                Next
                fortext = 1
            End If

        End If
        If startatex = 1000 Then
            y = 370
            GoTo Extratextkoll
        End If
        Static startAt As Integer = 0
        For i As Integer = startAt To rmax
            pprodnr = DataGridView1.Rows(i).Cells(0).Value
            pprodnamn = DataGridView1.Rows(i).Cells(1).Value
            pEnhet = DataGridView1.Rows(i).Cells(2).Value
            pprisfaktor = DataGridView1.Rows(i).Cells(3).Value
            prantal = DataGridView1.Rows(i).Cells(4).Value
            putpris = DataGridView1.Rows(i).Cells(5).Value
            psumma = DataGridView1.Rows(i).Cells(6).Value
            y = y + radavstand
            'malmgren dubblaprodukter
            If InStr(pprodnr, "+") > 1 Then
                pprodnrb = DataGridView1.Rows(i + 1).Cells(0).Value
                pprodnamnb = DataGridView1.Rows(i + 1).Cells(1).Value
                penhetb = DataGridView1.Rows(i + 1).Cells(2).Value
                pprisfaktorb = DataGridView1.Rows(i + 1).Cells(3).Value
                prantalb = DataGridView1.Rows(i + 1).Cells(4).Value
                putprisb = DataGridView1.Rows(i + 1).Cells(5).Value
                dblprod = pprodnr
                dblprod1 = Strings.Left(dblprod, InStr(dblprod, "+") - 1)
                dblprod2 = Strings.Right(dblprod, Len(dblprod) - InStr(dblprod, "+"))
                e.Graphics.DrawString(dblprod1, SmallFontR, BlackFont, 44, y)


                e.Graphics.DrawString(pprodnamn, SmallFontR, BlackFont, 164, y)


                e.Graphics.DrawString(dblprod2, SmallFontR, BlackFont, 44, y + 16)
                e.Graphics.DrawString(pprodnamnb, SmallFontR, BlackFont, 164, y + 16)
                e.Graphics.DrawString(rightpad(Format$(pprisfaktor, "###0.000")), SmallFontR, BlackFont, 364, y)
                e.Graphics.DrawString(rightpad(Format$(pprisfaktorb, "###0.000")), SmallFontR, BlackFont, 364, y + 16)
                e.Graphics.DrawString(pEnhet, SmallFontR, BlackFont, 440, y)
                e.Graphics.DrawString(penhetb, SmallFontR, BlackFont, 440, y + 16)
                e.Graphics.DrawString(rightpad(Format$(prantal, "###,##0.#")) + " =", SmallFontR, BlackFont, 470, y)
                e.Graphics.DrawString(rightpad(Format$(prantalb, "###,##0.#")) + " =", SmallFontR, BlackFont, 470, y + 16)
                e.Graphics.DrawString(rightpad(Format$(prantal * pprisfaktor, "###,##0.000 kg")), SmallFontR, BlackFont, 550, y)
                e.Graphics.DrawString(rightpad(Format$(prantalb * pprisfaktorb, "###,##0.000 kg")), SmallFontR, BlackFont, 550, y + 16)
                kkilo = (prantal * pprisfaktor) + (prantalb * pprisfaktorb)
                e.Graphics.DrawString(rightpad(Format$(kkilo, "###,##0.000 kg")), SmallFontR, BlackFont, 550, y + 32)
                genomutpris = ((prantalb * pprisfaktorb * putprisb) + (prantal * pprisfaktor * putpris)) / kkilo
                e.Graphics.DrawString(rightpad(Format$(genomutpris, "###,##0.00")), SmallFontR, BlackFont, 610, y + 32)
                e.Graphics.DrawString(rightpad(Format$(genomutpris * kkilo, "###,##0.00")), SmallFontR, BlackFont, 700, y + 32)

                point1 = New Point(400, y + 48) : point2 = New Point(780, y + 48)
                e.Graphics.DrawLine(blackPen, point1, point2)

                i = i + 1
                y = y + 40
                GoTo slutnext
            End If


            e.Graphics.DrawString(pprodnr, SmallFontR, BlackFont, 44, y)
            If specialversion = "Malmgren" Then
                e.Graphics.DrawString(pprodnamn, SmallFontR, BlackFont, 164, y)
            Else
                If Strings.Left(pprodnamn, 4) = "Följ" Then
                    e.Graphics.DrawString(pprodnamn, SmallFontB, BlackFont, 150, y)
                Else
                    e.Graphics.DrawString(pprodnamn, SmallFontR, BlackFont, 150, y)
                End If

            End If
            If orderformat = "MedPrisFaktor" Then
                If pprisfaktor <> 0 Then
                    e.Graphics.DrawString(rightpad(Format$(pprisfaktor, "###0.000")), SmallFontR, BlackFont, 364, y)
                End If
            ElseIf orderformat = "Färg" Then

                If pprisfaktor <> 0 Then
                    e.Graphics.DrawString(rightpad(Format$(pprisfaktor, "###0.#")), SmallFontR, BlackFont, 364, y)
                End If

            End If

            e.Graphics.DrawString(pEnhet, SmallFontR, BlackFont, 440, y)
            If prantal <> 0 Then
                e.Graphics.DrawString(rightpad(Format$(prantal, "###,##0.#")), SmallFontR, BlackFont, 470, y)
            End If

            If specialversion = "Malmgren" And FoljesedelRB.Checked = True Then


            Else

                If putpris <> 0 Then
                    If putpris < 0 Then
                        e.Graphics.DrawString(rightpad(Format$(putpris, "###,##0.00")), SmallFontR, BlackFont, 617, y)
                    Else
                        e.Graphics.DrawString(rightpad(Format$(putpris, "###,##0.00")), SmallFontR, BlackFont, 610, y)
                    End If
                End If
                If psumma <> 0 Then
                    If psumma < 0 Then
                        e.Graphics.DrawString(rightpad(Format$(psumma, "###,##0.00")), SmallFontR, BlackFont, 702, y)
                    Else
                        e.Graphics.DrawString(rightpad(Format$(psumma, "###,##0.00")), SmallFontR, BlackFont, 700, y)
                    End If
                End If

            End If

slutnext:
            e.HasMorePages = False
            ' -----------------------------------------------sidbyte
            If y >= 920 Then
                e.HasMorePages = True
                e.Graphics.DrawString("Sida " + CStr(pagenr), SmallFontR, BlackFont, 40, 1000 - 20)
                pagenr = pagenr + 1
                y = 100
                startAt = i + 1
                Return
            End If

        Next

        y = y + 40


        If CDbl(nullhantering(BetInTB.Text, "T")) <> 0 Then
            '----------------handpenning
            e.Graphics.DrawString("Ordersumma: ", SmallFontR, BlackFont, 570, y)
            e.Graphics.DrawString(rightpad(Format$(CDbl(attbetL.Text), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20

            e.Graphics.DrawString(MOMSLabel.Text, SmallFontR, BlackFont, 570, y)
            e.Graphics.DrawString(rightpad(Format$(CDbl(momsbelL.Text), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 30

            e.Graphics.DrawString("Betald handpenning: ", SmallFontB, BlackFont, 500, y)
            e.Graphics.DrawString(rightpad(Format$(CDbl(BetInTB.Text), "###,##0.00")), SmallFontB, BlackFont, 700, y) : y = y + 20
            handpmoms = BetInTB.Text * 0.2
            e.Graphics.DrawString(MOMSLabel.Text, SmallFontB, BlackFont, 500, y)

            e.Graphics.DrawString(rightpad(Format$(CDbl(handpmoms), "###,##0.00")), SmallFontB, BlackFont, 700, y) : y = y + 30

            e.Graphics.DrawString("Kvar att betala: ", SmallFontR, BlackFont, 500, y)
            e.Graphics.DrawString(rightpad(Format$(CDbl(TillbL.Text), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20

        Else

            If specialversion = "Malmgren" And FoljesedelRB.Checked = True Then
            Else
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
                If FakturaRB.Checked = True Or OffertRB.Checked = True Then
                    If specialversion = "Malmgren" Then


                        If momstyp = 4 Then
                            oresut = (CDbl(OrdsumL.Text) - CDbl(ExrabL.Text) - CDbl(attbetL.Text))
                        Else
                            oresut = (CDbl(OrdsumL.Text) - CDbl(ExrabL.Text) + CDbl(momsbelL.Text)) - CDbl(attbetL.Text)
                        End If



                        e.Graphics.DrawString("Öresutjämning: ", SmallFontR, BlackFont, 570, y)
                        e.Graphics.DrawString(rightpad(Format$(CDbl(oresut * -1), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20

                        e.Graphics.DrawString("Totalt SEK: ", SmallFontR, BlackFont, 570, y)


                    Else
                        e.Graphics.DrawString("Att Betala: ", SmallFontR, BlackFont, 570, y)
                    End If
                    e.Graphics.DrawString(rightpad(Format$(CDbl(attbetL.Text), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20
                ElseIf RakningRB.Checked = True Then
                    e.Graphics.DrawString("Betalat: ", SmallFontR, BlackFont, 570, y)
                    e.Graphics.DrawString(rightpad(Format$(CDbl(attbetL.Text), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20
                ElseIf OrderbekraftRB.Checked = True Then
                    e.Graphics.DrawString("Ordersumma: ", SmallFontR, BlackFont, 570, y)
                    e.Graphics.DrawString(rightpad(Format$(CDbl(attbetL.Text), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20
                Else

                End If
            End If
        End If
Extratextkoll:

        If TextPositB.Checked <> True Then
            If startatex <> 1000 Then
                y = y + 20
                dummy = CountCharacter(orderanteckning, Chr(10))


                If y + (dummy * 20) > 980 Then
                    e.HasMorePages = True
                    e.Graphics.DrawString("Sida " + CStr(pagenr), SmallFontR, BlackFont, 40, 1000 - 20)
                    pagenr = pagenr + 1
                    y = 100
                    startatex = 1000
                    Return
                End If
            End If

            If Len(orderanteckning) > 2 Then

                y = y + 10
                e.Graphics.DrawString(orderanteckning, SmallFontR, BlackFont, 32, y)
                y = y + 30

            End If
        End If

        'fot
        y = 1000
        e.Graphics.DrawString(Copytyp, SmallFontR, BlackFont, 40, y)
        'e.Graphics.DrawString("Sida 1 ", SmallFontR, BlackFont, 40, y)
        If specialversion = "Lodosestad" And nullhantering(MomsTypTB.Text, "T") = 0 Then
            x = CreateGraphics.MeasureString(komment1 + " Omvänd skattskyldighet för byggtjänst gäller", Me.Font).Width
            x = 400 - (x / 2)
            e.Graphics.DrawString(komment1 + " Omvänd skattskyldighet för byggtjänst gäller", SmallFontB, BlackFont, x, y - 20)

        ElseIf OffertRB.Checked = True Then


            x = CreateGraphics.MeasureString(komment1 + "OBS. Vi reserverar oss för framtida prisförändringar", Me.Font).Width
            x = 400 - (x / 2)
            e.Graphics.DrawString(komment1 + "OBS. Vi reserverar oss för framtida prisförändringar", SmallFontB, BlackFont, x, y - 20)

        Else
            x = CreateGraphics.MeasureString(komment1, Me.Font).Width
            x = 400 - (x / 2)

            e.Graphics.DrawString(komment1, SmallFontB, BlackFont, x, y - 20)
        End If
        x = CreateGraphics.MeasureString(komment2, Me.Font).Width
        x = 400 - (x / 2)
        e.Graphics.DrawString(komment2, SmallFontB, BlackFont, x, y)

        e.Graphics.DrawString("Sida " + CStr(pagenr), SmallFontR, BlackFont, 40, y - 20)
        e.Graphics.DrawString("Handl. " + VarrefTB.Text, SmallFontR, BlackFont, 630, y)
        y = y + 20
        ' If fortrycktfot <> True Then
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
        ' End If

        If e.HasMorePages = False Then
            startAt = 0

            pagenr = 1
        End If
        startatex = 0
        fortext = 0

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
        Dim myCmd As New OdbcCommand(mySQL, cn)
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
                falt = falt + "Ordertyp,"
                varden = varden + "'" + "Standard" + "',"
                falt = falt + "Momstyp,"
                varden = varden + "" + CStr(MomsTypTB.Text) + ","

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
                falt = falt + "epost,"
                varden = varden + "'" + nullhantering(kepost, "S") + "',"
                falt = falt + "Varreferens,"
                varden = varden + "'" + VarrefTB.Text + "',"
                falt = falt + "Leveranssatt,"
                varden = varden + "'" + nullhantering(LeveranssattTB.Text, "S") + "',"
                falt = falt + "Erordernummer,"
                varden = varden + "'" + ErOrdernrTB.Text + "',"
                falt = falt + "Markning,"
                varden = varden + "'" + MarkningTB.Text + "',"
                falt = falt + "Extrarab,"
                varden = varden + "" + kommatillpunkt(CStr(nullhantering((exrabTB.Text), "T"))) + ","
                falt = falt + "Levvilkor,"
                varden = varden + "'" + nullhantering(kleveransvilkor, "S") + "',"

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
                If kland <> "Sverige" Then
                    If Len(kland) > 3 Then
                        pkonto = "3105"
                    ElseIf kland = "Norge" Or kland = "NORGE" Then
                        pkonto = "3106"
                    End If
                End If

                varden = varden + "'" + nullhantering(pkonto, "S") + "',"
                
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
        utskriftstyp = ""
        today = Format(Now, "yyyy-MM-dd")
        HistFaktUT = ""
        kundnrTB.Text = ""
        kkundnr = ""
        KundnamnTB.Text = ""
        kkundnamn = ""
        Adress1TB.Text = ""
        kadress1 = ""
        Adress2TB.Text = ""
        kadress2 = ""
        PostnrTB.Text = ""
        kpostnr = ""
        OrtTB.Text = ""
        kort = ""
        LandTB.Text = ""
        kland = ""
        KreditDagarTB.Text = 0
        kkreditdagar = 0
        KontaktmanTB.Text = ""
        kkontaktman = ""
        exrabTB.Text = 0
        extrarab = 0
        MarkningTB.Text = ""
        markning = ""
        LevvilkorTB.Text = ""
        kleveransvilkor = ""
        LeveranssattTB.Text = ""
        kleveranssatt = ""
        LevdatumTB.Text = today
        levdatum = today
        EpostTB.Text = ""
        kepost = ""
        OrdsumL.Text = ""
        momsbelL.Text = ""
        ExrabL.Text = ""
        extrarab = 0
        attbetL.Text = ""
        DataGridView1.Columns.Clear()
        DataGridView1.Rows.Clear()
        DataGridView1.Refresh()

        If KontantRB.Checked = True Then
            makegrid("SimpleKontant")
            kkundnr = "0000"
            kkundnamn = "Kontantförsäljning"
        Else
            makegrid("SimpleKredit")
        End If

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
        orderanteckning = ""
        OrdertypTB.Text = "Standard"
        VarrefTB.Text = defaultref
        MomsTypTB.Text = stdmoms
        kfaktbettypL.Text = ""

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
        Dim dummy As String
        rr = rr + 1
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
            If orderformat = "Färg" Then

                DataGridView1.Item(0, rr - 1).Value = pprodnr
                DataGridView1.Item(5, rr - 1).Value = tillfpris * DataGridView1.Item(3, rr - 1).Value
                DataGridView1.Item(9, rr - 1).Value = tillfpris
                DataGridView1.Item(7, rr - 1).Value = ((putpris1 + (putpris1 * momsv(1))) * DataGridView1.Item(4, rr - 1).Value * DataGridView1.Item(3, rr - 1).Value)
                DataGridView1.Item(8, rr - 1).Value = 100 - ((DataGridView1.Item(5, rr - 1).Value / DataGridView1.Item(7, rr - 1).Value) * 100)
                DataGridView1.Item(6, rr - 1).Value = (DataGridView1.Item(5, rr - 1).Value * DataGridView1.Item(4, rr - 1).Value)
            Else
                DataGridView1.Item(5, rr - 1).Value = tillfpris
                DataGridView1.Item(9, rr - 1).Value = tillfpris
                DataGridView1.Item(7, rr - 1).Value = (putpris1 + (putpris1 * momsv(1))) * DataGridView1.Item(4, rr - 1).Value
                DataGridView1.Item(8, rr - 1).Value = 100 - ((DataGridView1.Item(5, rr - 1).Value / DataGridView1.Item(7, rr - 1).Value) * 100)
                DataGridView1.Item(6, rr - 1).Value = (DataGridView1.Item(3, rr - 1).Value * DataGridView1.Item(5, rr - 1).Value * DataGridView1.Item(4, rr - 1).Value)
            End If
        Else
            tillfpris = (getkundpris(krabattmall, kkundnr, pprodnr))

            tillfpris = tillfpris
            If orderformat = "Färg" Then

                DataGridView1.Item(5, rr - 1).Value = tillfpris * DataGridView1.Item(3, rr - 1).Value
                DataGridView1.Item(9, rr - 1).Value = tillfpris
                DataGridView1.Item(7, rr - 1).Value = (putpris1 * DataGridView1.Item(4, rr - 1).Value * DataGridView1.Item(3, rr - 1).Value)

                DataGridView1.Item(8, rr - 1).Value = 100 - ((DataGridView1.Item(5, rr - 1).Value / DataGridView1.Item(7, rr - 1).Value) * 100)
                DataGridView1.Item(6, rr - 1).Value = (DataGridView1.Item(5, rr - 1).Value * DataGridView1.Item(4, rr - 1).Value)
            Else
                DataGridView1.Item(5, rr - 1).Value = tillfpris
                DataGridView1.Item(9, rr - 1).Value = tillfpris
                DataGridView1.Item(7, rr - 1).Value = putpris1 * DataGridView1.Item(3, rr - 1).Value
                DataGridView1.Item(8, rr - 1).Value = 100 - ((DataGridView1.Item(5, rr - 1).Value / DataGridView1.Item(7, rr - 1).Value) * 100)
                DataGridView1.Item(6, rr - 1).Value = (DataGridView1.Item(3, rr - 1).Value * DataGridView1.Item(5, rr - 1).Value * DataGridView1.Item(4, rr - 1).Value)
            End If
        End If
        If tillfpris <> 0 Then
            DataGridView1.CurrentCell = DataGridView1.Item(4, rr - 1)

        Else
            DataGridView1.CurrentCell = DataGridView1.Item(1, rr - 1)

        End If

    End Sub

   
    Private Sub NyB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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
            If VarrefTB.Text <> "Finns ej" Then kundnrTB.Select()

        End If
    End Sub
    Function hamptalevsatt(ByVal kod As String)
        On Error Resume Next
        Dim cn As OdbcConnection, mySQL As String
        Dim connStr As String, l As Integer
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
        l = 1
        If tabel.HasRows = False Then
            GoTo slut
        Else
            While tabel.Read()
                levsatt(l) = nullhantering(tabel("levsatttext"), "S")
                l = l + 1
            End While

        End If
slut:
        cn.Close()
    End Function

    Private Sub LeveranssattTB_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        LeveranssattTB.Text = hamptalevsatt(LeveranssattTB.Text)
    End Sub

    Function hamptalevsvilkor(ByVal kod As String)
        On Error Resume Next
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
        l = 1
        If tabel.HasRows = False Then
            GoTo slut
        Else
            While tabel.Read()
                levvilk(l) = nullhantering(tabel("levvilktext"), "S")
                l = l + 1
            End While

        End If
slut:
        cn.Close()
    End Function

    Private Sub LevVilkorTB_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        LevvilkorTB.Text = hamptalevsvilkor(LevvilkorTB.Text)
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
            varden = varden + "'" + nullhantering(kkontaktman, "S") + "',"
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
            varden = varden + "'" + nullhantering(levdatum, "S") + "',"
            falt = falt + "Anteckning,"
            varden = varden + "'" + nullhantering(orderanteckning, "S") + "',"
            falt = falt + "Levadress1,"
            varden = varden + "'" + nullhantering(klev1, "S") + "',"
            falt = falt + "Levadress2,"
            varden = varden + "'" + nullhantering(klev2, "S") + "',"
            falt = falt + "Levadress3,"
            varden = varden + "'" + nullhantering(klev3, "S") + "',"

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
            varden = varden + "'" + nullhantering(ordernummer, "S") + "',"
            falt = falt + "Orderdatum,"
            varden = varden + "'" + today + "',"
            falt = falt + "Ordertid,"
            varden = varden + "'" + Mid$(Now, 12, 5) + "',"
            falt = falt + "Erreferens,"
            varden = varden + "'" + nullhantering(kkontaktman, "S") + "',"
            falt = falt + "Kundnamn,"
            varden = varden + "'" + nullhantering(kkundnamn, "S") + "',"
            falt = falt + "Kundnummer,"
            varden = varden + "'" + nullhantering(kkundnr, "S") + "',"
            falt = falt + "Adress1,"
            varden = varden + "'" + nullhantering(kadress1, "S") + "',"
            falt = falt + "Adress2,"
            varden = varden + "'" + nullhantering(kadress2, "S") + "',"
            falt = falt + "Postnummer,"
            varden = varden + "'" + nullhantering(kpostnr, "S") + "',"
            falt = falt + "Ort,"
            varden = varden + "'" + nullhantering(kort, "S") + "',"
            falt = falt + "Land,"
            varden = varden + "'" + nullhantering(kland, "S") + "',"
            falt = falt + "Kreditdagar,"
            varden = varden + "" + CStr(nullhantering(kkreditdagar, "T")) + ","
            falt = falt + "Telefon,"
            varden = varden + "'" + nullhantering(ktelefon, "S") + "',"
            falt = falt + "Telefax,"
            varden = varden + "'" + nullhantering(ktelefax, "S") + "',"
            falt = falt + "epost,"
            varden = varden + "'" + nullhantering(kepost, "S") + "',"
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
            varden = varden + "'" + nullhantering(varref, "S") + "',"
            falt = falt + "Leveranssatt,"
            varden = varden + "'" + nullhantering(kleveranssatt, "S") + "',"
            falt = falt + "Erordernummer,"
            varden = varden + "'" + nullhantering(ertordernr, "S") + "',"
            falt = falt + "Markning,"
            varden = varden + "'" + nullhantering(markning, "S") + "',"
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
            varden = varden + "" + kommatillpunkt(CStr(nullhantering(extrarab, "T"))) + ","
            falt = falt + "Levvilkor,"
            varden = varden + "'" + nullhantering(kleveransvilkor, "S") + "',"
            falt = falt + "Ordning,"
            varden = varden + "" + CStr(i) + ","
            falt = falt + "Ordertyp,"
            varden = varden + "'" + nullhantering(ordertyp, "S") + "',"
            falt = falt + "Momstyp,"
            varden = varden + "" + CStr(nullhantering(momstyp, "T")) + ","
            falt = falt + "Levdat,"
            varden = varden + "'" + nullhantering(levdatum, "S") + "',"
            falt = falt + "Anteckning,"
            varden = varden + "'" + nullhantering(orderanteckning, "S") + "',"
            falt = falt + "Levadress1,"
            varden = varden + "'" + nullhantering(klev1, "S") + "',"
            falt = falt + "Levadress2,"
            varden = varden + "'" + nullhantering(klev2, "S") + "',"
            falt = falt + "Levadress3,"
            varden = varden + "'" + nullhantering(klev3, "S") + "',"
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
            varden = varden + "'" + CStr(nummer) + "',"
            falt = falt + "ordernummer,"
            varden = varden + "'" + nullhantering((OrdnrTB.Text), "S") + "',"
            falt = falt + "Arkivdatum,"
            varden = varden + "'" + today + "',"
            falt = falt + "Arkivid,"
            varden = varden + "'" + Mid$(Now, 12, 5) + "',"
            falt = falt + "Erreferens,"
            varden = varden + "'" + nullhantering(kkontaktman, "S") + "',"
            falt = falt + "Kundnamn,"
            varden = varden + "'" + nullhantering(kkundnamn, "S") + "',"
            falt = falt + "Kundnummer,"
            varden = varden + "'" + nullhantering(kkundnr, "S") + "',"
            falt = falt + "Adress1,"
            varden = varden + "'" + nullhantering(kadress1, "S") + "',"
            falt = falt + "Adress2,"
            varden = varden + "'" + nullhantering(kadress2, "S") + "',"
            falt = falt + "Postnummer,"
            varden = varden + "'" + nullhantering(kpostnr, "S") + "',"
            falt = falt + "Ort,"
            varden = varden + "'" + nullhantering(kort, "S") + "',"
            falt = falt + "Land,"
            varden = varden + "'" + nullhantering(kland, "S") + "',"
            falt = falt + "Kreditdagar,"
            varden = varden + "" + nullhantering(kkreditdagar, "S") + ","
            falt = falt + "Telefon,"
            varden = varden + "'" + nullhantering(ktelefon, "S") + "',"
            falt = falt + "Telefax,"
            varden = varden + "'" + nullhantering(ktelefax, "S") + "',"
            falt = falt + "epost,"
            varden = varden + "'" + nullhantering(kepost, "S") + "',"
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
            If kland <> "Sverige" Then
                If Len(kland) > 3 Then
                    pkonto = "3105"
                ElseIf kland = "Norge" Or kland = "NORGE" Then
                    pkonto = "3106"
                End If
            End If


            varden = varden + "'" + nullhantering(pkonto, "S") + "',"
            falt = falt + "Varreferens,"
            varden = varden + "'" + nullhantering(varref, "S") + "',"
            falt = falt + "Leveranssatt,"
            varden = varden + "'" + nullhantering(kleveranssatt, "S") + "',"
            falt = falt + "Erordernummer,"
            varden = varden + "'" + nullhantering(ertordernr, "S") + "',"
            falt = falt + "Markning,"
            varden = varden + "'" + nullhantering(markning, "S") + "',"
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
            varden = varden + "" + kommatillpunkt(CStr(nullhantering((extrarab), "T"))) + ","
            falt = falt + "Levvilkor,"
            varden = varden + "'" + nullhantering(kleveransvilkor, "S") + "',"
            falt = falt + "Ordning,"
            varden = varden + "" + CStr(i) + ","
            falt = falt + "Ordertyp,"
            varden = varden + "'" + nullhantering(ordertyp, "S") + "',"
            falt = falt + "Momstyp,"
            varden = varden + "" + CStr(nullhantering(momstyp, "T")) + ","
            falt = falt + "Levdat,"
            varden = varden + "'" + nullhantering(levdatum, "S") + "',"
            falt = falt + "Anteckning,"
            varden = varden + "'" + nullhantering(orderanteckning, "S") + "',"
            falt = falt + "Bokfort,"
            varden = varden + "'" + nullhantering("", "S") + "',"
            falt = falt + "Levadress1,"
            varden = varden + "'" + nullhantering(klev1, "S") + "',"
            falt = falt + "Levadress2,"
            varden = varden + "'" + nullhantering(klev2, "S") + "',"
            falt = falt + "Levadress3,"
            varden = varden + "'" + nullhantering(klev3, "S") + "',"
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

    Private Sub Spara_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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
            docnummer = nullhantering(tabel("ordernummer"), "S")
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
            EpostTB.Text = nullhantering(tabel("epost"), "S")
            VarrefTB.Text = nullhantering(tabel("varreferens"), "S")
            orderanteckning = nullhantering(tabel("anteckning"), "S")
            KreditDagarTB.Text = nullhantering(tabel("kreditdagar"), "T")
            DataGridView1.Item(0, n - 1).Value = nullhantering(tabel("Produktnummer"), "S")
            DataGridView1.Item(1, n - 1).Value = nullhantering(tabel("Produktnamn"), "S")

            DataGridView1.Item(2, n - 1).Value = nullhantering(tabel("Enhet"), "S")
            If nullhantering(tabel("Prisfaktor"), "T") = 0 Then GoTo slutloop
            DataGridView1.Item(3, n - 1).Value = nullhantering(tabel("Prisfaktor"), "T")
            DataGridView1.Item(4, n - 1).Value = nullhantering(tabel("antal"), "T")
            klev1 = nullhantering(tabel("levadress1"), "S")
            klev2 = nullhantering(tabel("levadress2"), "S")
            klev3 = nullhantering(tabel("levadress3"), "S")

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
        rr = n - 2 : rmax = rr
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

    Private Sub SammanställningarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


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
            docnummer = nullhantering(tabel("Foljesedelnummer"), "S")
            OrdnrTB.Text = nullhantering(tabel("ordernummer"), "S")


            kundnrTB.Text = nullhantering(tabel("Kundnummer"), "S") : kkundnr = kundnrTB.Text
            KundnamnTB.Text = nullhantering(tabel("Kundnamn"), "S") : kkundnamn = KundnamnTB.Text
            Adress1TB.Text = nullhantering(tabel("adress1"), "S") : kadress1 = Adress1TB.Text
            Adress2TB.Text = nullhantering(tabel("adress2"), "S") : kadress2 = Adress2TB.Text
            PostnrTB.Text = nullhantering(tabel("Postnummer"), "S") : kpostnr = PostnrTB.Text
            OrtTB.Text = nullhantering(tabel("ort"), "S") : kort = OrtTB.Text
            LandTB.Text = nullhantering(tabel("land"), "S") : kland = LandTB.Text
            MomsTypTB.Text = 1
            KontaktmanTB.Text = nullhantering(tabel("erreferens"), "S") : kkontaktman = KontaktmanTB.Text

            EpostTB.Text = nullhantering(tabel("epost"), "S") : kepost = EpostTB.Text
            ktelefon = nullhantering(tabel("telefon"), "S")
            ErOrdernrTB.Text = nullhantering(tabel("erordernummer"), "S")
            MarkningTB.Text = nullhantering(tabel("markning"), "S")
            LevVilkorTB.Text = nullhantering(tabel("levvilkor"), "S")
            LeveranssattTB.Text = nullhantering(tabel("leveranssatt"), "S")
            LevdatumTB.Text = nullhantering(tabel("levdat"), "S")
            EpostTB.Text = nullhantering(tabel("epost"), "S")
            VarrefTB.Text = nullhantering(tabel("varreferens"), "S")
            orderanteckning = nullhantering(tabel("anteckning"), "S")
            klev1 = nullhantering(tabel("levadress1"), "S")
            klev2 = nullhantering(tabel("levadress2"), "S")
            klev3 = nullhantering(tabel("levadress3"), "S")
            KreditDagarTB.Text = nullhantering(tabel("kreditdagar"), "T")
            DataGridView1.Item(0, n - 1).Value = nullhantering(tabel("Produktnummer"), "S")
            DataGridView1.Item(1, n - 1).Value = nullhantering(tabel("Produktnamn"), "S")
            docdatum = nullhantering(tabel("Foljesedeldatum"), "S")
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

        rr = n - 2 : rmax = rr
        DataGridView1.CurrentCell = DataGridView1.Item(0, rmax + 1)
        DataGridView1.RefreshEdit()
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

    

    Private Sub OrdertypTB_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If OrdertypTB.Text = "Standard" Then
            OrdertypTB.Text = "Återkommande"
        ElseIf OrdertypTB.Text = "Återkommande" Then
            OrdertypTB.Text = "Offert"
        Else
            OrdertypTB.Text = "Standard"

        End If



    End Sub
    Private Function HTMLfakturahuvud()
        Dim logout As String = "Jam"
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
        If logout = "Ja" Then
            htmhuv = htmhuv + " <td width='505' colspan='2' height='3'>"
            htmhuv = htmhuv + "<img border='0' src='http://www.vadmin.net/" + KlientID + "logo.jpg' width='240' height='45'></td>"
        Else
            htmhuv = htmhuv + "<td width='505' colspan='2'  height='31'><b><font size='4'>" + Firmanamn + "</font></b></td>"
        End If

        htmhuv = htmhuv + "			<td width='285'  height='31'><b><font size='4'>" + utskriftsort + "</font></b></td>"
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
        htmhuv = htmhuv + "			<td width='505' colspan='2'><font size='2'>Tel: " + ktelefon + "</font></td>"
        htmhuv = htmhuv + "			<td width='285'>&nbsp;</td>"
        htmhuv = htmhuv + "		</tr>"
        htmhuv = htmhuv + "		<tr>"
        htmhuv = htmhuv + "			<td width='505' colspan='2'><font size='2'>Epost: " + kepost + "</font></td>"
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

    Private Sub EpostB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
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
        htmhuv = htmhuv + "		<td width='120' bgcolor='#C0C0C0'><b><font size='2'>Produktnr</font></b></td>"
        htmhuv = htmhuv + "		<td width='314' bgcolor='#C0C0C0'><b><font size='2'>Beskrivning</font></b></td>"
        htmhuv = htmhuv + "		<td align='center' bgcolor='#C0C0C0'><b><font size='2'>Enhet</font></b></td>"
        htmhuv = htmhuv + "		<td align='right' width='64' bgcolor='#C0C0C0'><b><font size='2'>Antal</font></b></td>"
        htmhuv = htmhuv + "		<td align='right' width='94' bgcolor='#C0C0C0'><b><font size='2'>apris</font></b></td>"
        htmhuv = htmhuv + "		<td align='right' width='101' bgcolor='#C0C0C0'><b><font size='2'>Summa</font></b></td>"
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
        htmhuv = htmhuv + "<td width='190'><font size='2'>Handläggare: " + VarrefTB.Text + "</font></td>"
        htmhuv = htmhuv + "</tr>"
        htmhuv = htmhuv + "<tr>"
        htmhuv = htmhuv + "<td width='123' align='center' bgcolor='#C0C0C0'><b>"
        htmhuv = htmhuv + "<font size='2'>Bankgiro</font></b></td>"
        htmhuv = htmhuv + "<td width='144' align='center' bgcolor='#C0C0C0'><b>"
        htmhuv = htmhuv + "<font size='2'>Postgiro</font></b></td>"
        htmhuv = htmhuv + "<td width='325' align='center' bgcolor='#C0C0C0'><b>"
        htmhuv = htmhuv + "<font size='2'>" + banknamn + "</font></b></td>"
        htmhuv = htmhuv + "<td width='190' align='center' bgcolor='#C0C0C0'><b>"
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
                dummy = TillLagerLogg(pprodnr, nummer, (prantal * -1))

                'dummy = hamptaprodukt("LikaNr")
                'dummy = productsaldouppdatering(pprodnr, prantal)
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

    Private Sub DataGridView1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles DataGridView1.KeyDown
        'If e.KeyCode = Keys.F2 Then
        '    MsgBox("f9")
        'End If

        If e.KeyCode = Keys.F2 Then
            franprog = "OrderFakt"


            Me.Visible = False
        End If
        'If e.KeyCode = Keys.F2 Then
        '    franprog = "OrderFakt"
        '    ProduktListaF.Show()
        '    Me.Visible = False

        'End If
    End Sub

    
    Private Sub EpostTB_TextChanged(sender As System.Object, e As System.EventArgs) Handles EpostTB.TextChanged
        kepost = EpostTB.Text
    End Sub



    Private Sub RaderaB_Click(sender As System.Object, e As System.EventArgs)
        suddaorder(OrdnrTB.Text)
        nolla()
    End Sub

    Private Sub LandTB_TextChanged(sender As Object, e As EventArgs) Handles LandTB.TextChanged
        kland = LandTB.Text
    End Sub

   
    Private Sub InfogB_Click(sender As Object, e As EventArgs)
        DataGridView1.Rows.Insert(rr)
    End Sub


    Private Sub MomsTypTB_TextChanged(sender As Object, e As EventArgs)

    End Sub

   


    Private Sub textb_Click(sender As Object, e As EventArgs)

    End Sub

    
    Private Sub LevadressB_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub DagskassaToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub
    Sub TillKassaBoken()
        Dim cn As OdbcConnection, mySQL As String
        Dim connStr As String, falt As String, varden As String
        connStr = "DSN=" + odbcsourcer + "; Database=" + databasnamn + ";Uid=v2000;Pwd=" + odbclosen
        cn = New OdbcConnection(connStr)
        cn.Open()
        falt = "" : varden = ""
        mySQL = "INSERT INTO Kassabok "
        falt = falt + "clientid,"
        varden = varden + "'" + KlientID + "',"
        falt = falt + "Rakningsnummer,"
        varden = varden + "'" + CStr(nummer) + "',"
        falt = falt + "Datum,"
        varden = varden + "'" + today + "',"
        falt = falt + "Tid,"
        varden = varden + "'" + Strings.Right(Now, 8) + "',"
        falt = falt + "Typ,"
        varden = varden + "'" + "Kontant" + "',"
        falt = falt + "VarRef,"
        varden = varden + "'" + VarrefTB.Text + "',"
        falt = falt + "Kontant,"
        varden = varden + "" + kommatillpunkt(CStr(nullhantering(KontantTB.Text, "T"))) + ","
        falt = falt + "Kort,"
        varden = varden + "" + kommatillpunkt(CStr(nullhantering(KortTB.Text, "T"))) + ","
        falt = falt + "Kassa,"
        varden = varden + "'" + stationsid + "',"
        falt = falt + "Belopp,"
        varden = varden + "" + kommatillpunkt(CStr(attbetL.Text - momsbelL.Text)) + ","
        falt = falt + "moms"
        varden = varden + "" + kommatillpunkt(CStr(momsbelL.Text)) + " "
        mySQL = mySQL & "(" & falt & ") VALUES (" & varden & ");"


        Dim myCmd As New OdbcCommand(mySQL, cn)
        myCmd.ExecuteNonQuery()
        'Stop
        cn.Close()
    End Sub

   

    Private Sub LevVilkorTB_KeyUp(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Enter Then
            LeveranssattTB.Select()
        End If
    End Sub

   

    Private Sub KontaktmanTB_KeyUp(sender As Object, e As KeyEventArgs) Handles KontaktmanTB.KeyUp
        If e.KeyCode = Keys.Enter Then
            ErOrdernrTB.Select()
        End If
    End Sub

    Private Sub ErOrdernrTB_KeyUp(sender As Object, e As KeyEventArgs) Handles ErOrdernrTB.KeyUp
        If e.KeyCode = Keys.Enter Then
            MarkningTB.Select()
        End If
    End Sub
   


    Private Sub ErOrdernrTB_TextChanged(sender As Object, e As EventArgs) Handles ErOrdernrTB.TextChanged
        ertordernr = ErOrdernrTB.Text
    End Sub

    Private Sub MarkningTB_KeyUp(sender As Object, e As KeyEventArgs) Handles MarkningTB.KeyUp
        If e.KeyCode = Keys.Enter Then
            LevVilkorTB.Select()
        End If
    End Sub

    Private Sub MarkningTB_TextChanged(sender As Object, e As EventArgs) Handles MarkningTB.TextChanged
        markning = MarkningTB.Text
    End Sub

    Private Sub exrabTB_KeyUp(sender As Object, e As KeyEventArgs) Handles exrabTB.KeyUp
        If e.KeyCode = Keys.Enter Then
            summering()
        End If
    End Sub

    
    Private Sub LeveranssattTB_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub DataGridView1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles DataGridView1.KeyPress
        If System.Windows.Forms.Keys.F2 = True Then

            franprog = "OrderFakt"

            Me.Visible = False
        End If

    End Sub

    Private Sub LevvilkorTB_KeyUp1(sender As Object, e As KeyEventArgs) Handles LevvilkorTB.KeyUp
        If e.KeyCode = Keys.Enter Then
            LeveranssattTB.Select()
        End If
    End Sub

    Private Sub LevvilkorTB_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LevvilkorTB.SelectedIndexChanged
        kleveransvilkor = LevvilkorTB.Text
    End Sub

    Private Sub LeveranssattTB_KeyUp(sender As Object, e As KeyEventArgs) Handles LeveranssattTB.KeyUp
        If e.KeyCode = Keys.Enter Then
            DataGridView1.Focus()
            DataGridView1.CurrentCell = DataGridView1.Item(0, 0)

            SendKeys.Send("{Enter}")
            DataGridView1.CurrentCell.Style.SelectionBackColor = Color.LightBlue
        End If
    End Sub

    Function TillLagerLogg(Artiknr As String, TransID As String, Antal As Double)
        TillLagerLogg = "Ja"
        Dim cn As OdbcConnection, mySQL As String
        Dim connStr As String, falt As String, varden As String
        connStr = "DSN=" + odbcsourcer + "; Database=" + databasnamn + ";Uid=v2000;Pwd=" + odbclosen
        cn = New OdbcConnection(connStr)
        cn.Open()
        falt = "" : varden = ""
        mySQL = "INSERT INTO LagerLogg "
        falt = falt + "clientid,"
        varden = varden + "'" + KlientID + "',"
        falt = falt + "MyTimeStamp,"
        varden = varden + "'" + gettimestamp() + "',"
        falt = falt + "Datum,"
        varden = varden + "'" + today + "',"
        falt = falt + "Id,"
        varden = varden + "'" + TransID + "',"
        falt = falt + "Prodnr,"
        varden = varden + "'" + Artiknr + "',"
        falt = falt + "Antal"
        varden = varden + "" + CStr(Antal) + ""
        mySQL = mySQL & "(" & falt & ") VALUES (" & varden & ");"

        Dim myCmd As New OdbcCommand(mySQL, cn)
        myCmd.ExecuteNonQuery()
        'Stop
        cn.Close()
    End Function

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellLeave(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellLeave
        Dim cellnr As Integer, rownr As Integer
        rownr = DataGridView1.CurrentCellAddress.Y
        cellnr = DataGridView1.CurrentCellAddress.X


    End Sub

    Private Sub LeveranssattTB_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LeveranssattTB.SelectedIndexChanged
        kleveranssatt = LeveranssattTB.Text
    End Sub

    Private Sub VarrefTB_TextChanged(sender As Object, e As EventArgs) Handles VarrefTB.TextChanged
        varref = VarrefTB.Text
    End Sub

    Private Sub MomsTypTB_SelectedIndexChanged(sender As Object, e As EventArgs) Handles MomsTypTB.SelectedIndexChanged
        ' momstyp = CDbl(MomsTypTB.Text)
    End Sub

    Private Sub LevdatumTB_TextChanged(sender As Object, e As EventArgs) Handles LevdatumTB.TextChanged
        levdatum = LevdatumTB.Text
    End Sub

    Private Sub KontaktmanTB_TextChanged(sender As Object, e As EventArgs) Handles KontaktmanTB.TextChanged
        kkontaktman = KontaktmanTB.Text
    End Sub

    Private Sub Adress1TB_TextChanged(sender As Object, e As EventArgs) Handles Adress1TB.TextChanged
        kadress1 = Adress1TB.Text
    End Sub

    Private Sub Adress2TB_TextChanged(sender As Object, e As EventArgs) Handles Adress2TB.TextChanged
        kadress2 = Adress2TB.Text
    End Sub

    Private Sub PostnrTB_TextChanged(sender As Object, e As EventArgs) Handles PostnrTB.TextChanged
        kpostnr = PostnrTB.Text
    End Sub

    Private Sub Huvud_Click(sender As Object, e As EventArgs) Handles Huvud.Click
        Try
            ' Create a new mail message
            Dim mail As New MailMessage()
            mail.From = New MailAddress("emailrelay@vadmin.net") ' Sender's email
            mail.To.Add("edgar.massey@gmail.com")                  ' Recipient's email
            mail.Subject = "Test Email from VB.NET"
            mail.Body = "This is a test email sent from a VB.NET program."

            ' Optional: Add an attachment
            ' mail.Attachments.Add(New Attachment("C:\path\to\file.txt"))

            Dim smtp As New SmtpClient("postal.oderland.com") ' Your SMTP server
            smtp.Port = 587                                ' SMTP port (587 for TLS, 465 for SSL, 25 for unencrypted)
            smtp.Credentials = New NetworkCredential("36780/vadminmail", "Po1e6SFqY5SV3GR97CT1dIqQ") ' Email login
            smtp.EnableSsl = True                          ' Enable SSL for secure communication

            ' Send the email
            smtp.Send(mail)
            Console.WriteLine("Email sent successfully!")
            Stop
        Catch ex As Exception
            Console.WriteLine("Failed to send email: " & ex.Message)
            Stop
        End Try

    End Sub

    Private Sub PrintDocument2_PrintPage(sender As Object, e As Printing.PrintPageEventArgs)
        Dim x As Integer, y As Integer, yy As Integer, point1 As Point, point2 As Point, penhetb As String, prantalb As Double, kkilo As Double, genomutpris As Double
        Dim dblprod As String, dblprod1 As String, dblprod2 As String, pprisfaktorb As Double, putprisb As Double, pprodnrb As String, pprodnamnb As String
        Dim capris As Double
        'väljerskrivare
        PrintDocument1.PrinterSettings.PrinterName = invoicepr
        If HistFaktUT <> "YES" Then

            'getnummer("Kolla", "Nästa fakturanummer")
        End If

        'header
        x = 40 : y = 8
        'If FakturaRB.Checked = True Then
        '    If Val(attbetL.Text) < 0 Then
        '        e.Graphics.DrawString("Kreditfaktura", LargeFontB, BlackFont, 580, y) : y = y + 22
        '    Else
        '        If ProformaCB.Checked = True Then
        '            e.Graphics.DrawString("ProformaFaktura", LargeFontB, BlackFont, 580, y) : y = y + 22
        '        Else

        '            e.Graphics.DrawString("Faktura", LargeFontB, BlackFont, 580, y) : y = y + 22
        '        End If
        '    End If


        '    If OBSL.Text = "Historisk faktura" Then

        '    Else
        '        If ProformaCB.Checked = True Then

        '            nummer = getnummer("Kolla", "Nästa profakturanummer")
        '        Else
        '            nummer = getnummer("Kolla", "Nästa fakturanummer")
        '        End If

        '    End If

        If FoljesedelRB.Checked = True Then
            If faktfolj = 1 Then
                If specialversion = "Målarboden" Then
                    e.Graphics.DrawString("Följesedel", LargeFontB, BlackFont, 630, y) : y = y + 22
                Else
                    e.Graphics.DrawString("Följesedel", LargeFontB, BlackFont, 580, y) : y = y + 22
                End If
                If Len(FoljnrTB.Text) < 1 Then
                    nummer = getnummer("Kolla", "Nästa följesedelnummer")
                    FoljnrTB.Text = nummer
                Else
                    nummer = FoljnrTB.Text
                End If
            Else
                e.Graphics.DrawString("Följesedel", LargeFontB, BlackFont, 580, y) : y = y + 22
                nummer = OrdnrTB.Text

            End If



        End If

        y = 30
        If specialversion = "Målarboden" Then
            e.Graphics.DrawString("Nummer: " + nummer, SmallFontR, BlackFont, 630, y) : y = y + 22
            e.Graphics.DrawString("Kund nr: " + kkundnr, SmallFontR, BlackFont, 630, y) : y = y + 22
            e.Graphics.DrawString(today + "  " + Strings.Mid(Now, 12, 5), SmallFontR, BlackFont, 630, y)
        Else
            e.Graphics.DrawString("Nummer: " + nummer, SmallFontR, BlackFont, 580, y) : y = y + 22
            e.Graphics.DrawString("Kund nr: " + kkundnr, SmallFontR, BlackFont, 580, y) : y = y + 22
            e.Graphics.DrawString("Datum: " + today + "  " + Strings.Mid(Now, 12, 5), SmallFontR, BlackFont, 580, y)
        End If


        If KlientID = "MD" Then
            Dim img As Image = Image.FromFile(sokvag + "MDlogo.jpg")
            e.Graphics.DrawImage(img, 10, 0)
            x = 40 : y = 70
        ElseIf KlientID = "LS" Then
            Dim img As Image = Image.FromFile(sokvag + "LSlogo.jpg")
            e.Graphics.DrawImage(img, 10, 0)
            x = 40 : y = 70
        Else
            x = 40 : y = 8
            e.Graphics.DrawString(Firmanamn, LargeFontB, BlackFont, x, y)
            y = y + 22

        End If


        e.Graphics.DrawString(Postadress1, SmallFontR, BlackFont, x, y) : y = y + 20
        e.Graphics.DrawString(Postadress2, SmallFontR, BlackFont, x, y) : y = y + 20
        e.Graphics.DrawString(Postadress3, SmallFontR, BlackFont, x, y) : y = y + 20
        e.Graphics.DrawString("Telefon: " + telefon, SmallFontR, BlackFont, x, y) : y = y + 20
        e.Graphics.DrawString("Epost: " + epost, SmallFontR, BlackFont, x, y) : y = y + 20
        point1 = New Point(40, y) : point2 = New Point(250, y)
        e.Graphics.DrawLine(blackPen, point1, point2) : y = y - 10



        y = 30

        y = 240
        e.Graphics.DrawString("Leveranssätt: ", SmallFontR, BlackFont, 44, y)
        e.Graphics.DrawString(kleveranssatt, SmallFontR, BlackFont, 164, y) : yy = y
        'e.Graphics.DrawString("Telefonnr.", SmallFontR, BlackFont, 400, y)
        'e.Graphics.DrawString(ktelefon, SmallFontR, BlackFont, 500, y) : 
        y = y + 16

        e.Graphics.DrawString("Leveransvilkor: ", SmallFontR, BlackFont, 44, y)
        e.Graphics.DrawString(kleveransvilkor, SmallFontR, BlackFont, 164, y)
        e.Graphics.DrawString("Er referens: ", SmallFontR, BlackFont, 400, y)
        e.Graphics.DrawString(kkontaktman, SmallFontR, BlackFont, 500, y) : yy = y : y = y + 16

        e.Graphics.DrawString("Betalningsvilkor ", SmallFontR, BlackFont, 44, y)
        e.Graphics.DrawString(CStr(kkreditdagar) + " dagar", SmallFontR, BlackFont, 164, y)

        e.Graphics.DrawString("Ert Ordernr: ", SmallFontR, BlackFont, 400, y)
        e.Graphics.DrawString(ErOrdernrTB.Text, SmallFontR, BlackFont, 500, y) : y = y + 16


        e.Graphics.DrawString("Märkning: ", SmallFontR, BlackFont, 400, y)
        e.Graphics.DrawString(MarkningTB.Text, SmallFontR, BlackFont, 500, y) : y = y + 16

        If OrderbekraftRB.Checked = False And OffertRB.Checked = False Then
            e.Graphics.DrawString("Leveransdatum: ", SmallFontR, BlackFont, 44, y)
            e.Graphics.DrawString(LevdatumTB.Text, SmallFontR, BlackFont, 164, y)
        End If

        'e.Graphics.DrawString("Ert Org nr: ", SmallFontR, BlackFont, 400, y)
        'e.Graphics.DrawString(kregnr, SmallFontR, BlackFont, 500, y) : y = y + 16

        'e.Graphics.DrawString("Epost " + kepost, SmallFontR, BlackFont, 400, y) : y = y + 16


        y = 140
        e.Graphics.DrawString(kkundnamn, OrdFontB, BlackFont, 400, y) : y = y + 20
        e.Graphics.DrawString(kadress1, SmallFontR, BlackFont, 400, y) : y = y + 16
        e.Graphics.DrawString(kadress2, SmallFontR, BlackFont, 400, y) : y = y + 16
        e.Graphics.DrawString(kpostnr + " " + kort, SmallFontR, BlackFont, 400, y) : y = y + 16
        e.Graphics.DrawString(kland, SmallFontR, BlackFont, 400, y) : y = y + 16



        'rubriker
        y = 340
        e.Graphics.DrawString("Prodnr:", SmallFontB, BlackFont, 44, y)


        e.Graphics.DrawString("Prodnamn", SmallFontB, BlackFont, 146, y)



        e.Graphics.DrawString("Enhet", SmallFontB, BlackFont, 440, y)
        e.Graphics.DrawString("Antal", SmallFontB, BlackFont, 520, y)
        If CaPrisRB.Checked = True Then
            e.Graphics.DrawString("Capris", SmallFontB, BlackFont, 650, y)
        Else
            e.Graphics.DrawString("Ertpris", SmallFontB, BlackFont, 650, y)
        End If


        'e.Graphics.DrawString("Belopp", SmallFontB, BlackFont, 722, y)

        y = y + 20

        point1 = New Point(30, y) : point2 = New Point(780, y)
        e.Graphics.DrawLine(blackPen, point1, point2) : y = y - 10

        If TextPositB.Checked = True Then
            If Len(orderanteckning) > 2 Then
                y = y + 10
                e.Graphics.DrawString(orderanteckning, SmallFontR, BlackFont, 32, y)
                y = y + 20

            End If

        End If
        Static startAt As Integer = 0
        For i As Integer = startAt To rmax
            pprodnr = DataGridView1.Rows(i).Cells(0).Value
            pprodnamn = DataGridView1.Rows(i).Cells(1).Value
            pEnhet = DataGridView1.Rows(i).Cells(2).Value
            pprisfaktor = DataGridView1.Rows(i).Cells(3).Value
            prantal = DataGridView1.Rows(i).Cells(4).Value
            putpris = DataGridView1.Rows(i).Cells(5).Value
            psumma = DataGridView1.Rows(i).Cells(6).Value
            capris = DataGridView1.Rows(i).Cells(7).Value
            y = y + 16
            'malmgren dubblaprodukter
            If InStr(pprodnr, "+") > 1 Then
                pprodnrb = DataGridView1.Rows(i + 1).Cells(0).Value
                pprodnamnb = DataGridView1.Rows(i + 1).Cells(1).Value
                penhetb = DataGridView1.Rows(i + 1).Cells(2).Value
                pprisfaktorb = DataGridView1.Rows(i + 1).Cells(3).Value
                prantalb = DataGridView1.Rows(i + 1).Cells(4).Value
                putprisb = DataGridView1.Rows(i + 1).Cells(5).Value
                dblprod = pprodnr
                dblprod1 = Strings.Left(dblprod, InStr(dblprod, "+") - 1)
                dblprod2 = Strings.Right(dblprod, Len(dblprod) - InStr(dblprod, "+"))
                e.Graphics.DrawString(dblprod1, SmallFontR, BlackFont, 44, y)
                e.Graphics.DrawString(pprodnamn, SmallFontR, BlackFont, 164, y)
                e.Graphics.DrawString(dblprod2, SmallFontR, BlackFont, 44, y + 16)
                e.Graphics.DrawString(pprodnamnb, SmallFontR, BlackFont, 164, y + 16)
                e.Graphics.DrawString(rightpad(Format$(pprisfaktor, "###0.000")), SmallFontR, BlackFont, 364, y)
                e.Graphics.DrawString(rightpad(Format$(pprisfaktorb, "###0.000")), SmallFontR, BlackFont, 364, y + 16)
                e.Graphics.DrawString(pEnhet, SmallFontR, BlackFont, 440, y)
                e.Graphics.DrawString(penhetb, SmallFontR, BlackFont, 440, y + 16)
                e.Graphics.DrawString(rightpad(Format$(prantal, "###,##0.#")) + " =", SmallFontR, BlackFont, 470, y)
                e.Graphics.DrawString(rightpad(Format$(prantalb, "###,##0.#")) + " =", SmallFontR, BlackFont, 470, y + 16)
                e.Graphics.DrawString(rightpad(Format$(prantal * pprisfaktor, "###,##0.000 kg")), SmallFontR, BlackFont, 550, y)
                e.Graphics.DrawString(rightpad(Format$(prantalb * pprisfaktorb, "###,##0.000 kg")), SmallFontR, BlackFont, 550, y + 16)
                kkilo = (prantal * pprisfaktor) + (prantalb * pprisfaktorb)
                e.Graphics.DrawString(rightpad(Format$(kkilo, "###,##0.000 kg")), SmallFontR, BlackFont, 550, y + 32)
                'genomutpris = ((prantalb * pprisfaktorb * putprisb) + (prantal * pprisfaktor * putpris)) / kkilo
                'e.Graphics.DrawString(rightpad(Format$(genomutpris, "###,##0.00")), SmallFontR, BlackFont, 610, y + 32)
                'e.Graphics.DrawString(rightpad(Format$(genomutpris * kkilo, "###,##0.00")), SmallFontR, BlackFont, 700, y + 32)

                point1 = New Point(400, y + 48) : point2 = New Point(780, y + 48)
                e.Graphics.DrawLine(blackPen, point1, point2)

                i = i + 1
                y = y + 40
                GoTo slutnext
            End If


            e.Graphics.DrawString(pprodnr, SmallFontR, BlackFont, 44, y)

            e.Graphics.DrawString(pprodnamn, SmallFontR, BlackFont, 150, y)

            If pprisfaktor > 0 Then
                e.Graphics.DrawString(pprisfaktor, SmallFontR, BlackFont, 400, y)
            End If



            e.Graphics.DrawString(pEnhet, SmallFontR, BlackFont, 440, y)
            If prantal <> 0 Then
                e.Graphics.DrawString(rightpad(Format$(prantal, "###,##0.#")), SmallFontR, BlackFont, 470, y)
            End If

            If specialversion <> "Malmgren" Then
                If CaPrisRB.Checked = True Then

                    If capris > 0 Then
                        dummy = hamptaprodukt("LikaNr")

                        If pprodnr = "BEST" Or pprodnr = "ART" Or pprodnr = "ÖVRIGT" Then

                            e.Graphics.DrawString(rightpad(Format$(capris * 0.8, "###,##0.00")), SmallFontR, BlackFont, 617, y)
                        Else
                            e.Graphics.DrawString(rightpad(Format$(putpris1, "###,##0.00")), SmallFontR, BlackFont, 617, y)

                        End If
                        'If MomsTypTB.Text = "4" Then
                        '    e.Graphics.DrawString(rightpad(Format$(capris * 0.8, "###,##0.00")), SmallFontR, BlackFont, 617, y)
                        'Else
                        '    e.Graphics.DrawString(rightpad(Format$(capris, "###,##0.00")), SmallFontR, BlackFont, 617, y)
                        'End If

                    End If
                Else

                    If putpris > 0 Then

                        If hamptaprodukt("LikaNr") = "NO" Then
                            e.Graphics.DrawString(rightpad(Format$(capris * 0.8, "###,##0.00")), SmallFontR, BlackFont, 617, y)
                        Else

                            tillfpris = (getkundpris(krabattmall, kkundnr, pprodnr))
                            e.Graphics.DrawString(rightpad(Format$(tillfpris, "###,##0.00")), SmallFontR, BlackFont, 617, y)

                        End If



                        'tillfpris = (getkundpris(krabattmall, kkundnr, pprodnr))

                        'e.Graphics.DrawString(rightpad(Format$(tillfpris, "###,##0.00")), SmallFontR, BlackFont, 617, y)
                        'If MomsTypTB.Text = "4" Then
                        '    e.Graphics.DrawString(rightpad(Format$((putpris / pprisfaktor) * 0.8, "###,##0.00")), SmallFontR, BlackFont, 617, y)
                        'Else
                        '    e.Graphics.DrawString(rightpad(Format$(putpris / pprisfaktor, "###,##0.00")), SmallFontR, BlackFont, 617, y)
                        'End If
                    End If

                End If


            End If
            'If psumma <> 0 Then
            '        If psumma < 0 Then
            '            e.Graphics.DrawString(rightpad(Format$(psumma, "###,##0.00")), SmallFontR, BlackFont, 702, y)
            '        Else
            '            e.Graphics.DrawString(rightpad(Format$(psumma, "###,##0.00")), SmallFontR, BlackFont, 700, y)
            '        End If
            '    End If



            e.HasMorePages = False
            ' -----------------------------------------------sidbyte
            If y >= 1060 Then
                e.HasMorePages = True
                pagenr = pagenr + 1
                y = 100
                startAt = i + 1
                Return
            End If

slutnext:
        Next

        y = y + 40


        If CDbl(BetInTB.Text) <> 0 Then
            '----------------handpenning
            e.Graphics.DrawString("Ordersumma: ", SmallFontR, BlackFont, 570, y)
            e.Graphics.DrawString(rightpad(Format$(CDbl(attbetL.Text), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20

            e.Graphics.DrawString(MOMSLabel.Text, SmallFontR, BlackFont, 570, y)
            e.Graphics.DrawString(rightpad(Format$(CDbl(momsbelL.Text), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 30

            e.Graphics.DrawString("Betald handpenning: ", SmallFontB, BlackFont, 500, y)
            e.Graphics.DrawString(rightpad(Format$(CDbl(BetInTB.Text), "###,##0.00")), SmallFontB, BlackFont, 700, y) : y = y + 20
            handpmoms = BetInTB.Text * 0.2
            e.Graphics.DrawString(MOMSLabel.Text, SmallFontB, BlackFont, 500, y)

            e.Graphics.DrawString(rightpad(Format$(CDbl(handpmoms), "###,##0.00")), SmallFontB, BlackFont, 700, y) : y = y + 30

            e.Graphics.DrawString("Kvar att betala: ", SmallFontR, BlackFont, 500, y)
            e.Graphics.DrawString(rightpad(Format$(CDbl(TillbL.Text), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20

        Else

            If specialversion = "Malmgren" And FoljesedelRB.Checked = True Then
            Else
                'If CDbl(nullhantering((ExrabL.Text), "T")) <> 0 Then
                '    e.Graphics.DrawString("Ordersumma: ", SmallFontR, BlackFont, 570, y)
                '    e.Graphics.DrawString(rightpad(Format$(CDbl(OrdsumL.Text), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20

                '    e.Graphics.DrawString("Extra rabatt: ", SmallFontR, BlackFont, 570, y)
                '    e.Graphics.DrawString(rightpad(Format$(CDbl(ExrabL.Text), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20
                'End If
                'If CDbl(momsbelL.Text) <> 0 Then
                '    e.Graphics.DrawString("Momspliktig belopp: ", SmallFontR, BlackFont, 570, y)
                '    e.Graphics.DrawString(rightpad(Format$(CDbl(OrdsumL.Text) - CDbl(ExrabL.Text), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20

                '    e.Graphics.DrawString(MOMSLabel.Text, SmallFontR, BlackFont, 570, y)
                '    e.Graphics.DrawString(rightpad(Format$(CDbl(momsbelL.Text), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20
                'End If
                If FakturaRB.Checked = True Then
                    If specialversion = "Malmgren" Then

                        oresut = (CDbl(OrdsumL.Text) - CDbl(ExrabL.Text) + CDbl(momsbelL.Text)) - CDbl(attbetL.Text)


                        e.Graphics.DrawString("Öresutjämning: ", SmallFontR, BlackFont, 570, y)
                        e.Graphics.DrawString(rightpad(Format$(CDbl(oresut * -1), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20

                        e.Graphics.DrawString("Totalt SEK: ", SmallFontR, BlackFont, 570, y)


                    Else
                        e.Graphics.DrawString("Att Betala: ", SmallFontR, BlackFont, 570, y)
                    End If
                    e.Graphics.DrawString(rightpad(Format$(CDbl(attbetL.Text), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20
                ElseIf RakningRB.Checked = True Then
                    e.Graphics.DrawString("Betalat: ", SmallFontR, BlackFont, 570, y)
                    e.Graphics.DrawString(rightpad(Format$(CDbl(attbetL.Text), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20
                ElseIf OrderbekraftRB.Checked = True Then
                    e.Graphics.DrawString("Ordersumma: ", SmallFontR, BlackFont, 570, y)
                    e.Graphics.DrawString(rightpad(Format$(CDbl(attbetL.Text), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20
                Else

                End If
            End If
        End If
        If TextPositB.Checked <> True Then
            If Len(orderanteckning) > 2 Then
                y = y + 10
                e.Graphics.DrawString(orderanteckning, SmallFontR, BlackFont, 32, y)
                y = y + 30

            End If
        End If

        ''fot
        y = 1000


        e.Graphics.DrawString(Copytyp, SmallFontR, BlackFont, 40, y)

        x = CreateGraphics.MeasureString("Mottagit av ________________________________ ", Me.Font).Width
        x = 400 - (x / 2)

        e.Graphics.DrawString("Mottagit av ________________________________ ", SmallFontB, BlackFont, x, y - 20)

        x = CreateGraphics.MeasureString("Priser är exklusiva moms", Me.Font).Width
        x = 400 - (x / 2)
        e.Graphics.DrawString("Priser är exklusiva moms", SmallFontB, BlackFont, x, y)

        e.Graphics.DrawString("Handl. " + VarrefTB.Text, SmallFontR, BlackFont, 630, y)
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


        If e.HasMorePages = False Then
            startAt = 0
        End If


    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick

        tiden = "Proceed"
    End Sub

    Private Sub GetFaktTB_TextChanged(sender As Object, e As EventArgs) Handles GetFaktTB.TextChanged

    End Sub

    Private Sub Ver_Enter(sender As Object, e As EventArgs) Handles Ver.Enter

    End Sub

    Private Sub PrintDocument2_PrintPage_1(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument2.PrintPage
        Dim x As Integer, y As Integer, yy As Integer, point1 As Point, point2 As Point, penhetb As String, prantalb As Double, kkilo As Double, genomutpris As Double
        Dim dblprod As String, dblprod1 As String, dblprod2 As String, pprisfaktorb As Double, putprisb As Double, pprodnrb As String, pprodnamnb As String
        Dim capris As Double
        'väljerskrivare

        If HistFaktUT <> "YES" Then

            'getnummer("Kolla", "Nästa fakturanummer")
        End If

        'header
        x = 40 : y = 8



        If faktfolj = 1 Then
            FoljnrTB.Text = GetFaktTB.Text
            If specialversion = "Målarboden" Then
                e.Graphics.DrawString("Följesedel", LargeFontB, BlackFont, 630, y) : y = y + 22
            Else
                e.Graphics.DrawString("Följesedel", LargeFontB, BlackFont, 580, y) : y = y + 22
            End If
            If Len(FoljnrTB.Text) < 1 Then
                nummer = getnummer("Kolla", "Nästa följesedelnummer")
                FoljnrTB.Text = nummer
            Else
                nummer = GetFaktTB.Text
            End If
        Else
            e.Graphics.DrawString("Följesedel", LargeFontB, BlackFont, 580, y) : y = y + 22
            nummer = GetFaktTB.Text

        End If





        y = 30

        If specialversion = "Målarboden" Then
            e.Graphics.DrawString("Nummer: " + nummer, SmallFontR, BlackFont, 630, y) : y = y + 22
            e.Graphics.DrawString("Kund nr: " + kkundnr, SmallFontR, BlackFont, 630, y) : y = y + 22
            e.Graphics.DrawString(today + "  " + Strings.Mid(Now, 12, 5), SmallFontR, BlackFont, 630, y)
        Else
            e.Graphics.DrawString("Nummer: " + nummer, SmallFontR, BlackFont, 580, y) : y = y + 22
            e.Graphics.DrawString("Kund nr: " + kkundnr, SmallFontR, BlackFont, 580, y) : y = y + 22
            e.Graphics.DrawString("Datum: " + today + "  " + Strings.Mid(Now, 12, 5), SmallFontR, BlackFont, 580, y)
        End If


        If KlientID = "MD" Then
            Dim img As Image = Image.FromFile(sokvag + "MDlogo.jpg")
            e.Graphics.DrawImage(img, 10, 0)
            x = 40 : y = 70
        ElseIf KlientID = "LS" Then
            Dim img As Image = Image.FromFile(sokvag + "LSlogo.jpg")
            e.Graphics.DrawImage(img, 10, 0)
            x = 40 : y = 70
        ElseIf KlientID = "MB" Then
            Dim img As Image = Image.FromFile(sokvag + "MBlogo.jpg")
            e.Graphics.DrawImage(img, 40, 0)
            x = 40 : y = 70
        ElseIf KlientID = "MC" Then
            Dim img As Image = Image.FromFile(sokvag + "MClogo.jpg")
            e.Graphics.DrawImage(img, 40, 0)
            x = 40 : y = 70
        Else
            x = 40 : y = 8
            e.Graphics.DrawString(Firmanamn, LargeFontB, BlackFont, x, y)
            y = y + 22

        End If


        e.Graphics.DrawString(Postadress1, SmallFontR, BlackFont, x, y) : y = y + 20
        e.Graphics.DrawString(Postadress2, SmallFontR, BlackFont, x, y) : y = y + 20
        e.Graphics.DrawString(Postadress3, SmallFontR, BlackFont, x, y) : y = y + 20
        e.Graphics.DrawString("Telefon: " + telefon, SmallFontR, BlackFont, x, y) : y = y + 20
        e.Graphics.DrawString("Epost: " + epost, SmallFontR, BlackFont, x, y) : y = y + 20
        point1 = New Point(40, y) : point2 = New Point(250, y)
        e.Graphics.DrawLine(blackPen, point1, point2) : y = y - 10



        y = 30

        y = 240
        e.Graphics.DrawString("Leveranssätt: ", SmallFontR, BlackFont, 44, y)
        e.Graphics.DrawString(kleveranssatt, SmallFontR, BlackFont, 164, y) : yy = y
        'e.Graphics.DrawString("Telefonnr.", SmallFontR, BlackFont, 400, y)
        'e.Graphics.DrawString(ktelefon, SmallFontR, BlackFont, 500, y) : 
        y = y + 16

        e.Graphics.DrawString("Leveransvilkor: ", SmallFontR, BlackFont, 44, y)
        e.Graphics.DrawString(kleveransvilkor, SmallFontR, BlackFont, 164, y)
        e.Graphics.DrawString("Er referens: ", SmallFontR, BlackFont, 400, y)
        e.Graphics.DrawString(kkontaktman, SmallFontR, BlackFont, 500, y) : yy = y : y = y + 16

        e.Graphics.DrawString("Betalningsvilkor ", SmallFontR, BlackFont, 44, y)
        e.Graphics.DrawString(CStr(kkreditdagar) + " dagar", SmallFontR, BlackFont, 164, y)

        e.Graphics.DrawString("Ert Ordernr: ", SmallFontR, BlackFont, 400, y)
        e.Graphics.DrawString(ErOrdernrTB.Text, SmallFontR, BlackFont, 500, y) : y = y + 16


        e.Graphics.DrawString("Märkning: ", SmallFontR, BlackFont, 400, y)
        e.Graphics.DrawString(MarkningTB.Text, SmallFontR, BlackFont, 500, y) : y = y + 16

        If OrderbekraftRB.Checked = False And OffertRB.Checked = False Then
            e.Graphics.DrawString("Leveransdatum: ", SmallFontR, BlackFont, 44, y)
            e.Graphics.DrawString(LevdatumTB.Text, SmallFontR, BlackFont, 164, y)
        End If

        'e.Graphics.DrawString("Ert Org nr: ", SmallFontR, BlackFont, 400, y)
        'e.Graphics.DrawString(kregnr, SmallFontR, BlackFont, 500, y) : y = y + 16

        'e.Graphics.DrawString("Epost " + kepost, SmallFontR, BlackFont, 400, y) : y = y + 16


        y = 140
        e.Graphics.DrawString(kkundnamn, OrdFontB, BlackFont, 400, y) : y = y + 20
        e.Graphics.DrawString(kadress1, SmallFontR, BlackFont, 400, y) : y = y + 16
        e.Graphics.DrawString(kadress2, SmallFontR, BlackFont, 400, y) : y = y + 16
        e.Graphics.DrawString(kpostnr + " " + kort, SmallFontR, BlackFont, 400, y) : y = y + 16
        e.Graphics.DrawString(kland, SmallFontR, BlackFont, 400, y) : y = y + 16



        'rubriker
        y = 340
        e.Graphics.DrawString("Prodnr:", SmallFontB, BlackFont, 44, y)


        e.Graphics.DrawString("Prodnamn", SmallFontB, BlackFont, 146, y)



        e.Graphics.DrawString("Enhet", SmallFontB, BlackFont, 440, y)
        e.Graphics.DrawString("Antal", SmallFontB, BlackFont, 520, y)
        If specialversion = "Målarboden" Then
            e.Graphics.DrawString("Capris", SmallFontB, BlackFont, 650, y)
        Else
            e.Graphics.DrawString("Ertpris", SmallFontB, BlackFont, 650, y)
        End If


        'e.Graphics.DrawString("Belopp", SmallFontB, BlackFont, 722, y)

        y = y + 20

        point1 = New Point(30, y) : point2 = New Point(780, y)
        e.Graphics.DrawLine(blackPen, point1, point2) : y = y - 10

        If TextPositB.Checked = True Then
            If Len(orderanteckning) > 2 Then
                y = y + 10
                e.Graphics.DrawString(orderanteckning, SmallFontR, BlackFont, 32, y)
                y = y + 20

            End If

        End If
        Static startAt As Integer = 0
        For i As Integer = startAt To rmax
            pprodnr = DataGridView1.Rows(i).Cells(0).Value
            pprodnamn = DataGridView1.Rows(i).Cells(1).Value
            pEnhet = DataGridView1.Rows(i).Cells(2).Value
            pprisfaktor = DataGridView1.Rows(i).Cells(3).Value
            prantal = DataGridView1.Rows(i).Cells(4).Value
            putpris = DataGridView1.Rows(i).Cells(5).Value
            psumma = DataGridView1.Rows(i).Cells(6).Value
            capris = DataGridView1.Rows(i).Cells(7).Value
            y = y + 16
            'malmgren dubblaprodukter



            e.Graphics.DrawString(pprodnr, SmallFontR, BlackFont, 44, y)

            e.Graphics.DrawString(pprodnamn, SmallFontR, BlackFont, 150, y)

            If pprisfaktor > 0 Then
                e.Graphics.DrawString(pprisfaktor, SmallFontR, BlackFont, 400, y)
            End If



            e.Graphics.DrawString(pEnhet, SmallFontR, BlackFont, 440, y)
            If prantal <> 0 Then
                e.Graphics.DrawString(rightpad(Format$(prantal, "###,##0.#")), SmallFontR, BlackFont, 470, y)
            End If

            If specialversion <> "Malmgren" Then
                If specialversion = "Målarboden" Then

                    If capris > 0 Then
                        dummy = hamptaprodukt("LikaNr")

                        If pprodnr = "BEST" Or pprodnr = "ART" Or pprodnr = "ÖVRIGT" Then

                            e.Graphics.DrawString(rightpad(Format$(capris * 0.8, "###,##0.00")), SmallFontR, BlackFont, 617, y)
                        Else
                            e.Graphics.DrawString(rightpad(Format$(putpris1, "###,##0.00")), SmallFontR, BlackFont, 617, y)

                        End If


                    End If
                Else

                    If putpris > 0 Then

                        If hamptaprodukt("LikaNr") = "NO" Then
                            e.Graphics.DrawString(rightpad(Format$(capris * 0.8, "###,##0.00")), SmallFontR, BlackFont, 617, y)
                        Else

                            tillfpris = (getkundpris(krabattmall, kkundnr, pprodnr))
                            e.Graphics.DrawString(rightpad(Format$(tillfpris, "###,##0.00")), SmallFontR, BlackFont, 617, y)

                        End If




                    End If

                End If


            End If




            e.HasMorePages = False
            ' -----------------------------------------------sidbyte
            If y >= 1060 Then
                e.HasMorePages = True
                pagenr = pagenr + 1
                y = 100
                startAt = i + 1
                Return
            End If

slutnext:
        Next

        y = y + 40



        If specialversion = "Malmgren" And FoljesedelRB.Checked = True Then
        Else
            '
            If FakturaRB.Checked = True Then
                If specialversion = "Malmgren" Then

                    oresut = (CDbl(OrdsumL.Text) - CDbl(ExrabL.Text) + CDbl(momsbelL.Text)) - CDbl(attbetL.Text)


                    e.Graphics.DrawString("Öresutjämning: ", SmallFontR, BlackFont, 570, y)
                    e.Graphics.DrawString(rightpad(Format$(CDbl(oresut * -1), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20

                    e.Graphics.DrawString("Totalt SEK: ", SmallFontR, BlackFont, 570, y)


                Else
                    e.Graphics.DrawString("Att Betala: ", SmallFontR, BlackFont, 570, y)
                End If
                e.Graphics.DrawString(rightpad(Format$(CDbl(attbetL.Text), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20
            ElseIf RakningRB.Checked = True Then
                e.Graphics.DrawString("Betalat: ", SmallFontR, BlackFont, 570, y)
                e.Graphics.DrawString(rightpad(Format$(CDbl(attbetL.Text), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20
            ElseIf OrderbekraftRB.Checked = True Then
                e.Graphics.DrawString("Ordersumma: ", SmallFontR, BlackFont, 570, y)
                e.Graphics.DrawString(rightpad(Format$(CDbl(attbetL.Text), "###,##0.00")), SmallFontR, BlackFont, 700, y) : y = y + 20
            Else

            End If
        End If

        If TextPositB.Checked <> True Then
            If Len(orderanteckning) > 2 Then
                y = y + 10
                e.Graphics.DrawString(orderanteckning, SmallFontR, BlackFont, 32, y)
                y = y + 30

            End If
        End If

        ''fot
        y = 1000


        e.Graphics.DrawString(Copytyp, SmallFontR, BlackFont, 40, y)

        x = CreateGraphics.MeasureString("Mottagit av ________________________________ ", Me.Font).Width
        x = 400 - (x / 2)

        e.Graphics.DrawString("Mottagit av ________________________________ ", SmallFontB, BlackFont, x, y - 20)

        x = CreateGraphics.MeasureString("Priser är exklusiva moms", Me.Font).Width
        x = 400 - (x / 2)
        e.Graphics.DrawString("Priser är exklusiva moms", SmallFontB, BlackFont, x, y)

        e.Graphics.DrawString("Handl. " + VarrefTB.Text, SmallFontR, BlackFont, 630, y)
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


        If e.HasMorePages = False Then
            startAt = 0
        End If

    End Sub



    Private Sub OrtTB_TextChanged(sender As Object, e As EventArgs) Handles OrtTB.TextChanged
        kort = OrtTB.Text
    End Sub

    Private Sub OrdertypTB_SelectedIndexChanged(sender As Object, e As EventArgs) Handles OrdertypTB.SelectedIndexChanged
        ordertyp = OrdertypTB.Text
    End Sub

    Private Sub exrabTB_TextChanged(sender As Object, e As EventArgs) Handles exrabTB.TextChanged
        extrarab = nullhantering(exrabTB.Text, "T")
    End Sub


    Private Sub DataGridView1_Enter(sender As Object, e As EventArgs) Handles DataGridView1.Enter
        If Len(kkundnamn) < 2 And KontantRB.Checked = False Then
            MessageBox.Show("Välj kund!!")
            kundnrTB.Select()
        Else
            DataGridView1.CurrentCell.Style.SelectionBackColor = Color.LightBlue
            SendKeys.Send("{Enter}")

        End If

    End Sub

    Private Sub KreditRB_CheckedChanged(sender As Object, e As EventArgs) Handles KreditRB.CheckedChanged

    End Sub

    Private Sub DataGridView1_GotFocus(sender As Object, e As EventArgs) Handles DataGridView1.GotFocus

    End Sub

    Private Sub DataGridView1_EditModeChanged(sender As Object, e As EventArgs) Handles DataGridView1.EditModeChanged

    End Sub

    Function getFAKTURA(ByVal ordnr As String)
        'On Error Resume Next
        Dim n As Integer
        Dim cn As OdbcConnection, mySQL As String
        Dim connStr As String, utpris(10) As Double
        Dim myCmd As OdbcCommand
        Dim tabel As OdbcDataReader
        HistFaktUT = "YES"
     
        connStr = "DSN=" + odbcsourcer + "; Database=" + databasnamn + ";Uid=v2000;Pwd=" + odbclosen
        cn = New OdbcConnection(connStr)
        cn.Open()
        mySQL = "SELECT * FROM orderarkiv "
        mySQL = mySQL + "Where debitid = " & "'" + ordnr + "'"
        mySQL = mySQL + " And clientid = '" + KlientID + "' "
        mySQL = mySQL + " order by mytimestamp "
        myCmd = New OdbcCommand(mySQL, cn)
        tabel = myCmd.ExecuteReader(CommandBehavior.CloseConnection)
        n = 1
        getFAKTURA = "Yes"
        nummer = ordnr


        While tabel.Read()
            faktdatum = nullhantering(tabel("regdatum"), "S")
            OrdnrTB.Text = nullhantering(tabel("ordernummer"), "S")

            docnummer = nullhantering(tabel("debitid"), "S")
            kundnrTB.Text = nullhantering(tabel("Kundnummer"), "S") : kkundnr = kundnrTB.Text
            KundnamnTB.Text = nullhantering(tabel("Kundnamn"), "S") : kkundnamn = KundnamnTB.Text
            Adress1TB.Text = nullhantering(tabel("adress1"), "S") : kadress1 = Adress1TB.Text
            Adress2TB.Text = nullhantering(tabel("adress2"), "S") : kadress2 = Adress2TB.Text
            PostnrTB.Text = nullhantering(tabel("Postnummer"), "S") : kpostnr = PostnrTB.Text
            OrtTB.Text = nullhantering(tabel("ort"), "S") : kort = OrtTB.Text
            LandTB.Text = nullhantering(tabel("land"), "S") : kland = LandTB.Text
            kregnr = nullhantering(tabel("status"), "S")
            MomsTypTB.Text = nullhantering(tabel("momstyp"), "T") : momstyp = MomsTypTB.Text

            KontaktmanTB.Text = nullhantering(tabel("erreferens"), "S") : kkontaktman = KontaktmanTB.Text
            EpostTB.Text = nullhantering(tabel("epost"), "S") : kepost = EpostTB.Text
            ktelefon = nullhantering(tabel("telefon"), "S")
            OrdertypTB.Text = nullhantering(tabel("ordertyp"), "S")
            ErOrdernrTB.Text = nullhantering(tabel("erordernummer"), "S")
            MarkningTB.Text = nullhantering(tabel("markning"), "S")
            LevvilkorTB.Text = nullhantering(tabel("levvilkor"), "S")
            LeveranssattTB.Text = nullhantering(tabel("leveranssatt"), "S")
            LevdatumTB.Text = nullhantering(tabel("levdat"), "S")
            EpostTB.Text = nullhantering(tabel("epost"), "S")
            VarrefTB.Text = nullhantering(tabel("varreferens"), "S")
            orderanteckning = nullhantering(tabel("anteckning"), "S")
            KreditDagarTB.Text = nullhantering(tabel("kreditdagar"), "T")
            kkreditdagar = nullhantering(tabel("kreditdagar"), "T")
            DataGridView1.Item(0, n - 1).Value = nullhantering(tabel("Produktnummer"), "S")
            DataGridView1.Item(1, n - 1).Value = nullhantering(tabel("Produktnamn"), "S")
            DataGridView1.Item(2, n - 1).Value = nullhantering(tabel("Enhet"), "S")
            If nullhantering(tabel("Prisfaktor"), "T") = 0 Then GoTo slutloop
            DataGridView1.Item(3, n - 1).Value = nullhantering(tabel("Prisfaktor"), "T")
            DataGridView1.Item(4, n - 1).Value = nullhantering(tabel("antal"), "T")
            klev1 = nullhantering(tabel("levadress1"), "S")
            klev2 = nullhantering(tabel("levadress2"), "S")
            klev3 = nullhantering(tabel("levadress3"), "S")

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

        rr = n - 2 : rmax = rr
        DataGridView1.CurrentCell = DataGridView1.Item(0, rmax + 1)
        DataGridView1.RefreshEdit()
        summering()
        cn.Close()

    End Function
    Function getfoljesedelmulti(ByVal antal As Integer)
        'On Error Resume Next
        Dim n As Integer, ll As Integer, firstrow As String
        Dim cn As OdbcConnection, mySQL As String
        Dim connStr As String, utpris(10) As Double
        Dim myCmd As OdbcCommand
        Dim tabel As OdbcDataReader
        n = 1
        firstrow = ""
        For ll = 0 To antal
            connStr = "DSN=" + odbcsourcer + "; Database=" + databasnamn + ";Uid=v2000;Pwd=" + odbclosen
            cn = New OdbcConnection(connStr)
            cn.Open()
            mySQL = "SELECT * FROM foljesedlar "
            mySQL = mySQL + "Where foljesedelnummer = " & "'" + foljnr(ll) + "'"
            mySQL = mySQL + " And clientid = '" + KlientID + "' "
            mySQL = mySQL + " order by mytimestamp "
            myCmd = New OdbcCommand(mySQL, cn)
            tabel = myCmd.ExecuteReader(CommandBehavior.CloseConnection)

            getfoljesedelmulti = "Yes"
            firstrow = ""
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
                LevvilkorTB.Text = nullhantering(tabel("levvilkor"), "S")
                LeveranssattTB.Text = nullhantering(tabel("leveranssatt"), "S")
                LevdatumTB.Text = nullhantering(tabel("levdat"), "S")
                EpostTB.Text = nullhantering(tabel("epost"), "S")
                If Len(firstrow) < 1 Then
                    DataGridView1.Item(1, n - 1).Value = "_________________________________________________________________________"
                    n = n + 1
                    DataGridView1.Item(1, n - 1).Value = "Följ.nr:" + nullhantering(tabel("Foljesedelnummer"), "S") + "  Datum:" + nullhantering(tabel("Foljesedeldatum"), "S") + " " + nullhantering(tabel("Erordernummer"), "S") + " " + nullhantering(tabel("Markning"), "S")
                    n = n + 1
                   
                    firstrow = "Done"
                End If
                VarrefTB.Text = nullhantering(tabel("varreferens"), "S")
                orderanteckning = nullhantering(tabel("anteckning"), "S")
                klev1 = nullhantering(tabel("levadress1"), "S")
                klev2 = nullhantering(tabel("levadress2"), "S")
                klev3 = nullhantering(tabel("levadress3"), "S")
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
                n = n + 1 : rmax = n
            End While
        Next
        getfoljesedelmulti = ""
        summering()
        ' cn.Close()
    End Function

   
    Private Sub Gnamn_Click(sender As Object, e As EventArgs)

        PDFCB.Enabled = True
    End Sub
    Sub tillpdfrutin()
        Dim epostbody As String


        epostbody = "En pdf faktura med Fakturanr " + nummer + " skapas och skickas till " + EpostTB.Text + " : Från " + " (" + KlientID + ") "

        skickatillepostrelay("PDF@vadmincloud.com", epost, "PDFfil skapas", epostbody, Firmanamn)

        skickatillepostrelay("PDF@vadmincloud.com", "edgar.massey@telia.com", "PDFfil skapas", epostbody, Firmanamn)
        'MsgBox("Epost skickat!!")

    End Sub

    

    Private Sub KlientIdl_Click(sender As Object, e As EventArgs)
        pdfval = "YES"

    End Sub

    Private Sub GetfakturaB_Click(sender As Object, e As EventArgs) Handles GetfakturaB.Click
        skickapdf()
    End Sub

    Private Sub KlientIDB_TextChanged(sender As Object, e As EventArgs) Handles KlientIDB.TextChanged
        KlientID = KlientIDB.Text

    End Sub

    Function hamptaforetag(ByVal clientid As String)
        getfiles()

        odbcsource = "VadminODBC"
        odbcsourcer = odbcsource


        ' On Error GoTo nocon
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


    Function CountCharacter(value As String, ch As String) As Integer

        CountCharacter = 0

        While InStr(value, ch) > 0
            value = Strings.Right(value, (Len(value) - InStr(value, ch)))
            CountCharacter = CountCharacter + 1

        End While

    End Function





    Function skickaepost(till As String)



        Dim pdfname
        till = BLANKBORT(till)
        skickaepost = "YES"

        'SMTP Credentials
        SMTPUser = "36780/vadminmail"
        SMTPPassW = "Po1e6SFqY5SV3GR97CT1dIqQ"
        SMTPservernamn = "postal.oderland.com"

        Dim epostsubject = doctyp + "# " + GetFaktTB.Text + " från " + Strings.Left(Firmanamn, 10)
        Dim epostfran = "epostrelay@vadmin.net"
        Dim epostbody = doctyp + " nr: " + GetFaktTB.Text + " i pdf format"

        Try
            Dim SmtpServer As New SmtpClient()
            Dim mail As New MailMessage()
            SmtpServer.Host = SMTPservernamn
            SmtpServer.Credentials = New Net.NetworkCredential(SMTPUser, SMTPPassW)
            SmtpServer.Port = 587
            SmtpServer.EnableSsl = True
            mail.From = New MailAddress(epostfran)

            ' Attach PDF
            ' pdfname = getpdfname()
            pdfname = "C:\PDF\" & GetFaktTB.Text & ".pdf"
            mail.Attachments.Add(New Attachment(pdfname))

            ' Add primary recipient (To)

            mail.To.Add(till)
            mail.CC.Add(BLANKBORT(epost))

            ' Add BCC (Hidden email)
            Dim bccEmail As String = "epostarkiv25@gmail.com"
            mail.Bcc.Add(bccEmail)


            ' Set email subject and body
            mail.Subject = epostsubject
            mail.Body = epostbody
            mail.IsBodyHtml = True

            ' Debugging output
            Debug.WriteLine("To: " & String.Join(",", mail.To.Select(Function(m) m.Address)))
            Debug.WriteLine("CC: " & String.Join(",", mail.CC.Select(Function(m) m.Address)))
            Debug.WriteLine("BCC: " & String.Join(",", mail.Bcc.Select(Function(m) m.Address)))

            ' Send Email
            System.Threading.Thread.Sleep(2000)
            SmtpServer.Send(mail)
            System.Threading.Thread.Sleep(2000)

        Catch ex As Exception
            MsgBox("Error while sending email: " & ex.ToString())
        End Try
    End Function
    Function getfiles()
        Dim fileArray() As String, test As String
        test = ""
        ' fileArray = Directory.GetFiles(System.AppDomain.CurrentDomain.BaseDirectory & "\pdf\")
        fileArray = Directory.GetFiles("C:\pdf\")
        Array.Sort(fileArray)
        For Each test In fileArray
            ' do stuff with item
        Next
        While Strings.InStr(test, "\") > 0
            test = Strings.Right(test, Len(test) - Strings.InStr(test, "\"))
        End While
        getfiles = test
    End Function

    Sub getpdfinfo()
        Dim cn As OdbcConnection, mySQL As String
        Dim connStr As String, firstbort As String
        epostfrom = "" : eposttill = "" : epostsubject = "" : epostbody = ""
        connStr = "DSN=" + "VadminODBC" + "; Database=" + "emailrelay" + ";Uid=v2000;Pwd=" + "alfons"
        cn = New OdbcConnection(connStr)
        cn.Open()
        doctyp = ""
        mySQL = "SELECT top (1) *  FROM inkommandePDF  order by regdatum"
        Dim myCmd As New OdbcCommand(mySQL, cn)
        Dim tabel As OdbcDataReader = myCmd.ExecuteReader(CommandBehavior.CloseConnection)
        firstbort = ""
        If tabel.Read = True Then
            mytimestamp = nullhantering(tabel("mytimestamp"), "S")
            fromsystem = nullhantering(tabel("fromsystem"), "S")
            KlientIDB.Text = fromsystem
            'KlientID = fromsystem
            regdatum = nullhantering(tabel("regdatum"), "S")
            databasnamn = nullhantering(tabel("epostfrom"), "S")
            epostfrom = databasnamn
            eposttill = BLANKBORT(nullhantering(tabel("eposttill"), "S"))
            KundepostTB.Text = eposttill
            epostsubject = nullhantering(tabel("epostsubject"), "S")
            GetFaktTB.Text = epostsubject
            epostbody = nullhantering(tabel("epostbody"), "S")
            doctyp = epostbody
            If epostbody = "Faktura" Then
                FakturaRB.Checked = True
            ElseIf epostbody = "Följesedel" Then
                FoljesedelRB.Checked = True
            ElseIf epostbody = "Orderbekräftelse" Then
                OrderbekraftRB.Checked = True
            End If


            sparatillpdfarkiv()
            cn.Close()
            cn = New OdbcConnection(connStr)
            cn.Open()
            mySQL = "DELETE FROM inkommandePDF "
            mySQL = mySQL + " Where mytimestamp = '" + mytimestamp + "' "
            myCmd = New OdbcCommand(mySQL, cn)
            myCmd.ExecuteNonQuery()
            cn.Close()

        End If
    End Sub
    Sub sparatillpdfarkiv()
        Dim cn As OdbcConnection, mySQL As String
        Dim connStr As String, falt As String, varden As String
        connStr = "DSN=" + odbcsourcer + "; Database=" + "emailrelay" + ";Uid=v2000;Pwd=" + odbclosen
        cn = New OdbcConnection(connStr)
        cn.Open()
        epostfrom = "PDF@vadmincloud.com"
        falt = "" : varden = ""
        mySQL = "INSERT INTO pdfarkiv "
        falt = falt + "Mytimestamp,"
        varden = varden + "'" + gettimestamp() + "',"
        falt = falt + "fromsystem,"
        varden = varden + "'" + fromsystem + "',"
        falt = falt + "Regdatum,"
        varden = varden + "'" + CStr(Now) + "',"
        falt = falt + "epostfrom,"
        varden = varden + "'" + epostfrom + "',"
        falt = falt + "eposttill,"
        varden = varden + "'" + eposttill + "',"
        falt = falt + "epostsubject,"
        varden = varden + "'" + epostsubject + "',"
        falt = falt + "epostbody"
        varden = varden + "'" + epostbody + "'"
        mySQL = mySQL & "(" & falt & ") VALUES (" & varden & ");"
        Dim myCmd As New OdbcCommand(mySQL, cn)
        myCmd.ExecuteNonQuery()
        cn.Close()
    End Sub

    Private Sub TimerB_Click(sender As Object, e As EventArgs) Handles TimerB.Click
        If TimerB.Text = "Start timern" Then

            Timer1.Interval = 120000  '2 minute
            Timer1.Start()
            TimerB.Text = "Timer enabled"
        Else
            

            Timer1.Stop()
            TimerB.Text = "Start timern"
        End If
    End Sub
    Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick


        ticks = ticks + 1
        TicksL.Text = CStr(ticks) + " - " + CStr(antpdf)

        If FardigStatus <> "Pagar" Then
            skickapdf()
        End If
    End Sub
    Sub skickapdf()
        On Error Resume Next
        Dim till As String = epost
        Dim kundepost As String = ""
        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        FardigStatus = "Pagar"
        My.Computer.FileSystem.CopyDirectory("C:\PDF", "C:\PDFArkiv", True)
        today = Format(Now, "yyyy-MM-dd")
        odbclosen = "alfons"

        getpdfinfo()


        If epostfrom <> "" Then
            Kill("C:\PDF\*.*")
            wait(10)

            hamptaforetag(KlientID)
            If epostbody = "Faktura" Then
                getFAKTURA(GetFaktTB.Text)
            End If

            If epostbody = "Följesedel" Then
                getfoljesedel(GetFaktTB.Text)
            End If

            If epostbody = "Orderbekräftelse" Then
                getorder(GetFaktTB.Text)
            End If
            If epostbody = "Offert" Then
                OffertRB.Checked = True
                getorder(GetFaktTB.Text)
            End If

            antpdf = antpdf + 1


            If epostbody = "Följesedel" Then
                'PrintDocument2.PrinterSettings.Copies = 1
                'PrintDocument2.PrinterSettings.PrinterName = pdfprintername
                'PrintDocument2.Print()
                PrintDocument2.PrinterSettings.PrinterName = "Microsoft Print to PDF"
                PrintDocument2.PrinterSettings.PrintToFile = True
                PrintDocument2.PrinterSettings.PrintFileName = "C:\PDF\" & GetFaktTB.Text & ".pdf"
                PrintDocument2.Print()

            Else
                'PrintDocument1.PrinterSettings.Copies = 1
                'PrintDocument1.PrinterSettings.PrinterName = pdfprintername
                'PrintDocument1.Print()

                PrintDocument1.PrinterSettings.PrinterName = "Microsoft Print to PDF"
                PrintDocument1.PrinterSettings.PrintToFile = True
                PrintDocument1.PrinterSettings.PrintFileName = "C:\PDF\" & GetFaktTB.Text & ".pdf"
                PrintDocument1.Print()
            End If


            wait(45)
                If PDFKontrolCB.Checked = True Then
                    Dim Message As String = "Godkänns PDFfilen ?"
                    Dim Caption As String = "Kontrollera"
                    Dim Buttons As MessageBoxButtons = MessageBoxButtons.OKCancel
                    Dim Result As DialogResult

                    Result = MessageBox.Show(Message, Caption, Buttons)

                    If Result <> 1 Then
                        End
                    End If
                End If

                If Strings.InStr(eposttill, "@") > 1 Then
                    skickaepost(eposttill)
                End If
                nolla()


                KlientIDB.Text = ""
                KundepostTB.Text = ""
                GetFaktTB.Text = ""



                'Dim fileArray() As String, test As String
                'test = ""
                'fileArray = Directory.GetFiles("C:\pdf\")

                'For Each test In fileArray

                '    My.Computer.FileSystem.DeleteFile(test, FileIO.UIOption.AllDialogs, FileIO.RecycleOption.DeletePermanently)
                'Next

                wait(10)
                My.Computer.FileSystem.CopyDirectory("C:\PDF", "C:\PDFArkiv", True)

            ' Kill("C:\PDF\*.*")
        End If
            Me.BringToFront()
        Me.Cursor = System.Windows.Forms.Cursors.Arrow
        FardigStatus = "Färdig"
    End Sub
    Function getpdfname() As String
        Dim Files() As String, pdfname As String = "", nr As Integer = 0
        Files = System.IO.Directory.GetFiles("C:\pdf\")
        Dim filePaths As Linq.IOrderedEnumerable(Of IO.FileInfo) =
  New DirectoryInfo("c:\pdf\").GetFiles().OrderBy(Function(f As FileInfo) f.CreationTime)
        For Each fi As IO.FileInfo In filePaths
            pdfname = Files(nr)
            nr = nr + 1
        Next
        getpdfname = pdfname
    End Function
    Public Sub wait(ByVal seconds As Single)
        Static start As Single
        start = VB.Timer()
        Do While VB.Timer() < start + seconds
            System.Windows.Forms.Application.DoEvents()
        Loop
    End Sub
End Class
