﻿Imports System.Math
Imports System
Imports System.Globalization
Imports System.Threading
Imports Word = Microsoft.Office.Interop.Word
Imports System.Runtime.InteropServices

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, NumericUpDown9.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown1.ValueChanged, MyBase.Load, NumericUpDown3.ValueChanged, NumericUpDown12.ValueChanged, NumericUpDown11.ValueChanged
        Dim power, moment_motor, speed, factor As Double
        Dim OD_as, OD_naaf, ID_naaf, ring_dikte As Double
        Dim lengte_as, Coeffie_slip, moment_slip, p_vlaktedruk, trekspanning_ring, combi_spanning As Double
        Dim Elast_mod, sd_verhouding, s_maat_mu As Double
        Dim uitzetting, delta_temp, CoeffI_term As Double
        Dim F_pers, S_verlies As Double

        power = NumericUpDown6.Value * 1000
        speed = NumericUpDown5.Value
        factor = NumericUpDown7.Value

        OD_as = NumericUpDown1.Value
        OD_naaf = NumericUpDown4.Value
        ID_naaf = NumericUpDown1.Value                      'ID_naaf is praktisch OD_as
        lengte_as = NumericUpDown2.Value                    'Insteeklengte as
        Coeffie_slip = NumericUpDown3.Value                 'Coefficient slip
        ring_dikte = (OD_naaf - ID_naaf) / 2
        Elast_mod = NumericUpDown8.Value

        delta_temp = NumericUpDown9.Value
        CoeffI_term = NumericUpDown10.Value

        '----------- rekenen -----------
        Try
            moment_motor = power * factor / (speed * PI / 30)

            '----- as -----------------------
            p_vlaktedruk = 2 * moment_motor * 1000 / (PI * OD_as ^ 2 * lengte_as * Coeffie_slip)    '[N/mm]
            moment_slip = p_vlaktedruk * PI * OD_as * lengte_as * Coeffie_slip * 0.5 * OD_as / 1000 '[N.m]


            '----- ring ---------------------
            trekspanning_ring = p_vlaktedruk * (OD_naaf ^ 2 + ID_naaf ^ 2) / (OD_naaf ^ 2 - ID_naaf ^ 2)

            '----- s/d ---------------------
            sd_verhouding = 2 * p_vlaktedruk * OD_naaf ^ 2 / (Elast_mod * (OD_naaf ^ 2 - ID_naaf ^ 2))


            '----- Uitzetting --
            uitzetting = delta_temp * CoeffI_term * ID_naaf * 1000   '[mu]
            s_maat_mu = sd_verhouding * OD_as * 1000

            '------ Gecombineerde spanning ------------------ 
            combi_spanning = Sqrt(trekspanning_ring ^ 2 + p_vlaktedruk ^ 2 + trekspanning_ring * p_vlaktedruk)

            '----- Perskracht --
            F_pers = p_vlaktedruk * Coeffie_slip * PI * OD_as * lengte_as   '[N]
            F_pers /= 10000                                                 '[N-> ton

            '----- Oppervlakte ruwheid --------
            S_verlies = Round(1.2 * 4 * (NumericUpDown11.Value + NumericUpDown12.Value), 0)        '60% verlies 

            '----- Presenteren --------------
            TextBox1.Text = Round(moment_motor, 0).ToString
            TextBox2.Text = Round(uitzetting, 0).ToString               'Thermische expansie
            TextBox3.Text = Round(ring_dikte, 1).ToString
            TextBox4.Text = Round(p_vlaktedruk, 1).ToString             'Radiale spanning = vlaktedruk
            TextBox5.Text = Round(p_vlaktedruk, 1).ToString             'Vlaktedruk as
            TextBox7.Text = Round(trekspanning_ring, 1).ToString        'Trekspanning ring
            TextBox6.Text = Round(sd_verhouding, 4).ToString            's/d 
            TextBox8.Text = Round(1 / sd_verhouding, 0).ToString        'd/s
            TextBox9.Text = Round(s_maat_mu, 0).ToString                's_maat
            TextBox10.Text = Round(combi_spanning, 0).ToString           'gecombineerde spanning naaf

            TextBox11.Text = Round(F_pers, 1).ToString                  'Perskracht [ton]
            TextBox14.Text = Round(S_verlies, 1).ToString               'Verlies door oppervlakte ruwheid [mu]
            TextBox15.Text = Round(moment_slip, 0).ToString             'As begint te slippen [Nm]

            If p_vlaktedruk < 90 Then           'Check vlakte druk
                TextBox5.BackColor = Color.LightGreen
            Else
                TextBox5.BackColor = Color.Red
            End If

            If 1 / sd_verhouding > 750 Then     'Check krimpmaat
                TextBox6.BackColor = Color.LightGreen
                TextBox8.BackColor = Color.LightGreen
                TextBox9.BackColor = Color.LightGreen
            Else
                TextBox6.BackColor = Color.Red
                TextBox8.BackColor = Color.Red
                TextBox9.BackColor = Color.Red
            End If

            If NumericUpDown1.Value < NumericUpDown4.Value - 20 Then      'Onmogelijke naaf dimensie
                NumericUpDown4.BackColor = SystemColors.Window
            Else
                NumericUpDown4.BackColor = Color.Red
            End If

        Catch
            MessageBox.Show("Exception")
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1, oPara2, oPara3 As Word.Paragraph

        'Start Word and open the document template. 
        oWord = CreateObject("Word.Application")
        oWord.Visible = True
        oDoc = oWord.Documents.Add


        'Insert a paragraph at the beginning of the document. 
        oPara1 = oDoc.Content.Paragraphs.Add
        oPara1.Range.Text = "VTK Engineering department"
        oPara1.Range.Font.Name = "Arial"
        oPara1.Range.Font.Size = 16
        oPara1.Range.Font.Bold = True
        oPara1.Format.SpaceAfter = 4                '24 pt spacing after paragraph. 
        oPara1.Range.InsertParagraphAfter()

        oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara2.Range.Font.Size = 11
        oPara2.Format.SpaceAfter = 2
        oPara2.Range.Font.Bold = False
        oPara2.Range.Text = "Berekening krimpen en persen van as en naaf" & vbCrLf
        oPara2.Range.InsertParagraphAfter()

        '----------------------------------------------
        'Insert a table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 4, 2)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = 11
        oTable.Range.Font.Bold = False
        oTable.Rows.Item(1).Range.Font.Bold = True

        oTable.Cell(1, 1).Range.Text = "Project Name"
        oTable.Cell(1, 2).Range.Text = TextBox16.Text
        oTable.Cell(2, 1).Range.Text = "Project number "
        oTable.Cell(2, 2).Range.Text = TextBox17.Text
        oTable.Cell(3, 1).Range.Text = "Author "
        oTable.Cell(3, 2).Range.Text = Environment.UserName
        oTable.Cell(4, 1).Range.Text = "Date "
        oTable.Cell(4, 2).Range.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        oTable.Columns.Item(1).Width = oWord.InchesToPoints(2.5)   'Change width of columns 1 & 2.
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(2)
        oTable.Rows.Item(1).Range.Font.Bold = True
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        '----------------------------------------------
        'Insert a 14 x 5 table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 13, 3)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = 10
        oTable.Range.Font.Bold = False
        oTable.Rows.Item(1).Range.Font.Bold = True

        oTable.Cell(1, 1).Range.Text = "Input Data"
        oTable.Cell(1, 2).Range.Text = ""
        oTable.Cell(1, 3).Range.Text = ""

        oTable.Cell(2, 1).Range.Text = "Diameter as (d_a)"
        oTable.Cell(2, 2).Range.Text = NumericUpDown1.Value
        oTable.Cell(2, 3).Range.Text = "[mm]"

        oTable.Cell(3, 1).Range.Text = "Insteek lengte as (l)"
        oTable.Cell(3, 2).Range.Text = NumericUpDown2.Value
        oTable.Cell(3, 3).Range.Text = "[mm]"

        oTable.Cell(4, 1).Range.Text = "Buiten diameter naaf (D)"
        oTable.Cell(4, 2).Range.Text = NumericUpDown4.Value
        oTable.Cell(4, 3).Range.Text = "[mm]"

        oTable.Cell(5, 1).Range.Text = "Motor vermogen"
        oTable.Cell(5, 2).Range.Text = NumericUpDown6.Value
        oTable.Cell(5, 3).Range.Text = "[Kw]"


        oTable.Cell(6, 1).Range.Text = "Toerental"
        oTable.Cell(6, 2).Range.Text = NumericUpDown5.Value
        oTable.Cell(6, 3).Range.Text = "[rpm]"


        '---- -----
        oTable.Cell(7, 1).Range.Text = "Bedrijfstoeslag factor"
        oTable.Cell(7, 2).Range.Text = NumericUpDown7.Value
        oTable.Cell(7, 3).Range.Text = "[-]"

        oTable.Cell(8, 1).Range.Text = "Frictie coefficient"
        oTable.Cell(8, 2).Range.Text = NumericUpDown3.Value
        oTable.Cell(8, 3).Range.Text = "[-]"

        oTable.Cell(9, 1).Range.Text = "E modulus"
        oTable.Cell(9, 2).Range.Text = NumericUpDown8.Value
        oTable.Cell(9, 3).Range.Text = "[N/mm2]"

        '---- -----
        oTable.Cell(10, 1).Range.Text = "Thermal expansion coefficient"
        oTable.Cell(10, 2).Range.Text = NumericUpDown10.Value
        oTable.Cell(10, 3).Range.Text = "[mm/mm.K]"

        oTable.Cell(11, 1).Range.Text = "Opwarming"
        oTable.Cell(11, 2).Range.Text = NumericUpDown9.Value
        oTable.Cell(11, 3).Range.Text = "[C]"

        oTable.Cell(12, 1).Range.Text = "Ruwheid as"
        oTable.Cell(12, 2).Range.Text = NumericUpDown12.Value
        oTable.Cell(12, 3).Range.Text = "[-]"

        '---- --------
        oTable.Cell(13, 1).Range.Text = "Ruwheid naaf"
        oTable.Cell(13, 2).Range.Text = NumericUpDown11.Value
        oTable.Cell(13, 3).Range.Text = "[-]"

        oTable.Columns.Item(1).Width = oWord.InchesToPoints(2.9)   'Change width of columns 1 & 2.
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(0.8)
        oTable.Columns.Item(3).Width = oWord.InchesToPoints(0.9)


        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()
        'Insert a 14 x 5 table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 12, 3)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = 10
        oTable.Range.Font.Bold = False
        oTable.Rows.Item(1).Range.Font.Bold = True

        oTable.Cell(1, 1).Range.Text = "Results"
        oTable.Cell(1, 2).Range.Text = ""
        oTable.Cell(1, 3).Range.Text = ""

        oTable.Cell(2, 1).Range.Text = "Motor koppel"
        oTable.Cell(2, 2).Range.Text = TextBox1.Text
        oTable.Cell(2, 3).Range.Text = "[N.m]"

        oTable.Cell(3, 1).Range.Text = "Vlaktedruk (p)"
        oTable.Cell(3, 2).Range.Text = TextBox5.Text
        oTable.Cell(3, 3).Range.Text = "[N/mm]"

        oTable.Cell(4, 1).Range.Text = "Slipmoment"
        oTable.Cell(4, 2).Range.Text = TextBox15.Text
        oTable.Cell(4, 3).Range.Text = "[N.m]"

        oTable.Cell(5, 1).Range.Text = "Trekspanning (sigma_t max)"
        oTable.Cell(5, 2).Range.Text = TextBox7.Text
        oTable.Cell(5, 3).Range.Text = "[N/mm]"

        oTable.Cell(6, 1).Range.Text = "Radiale drukspanning (sigma_r max)"
        oTable.Cell(6, 2).Range.Text = TextBox4.Text
        oTable.Cell(6, 3).Range.Text = "[N/mm]"

        oTable.Cell(7, 1).Range.Text = "Gecombineerde spanning (sigma_i max)"
        oTable.Cell(7, 2).Range.Text = TextBox10.Text
        oTable.Cell(7, 3).Range.Text = "[N/mm]"

        oTable.Rows.Item(8).Range.Font.Bold = True

        oTable.Cell(8, 1).Range.Text = "Pers of krimpmaat (koude maat)"
        oTable.Cell(8, 2).Range.Text = ""
        oTable.Cell(8, 3).Range.Text = ""

        oTable.Cell(9, 1).Range.Text = "s/d verhouding"
        oTable.Cell(9, 2).Range.Text = TextBox6.Text
        oTable.Cell(9, 3).Range.Text = "[-]"

        oTable.Cell(10, 1).Range.Text = "s maat "
        oTable.Cell(10, 2).Range.Text = TextBox9.Text
        oTable.Cell(10, 3).Range.Text = "[mu]"

        oTable.Cell(11, 1).Range.Text = "Perskracht "
        oTable.Cell(11, 2).Range.Text = TextBox11.Text
        oTable.Cell(11, 3).Range.Text = "[ton]"

        oTable.Rows.Item(11).Range.Font.Bold = True

        oTable.Cell(11, 1).Range.Text = "Warme maat"
        oTable.Cell(11, 2).Range.Text = ""
        oTable.Cell(11, 3).Range.Text = ""

        oTable.Cell(12, 1).Range.Text = "thermische uitzetting"
        oTable.Cell(12, 2).Range.Text = TextBox2.Text
        oTable.Cell(12, 3).Range.Text = "[mu]"

        oTable.Columns.Item(1).Width = oWord.InchesToPoints(2.9)   'Change width of columns 1 & 2.
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(0.8)
        oTable.Columns.Item(3).Width = oWord.InchesToPoints(0.9)

        '    Me.PictureBox1.Image = New System.Drawing.Bitmap("vervormingvanasennaaf.gif")

        '    '    ' PictureBox1.Image = Image.FromFile("vervormingvanasennaaf.gif")

        oPara3 = oDoc.Content.Paragraphs.Add
        'oPara3.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
        'Dim memstream As New MemoryStream
        'Dim bitmap = New System.Drawing.Bitmap(memstream)
        Try
            With oPara3.Range.InlineShapes.AddPicture("PictureBox1.Image") '.ImageLocation)
                .LockAspectRatio = True
                .Width = 300
            End With
        Catch ex As Exception
            '    'MessageBox.Show(ex.Message & "Line 1780")  ' Show the exception's message.
        End Try
    End Sub

    'Private Sub PictureBox1_Click(sender As Object, e As EventArgs)
    '    PictureBox1.Image = My.Resources.vervormingvanasennaaf
    'End Sub



End Class
