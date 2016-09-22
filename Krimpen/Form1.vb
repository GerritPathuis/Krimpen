Imports System.Math
Imports System
Imports System.Globalization
Imports System.Threading
Imports Word = Microsoft.Office.Interop.Word
Imports System.Runtime.InteropServices

Public Class Form1

    Public Shared RA_surf() As String = {"  0.2", "  0.4", "  0.8", "  1.6", "  3.2", "  6.3", "  12.5", "  25"}

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")      'Decimal separator "."
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")    'Decimal separator "."

        '-------Fill combobox1, Surface with RA------------------
        For hh = 0 To RA_surf.Length - 1               'Fill combobox 5 emotor data
            ComboBox1.Items.Add(RA_surf(hh))
            ComboBox2.Items.Add(RA_surf(hh))
        Next hh
        ComboBox1.SelectedIndex = 2         'Ra 0.8 voor krimp of persvlak
        ComboBox2.SelectedIndex = 2         'Ra 0.8 voor krimp of persvlak
        Calc()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, NumericUpDown9.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown1.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown13.ValueChanged, NumericUpDown10.ValueChanged, RadioButton1.CheckedChanged, ComboBox2.SelectedIndexChanged, ComboBox1.SelectedIndexChanged
        Calc()
    End Sub

    Private Sub Calc()
        Dim power, moment_motor, speed, factor As Double
        Dim OD_as, OD_naaf, ID_naaf, ring_dikte As Double
        Dim lengte_as, Coeffie_slip, moment_slip, p_vlaktedruk, trekspanning_ring, combi_spanning As Double
        Dim Elast_mod, sd_verhouding, s_maat_mu As Double
        Dim Therm_uitzetting_rvs, Therm_uitzetting_staal, delta_temp As Double
        Dim Coef_exp_staal, Coef_exp_rvs, Bedrijfs_temp, exp_verschil As Double
        Dim F_pers, S_verlies, actual_hot_s, production_s As Double
        Dim ra1, ra2 As Double

        Double.TryParse(TextBox18.Text, Coef_exp_staal)
        Double.TryParse(TextBox19.Text, Coef_exp_rvs)
        Double.TryParse(ComboBox1.SelectedItem, ra1)
        Double.TryParse(ComboBox2.SelectedItem, ra2)

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
        Bedrijfs_temp = NumericUpDown13.Value - NumericUpDown11.Value
        production_s = NumericUpDown10.Value                'koude productie S maat

        '--------------------------------
        If RadioButton1.Checked = True Then
            GroupBox11.Visible = True
        Else
            GroupBox11.Visible = False
        End If

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
            s_maat_mu = sd_verhouding * OD_as * 1000    'Maat noodzakelijk voor het overbrengen Slip moment

            '----- Therm_uitzetting --
            Therm_uitzetting_rvs = delta_temp * Coef_exp_rvs * ID_naaf * 1000           '[mu]
            Therm_uitzetting_staal = delta_temp * Coef_exp_staal * ID_naaf * 1000       '[mu]

            '------ Gecombineerde spanning ------------------ 
            combi_spanning = Sqrt(trekspanning_ring ^ 2 + p_vlaktedruk ^ 2 + trekspanning_ring * p_vlaktedruk)

            '----- Perskracht --
            F_pers = p_vlaktedruk * Coeffie_slip * PI * OD_as * lengte_as   '[N]
            F_pers /= 10000                                                 '[N-> ton

            '----- Oppervlakte ruwheid --------
            S_verlies = Round(1.2 * (ra1 + ra2), 0)        '60% verlies 

            '----- naaf rvs en as staal -----------------------
            exp_verschil = Bedrijfs_temp * OD_as * (Coef_exp_rvs - Coef_exp_staal) * 1000

            '----------------- actual_hot_s ----------------------
            actual_hot_s = production_s - S_verlies - exp_verschil

            '----- Presenteren --------------
            TextBox1.Text = Round(moment_motor, 0).ToString
            TextBox2.Text = Round(Therm_uitzetting_rvs, 0).ToString     'Thermische expansie naaf rvs
            TextBox21.Text = Round(Therm_uitzetting_staal, 0).ToString  'Thermische expansie naaf staal
            TextBox3.Text = Round(ring_dikte, 1).ToString
            TextBox4.Text = Round(p_vlaktedruk, 1).ToString             'Radiale spanning = vlaktedrukColor.Red
            TextBox5.Text = Round(p_vlaktedruk, 1).ToString             'Vlaktedruk as
            TextBox7.Text = Round(trekspanning_ring, 1).ToString        'Trekspanning ring
            TextBox6.Text = Round(sd_verhouding, 4).ToString            's/d 
            TextBox8.Text = Round(1 / sd_verhouding, 0).ToString        'd/s
            TextBox9.Text = Round(s_maat_mu, 0).ToString                's_maat
            TextBox10.Text = Round(combi_spanning, 0).ToString           'gecombineerde spanning naaf

            TextBox11.Text = Round(F_pers, 1).ToString                  'Perskracht [ton]
            TextBox14.Text = Round(S_verlies, 1).ToString               'Verlies door oppervlakte ruwheid [mu]
            TextBox23.Text = Round(S_verlies, 1).ToString               'Verlies door oppervlakte ruwheid [mu]
            TextBox15.Text = Round(moment_slip, 0).ToString             'As begint te slippen [Nm]
            TextBox20.Text = Round(exp_verschil, 0).ToString
            TextBox24.Text = Round(actual_hot_s, 0).ToString            'As begint te slippen [Nm]

            If p_vlaktedruk < 90 Then           'Check vlakte druk
                TextBox5.BackColor = Color.LightGreen
            Else
                TextBox5.BackColor = Color.Red
            End If

            If 1 / sd_verhouding > 750 Then     'Check krimpmaat op over-stressed
                TextBox6.BackColor = Color.LightGreen
                TextBox8.BackColor = Color.LightGreen
                TextBox9.BackColor = Color.LightGreen
            Else
                TextBox6.BackColor = Color.Red
                TextBox8.BackColor = Color.Red
                TextBox9.BackColor = Color.Red
            End If

            If actual_hot_s > s_maat_mu Then     'Slip at operating temperatuur
                TextBox24.BackColor = Color.LightGreen
                NumericUpDown13.BackColor = SystemColors.Window
                Label37.Visible = False
            Else
                TextBox24.BackColor = Color.Red
                NumericUpDown13.BackColor = Color.Red
                Label37.Visible = True
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
        Dim row, font_sizze As Integer
        Dim ss_hot, ss_slip As Double
        Dim ufilename As String

        'Start Word and open the document template. 
        font_sizze = 9
        oWord = CreateObject("Word.Application")
        oWord.Visible = True
        oDoc = oWord.Documents.Add

        'Insert a paragraph at the beginning of the document. 
        oPara1 = oDoc.Content.Paragraphs.Add
        oPara1.Range.Text = "VTK Engineering"
        oPara1.Range.Font.Name = "Arial"
        oPara1.Range.Font.Size = font_sizze + 2
        oPara1.Range.Font.Bold = True
        oPara1.Format.SpaceAfter = 1                '24 pt spacing after paragraph. 
        oPara1.Range.InsertParagraphAfter()

        oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara2.Range.Font.Size = font_sizze
        oPara2.Format.SpaceAfter = 1
        oPara2.Range.Font.Bold = False
        oPara2.Range.Text = "Berekening krimpen en persen van as en naaf" & vbCrLf
        oPara2.Range.InsertParagraphAfter()

        '----------------------------------------------
        'Insert a table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 2)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = False
        oTable.Rows.Item(1).Range.Font.Bold = True

        row = 1
        oTable.Cell(row, 1).Range.Text = "Project Name"
        oTable.Cell(row, 2).Range.Text = TextBox16.Text
        row += 1
        oTable.Cell(row, 1).Range.Text = "Project number "
        oTable.Cell(row, 2).Range.Text = TextBox17.Text
        row += 1
        oTable.Cell(row, 1).Range.Text = "Naaf nummer "
        oTable.Cell(row, 2).Range.Text = TextBox22.Text
        row += 1
        oTable.Cell(row, 1).Range.Text = "Author "
        oTable.Cell(row, 2).Range.Text = Environment.UserName
        row += 1
        oTable.Cell(row, 1).Range.Text = "Date "
        oTable.Cell(row, 2).Range.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        oTable.Columns.Item(1).Width = oWord.InchesToPoints(2.5)   'Change width of columns 1 & 2.
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(2)
        oTable.Rows.Item(1).Range.Font.Bold = True
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        '----------------------------------------------
        'Insert a 14 x 5 table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 13, 3)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = False
        oTable.Rows.Item(1).Range.Font.Bold = True
        row = 1
        oTable.Cell(row, 1).Range.Text = "Input Data"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Diameter as (d_a)"
        oTable.Cell(row, 2).Range.Text = NumericUpDown1.Value
        oTable.Cell(row, 3).Range.Text = "[mm]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Insteek lengte as (l)"
        oTable.Cell(row, 2).Range.Text = NumericUpDown2.Value
        oTable.Cell(row, 3).Range.Text = "[mm]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Buiten diameter naaf (D)"
        oTable.Cell(row, 2).Range.Text = NumericUpDown4.Value
        oTable.Cell(row, 3).Range.Text = "[mm]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Motor vermogen"
        oTable.Cell(row, 2).Range.Text = NumericUpDown6.Value
        oTable.Cell(row, 3).Range.Text = "[Kw]"
        row += 1

        oTable.Cell(row, 1).Range.Text = "Toerental"
        oTable.Cell(row, 2).Range.Text = NumericUpDown5.Value
        oTable.Cell(row, 3).Range.Text = "[rpm]"

        '---- -----
        row += 1
        oTable.Cell(row, 1).Range.Text = "Bedrijfstoeslag factor"
        oTable.Cell(row, 2).Range.Text = NumericUpDown7.Value
        oTable.Cell(row, 3).Range.Text = "[-]"

        row += 1
        oTable.Cell(row, 1).Range.Text = "Frictie coefficient"
        oTable.Cell(row, 2).Range.Text = NumericUpDown3.Value
        oTable.Cell(row, 3).Range.Text = "[-]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "E modulus"
        oTable.Cell(row, 2).Range.Text = NumericUpDown8.Value
        oTable.Cell(row, 3).Range.Text = "[N/mm2]"
        row += 1
        '---- -----
        oTable.Cell(row, 1).Range.Text = "Thermal expansion coefficient rvs"
        oTable.Cell(row, 2).Range.Text = TextBox19.Text
        oTable.Cell(row, 3).Range.Text = "[mm/mm.K]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Opwarming tbv montage"
        oTable.Cell(row, 2).Range.Text = NumericUpDown9.Value
        oTable.Cell(row, 3).Range.Text = "[C]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Ruwheid as"
        oTable.Cell(row, 2).Range.Text = ComboBox1.SelectedItem
        oTable.Cell(row, 3).Range.Text = "[mu]"
        row += 1
        '---- --------
        oTable.Cell(row, 1).Range.Text = "Ruwheid naaf"
        oTable.Cell(row, 2).Range.Text = ComboBox2.SelectedItem
        oTable.Cell(row, 3).Range.Text = "[mu]"

        oTable.Columns.Item(1).Width = oWord.InchesToPoints(2.9)   'Change width of columns 1 & 2.
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(0.8)
        oTable.Columns.Item(3).Width = oWord.InchesToPoints(0.9)
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        'Insert a 16 x 3 table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 16, 3)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = False
        oTable.Rows.Item(1).Range.Font.Bold = True
        row = 1
        oTable.Cell(row, 1).Range.Text = "Results"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Motor koppel"
        oTable.Cell(row, 2).Range.Text = TextBox1.Text
        oTable.Cell(row, 3).Range.Text = "[N.m]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Vlaktedruk (p)"
        oTable.Cell(row, 2).Range.Text = TextBox5.Text
        oTable.Cell(row, 3).Range.Text = "[N/mm]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Slipmoment"
        oTable.Cell(row, 2).Range.Text = TextBox15.Text
        oTable.Cell(row, 3).Range.Text = "[N.m]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Trekspanning (sigma_t max)"
        oTable.Cell(row, 2).Range.Text = TextBox7.Text
        oTable.Cell(row, 3).Range.Text = "[N/mm]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Radiale drukspanning (sigma_r max)"
        oTable.Cell(row, 2).Range.Text = TextBox4.Text
        oTable.Cell(row, 3).Range.Text = "[N/mm]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Gecombineerde spanning (sigma_i max)"
        oTable.Cell(row, 2).Range.Text = TextBox10.Text
        oTable.Cell(row, 3).Range.Text = "[N/mm]"
        row += 2
        oTable.Rows.Item(row).Range.Font.Bold = True
        oTable.Cell(row, 1).Range.Text = "Pers of krimpmaat (koude maat)"
        row += 1
        oTable.Cell(row, 1).Range.Text = "d/s verhouding > 850"
        oTable.Cell(row, 2).Range.Text = TextBox8.Text
        oTable.Cell(row, 3).Range.Text = "[-]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "s maat "
        oTable.Cell(row, 2).Range.Text = TextBox9.Text
        oTable.Cell(row, 3).Range.Text = "[mu]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Perskracht "
        oTable.Cell(row, 2).Range.Text = TextBox11.Text
        oTable.Cell(row, 3).Range.Text = "[ton]"
        row += 2
        oTable.Rows.Item(row).Range.Font.Bold = True
        oTable.Cell(row, 1).Range.Text = "Warme maat"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Thermische Therm_uitzetting_rvs"
        oTable.Cell(row, 2).Range.Text = TextBox2.Text
        oTable.Cell(row, 3).Range.Text = "[mu]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Thermische Therm_uitzetting_staal"
        oTable.Cell(row, 2).Range.Text = TextBox21.Text
        oTable.Cell(row, 3).Range.Text = "[mu]"

        oTable.Columns.Item(1).Width = oWord.InchesToPoints(2.9)   'Change width of columns 1 & 2.
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(0.8)
        oTable.Columns.Item(3).Width = oWord.InchesToPoints(0.9)
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        If RadioButton1.Checked = True Then
            'Insert a 7 x 3 table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 7, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = font_sizze
            oTable.Range.Font.Bold = False
            oTable.Rows.Item(1).Range.Font.Bold = True
            row = 1
            oTable.Rows.Item(row).Range.Font.Bold = True
            oTable.Cell(row, 1).Range.Text = "Samenvatting RVS naaf op stalen as"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Bedrijfstemperatuur"
            oTable.Cell(row, 2).Range.Text = NumericUpDown13.Text.ToString
            oTable.Cell(row, 3).Range.Text = "[C]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Koude s maat"
            oTable.Cell(row, 2).Range.Text = NumericUpDown10.Text.ToString
            oTable.Cell(row, 3).Range.Text = "[mu]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Ra diameter verlies"
            oTable.Cell(row, 2).Range.Text = TextBox23.Text
            oTable.Cell(row, 3).Range.Text = "[mu]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Bedrijfs thermische expansie verlies"
            oTable.Cell(row, 2).Range.Text = TextBox20.Text
            oTable.Cell(row, 3).Range.Text = "[mu]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Resulterende warme bedrijfs S maat"
            oTable.Cell(row, 2).Range.Text = TextBox24.Text
            oTable.Cell(row, 3).Range.Text = "[mu]"

            Double.TryParse(TextBox24.Text, ss_hot)
            Double.TryParse(TextBox9.Text, ss_slip)

            If ss_hot < ss_slip Then     'Slip at operating temperatuur
                row += 1
                oTable.Rows.Item(row).Range.Font.Bold = True
                oTable.Cell(row, 1).Range.Text = "Naaf zit los !!"
            Else
                oTable.Cell(row, 1).Range.Text = "Naaf zit vast !!"
            End If
            oTable.Columns.Item(1).Width = oWord.InchesToPoints(2.9)   'Change width of columns 1 & 2.
            oTable.Columns.Item(2).Width = oWord.InchesToPoints(0.8)
            oTable.Columns.Item(3).Width = oWord.InchesToPoints(0.9)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()
        End If


        ' save the image to c:\Temp
        Dim FilePath As String = My.Computer.FileSystem.SpecialDirectories.MyPictures
        If Me.PictureBox1.Image IsNot Nothing Then
            Me.PictureBox1.Image.Save(IO.Path.Combine(FilePath, "TestFile.jpg"))
        End If

        oPara3 = oDoc.Content.Paragraphs.Add
        Try
            With oPara3.Range.InlineShapes.AddPicture(FilePath & "\TestFile.jpg")
                .LockAspectRatio = True
                .Width = 250
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message & "Could not Load picture")  ' Show the exception's message.
        End Try
        IO.File.Delete(FilePath & "\TestFile.jpg")

        Try
            ufilename = "C:\temp\Naaf_krimp_" & DateTime.Now.ToString("yyyy_MM_dd__HH_mm_ss") & ".docx"
            oDoc.SaveAs(ufilename)
        Catch ex As Exception
            MessageBox.Show(ex.Message)  ' Show the exception's message.
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, RadioButton3.CheckedChanged, NumericUpDown18.ValueChanged, NumericUpDown17.ValueChanged, NumericUpDown16.ValueChanged, NumericUpDown15.ValueChanged, NumericUpDown14.ValueChanged, NumericUpDown12.ValueChanged, TabPage3.Enter
        Dim L1, uitz1, Delta1 As Double
        Dim L2, uitz2, Delta2 As Double
        Dim L3, uitz3, Delta3 As Double
        Dim expansie_coef, uitz_tot As Double

        If RadioButton3.Checked Then        'Staal
            Double.TryParse(TextBox28.Text, expansie_coef)
        Else
            Double.TryParse(TextBox27.Text, expansie_coef)
        End If

        L1 = NumericUpDown12.Value
        L2 = NumericUpDown15.Value
        L3 = NumericUpDown17.Value

        Delta1 = NumericUpDown14.Value
        Delta2 = NumericUpDown16.Value
        Delta3 = NumericUpDown18.Value

        uitz1 = L1 * Delta1 * expansie_coef
        uitz2 = L2 * Delta2 * expansie_coef
        uitz3 = L3 * Delta3 * expansie_coef
        uitz_tot = uitz1 + uitz2 + uitz3

        TextBox29.Text = uitz1.ToString
        TextBox30.Text = uitz2.ToString
        TextBox31.Text = uitz3.ToString
        TextBox32.Text = uitz_tot.ToString
    End Sub
End Class
