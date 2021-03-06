﻿Imports System.Math

Imports System.Globalization
Imports System.Threading
Imports Word = Microsoft.Office.Interop.Word
Imports System.Management

Public Class Form1

    Public Shared RA_surf() As String = {"  0.2", "  0.4", "  0.8", "  1.6", "  3.2", "  6.3", "  12.5", "  25"}
    Public Shared metal_expansion() As String = {
   "Admiralty Brass;20.2",
    "Aluminum;23.6",
    "Aluminum Alloy - 2011;23",
    "Aluminum Alloy - 2017;23.6",
    "Aluminum Alloy - 2024;23.2",
    "Aluminum Alloy - 3003;23.2",
    "Aluminum Alloy - 5052;23.8",
    "Aluminum Alloy - 5086;23.8",
    "Aluminum Alloy - 6061;23.4",
    "Aluminum Alloy - 7075;23.6",
    "Aluminum Bronze;16.2",
    "Antimony;9",
    "Beryllium;12.1",
    "Beryllium Copper;16.7",
    "Bismuth;13",
    "Cast Iron, grey;10.4",
    "Chromium;5.94",
    "Cobalt;12.1",
    "Copper;17.6",
    "Copper-Alloy -C1100 (Electrolytic tough pitch);17.6",
    "Copper-Alloy -C14500 (Free Machining Cu);17.8",
    "Copper-Alloy -C17200, C17300 (Beryllium Cu);17.8",
    "Copper-Alloy -C18200 (Chromium Cu);17.6",
    "Copper-Alloy -C18700 (Leaded Cu);17.6",
    "Copper-Alloy -C22000 (Commercial bronze, 90%);18.4",
    "Copper-Alloy -C23000 (Red brass, 85%);18.7",
    "Copper-Alloy -C26000 (Cartridge brass, 70%);20",
    "Copper-Alloy -C27000 (Yellow brass);20.3",
    "Copper-Alloy -C28000 (Muntz metal, 60%);20.9",
    "Copper-Alloy -C33000 (Low-leaded brass tube);20.2",
    "Copper-Alloy -C35300 (High-leaded brass);20.3",
    "Copper-Alloy -C35600 (Extra high-leaded brass) ;20.5",
    "Copper-Alloy -C36000 (Free machining brass);20.5",
    "Copper-Alloy -C36500 (Leaded Muntz metal);20.9",
    "Copper-Alloy -C46400 (Naval brass);21.2",
    "Copper-Alloy -C51000 (Phosphor bronze, 5% A);17.8",
    "Copper-Alloy -C54400 (Free cutting phos. bronze);17.3",
    "Copper-Alloy -C62300 (Aluminum bronze, 9%);16.2",
    "Copper-Alloy -C62400 (Aluminum bronze, 11%);16.6",
    "Copper-Alloy -Manganese Bronze;21.2",
    "Copper-Alloy -Nickel-Silver;16.2",
    "Copper-Alloy -C63000 (Ni-Al bronze) ;16.2",
    "Cupronickel;16.2",
    "Ductile Iron, A536 (120-90-02);10.9",
    "Gold;14.2",
    "Hastelloy C;9.54",
    "Incoloy;14.4",
    "Inconel;11.5",
    "Iridium;5.94",
    "Iron, nodular pearlitic;11.7",
    "Iron, pure;12.2",
    "Magnesium;25.2",
    "Malleable Iron, A220 (50005, 60004, 80002);13.5",
    "Manganese;21.6",
    "Manganese Bronze;21.2",
    "Molybdenum;5.4",
    "Monel;14",
    "Nickel Wrought;13.3",
    "Nickel-Alloy - Hastelloy C-22;12.4",
    "Nickel-Alloy - Hastelloy C-276;11.2",
    "Nickel-Alloy - Inconel 718;13",
    "Nickel-Alloy - K500;13.7",
    "Nickel-Alloy - Monel;15.7",
    "Nickel-Alloy - Monel 400;13.9",
    "Nickel-Alloy - Nickel 200, 201, 205;15.3",
    "Nickel-Alloy - R405;13.7",
    "Niobium (Columbium);7.02",
    "Osmium;5.04",
    "Platinum;9",
    "Plutonium;35.7",
    "Potassium;82.8",
    "Red Brass;18.7",
    "Rhodium;7.92",
    "Selenium;37.8",
    "Silicon;5.04",
    "Silver;19.8",
    "Sodium;70.2",
    "Steel S355 (@ 100c);11.1",
    "Steel S355 (@ 200c);12.1",
    "Steel S355 (@ 300c);12.9",
    "Stainless - 30100;16.9",
    "Stainless - S30200;17.3",
    "Stainless - S30215;16.2",
    "Stainless - S30300;17.3",
    "Stainless - S30323;17.3",
    "Stainless - S30400;17.3",
    "Stainless - S30430;17.3",
    "Stainless - S30500;17.3",
    "Stainless - S30800;17.3",
    "Stainless - S30900;14.9",
    "Stainless - S30908;14.9",
    "Stainless - S31000;15.8",
    "Stainless - S31008;15.8",
    "Stainless - S31600;15.8",
    "Stainless - S31700;15.8",
    "Stainless - S31703;16.6",
    "Stainless - S32100;16.6",
    "Super Duplex - S32760;13.0",
    "Stainless - S34700;16.6",
    "Stainless - S34800;16.7",
    "Stainless - S38400;17.3",
    "Stainless - S41000;9.9",
    "Stainless - S40300;9.9",
    "Stainless - S40500;10.8",
    "Stainless - S41400;10.4",
    "Stainless - S41600 ;9.9",
    "Stainless - S41623;9.9",
    "Stainless - S42000;10.3",
    "Stainless - S42020;10.3",
    "Stainless - S42200;11.2",
    "Stainless - S42900;10.3",
    "Stainless - S43000;10.4",
    "Stainless - S43020;10.4",
    "Stainless - S43023;10.4",
    "Stainless - S43600;9.36",
    "Stainless - S44002;10.3",
    "Stainless - S44003;10.1",
    "Stainless - S44004;10.3",
    "Stainless - S44600;10.4",
    "Stainless - S50100;11.2",
    "Stainless - S50200;11.2",
    "Tantalum;6.48",
    "Thorium;12.1",
    "Ti-8Mn;10.8",
    "Tin;23",
    "Titanium;8.64",
    "Titanium Alloy - Ti-5Al-2.5Sn;9.54",
    "Tungsten;4.5",
    "Uranium;13.3",
    "Vanadium;7.92",
    "Wrought Carbon Steel;14",
    "Yellow Brass;20.3",
    "Zinc;34.2"}

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Pro_user, HD_number As String
        Dim user_list As New List(Of String)
        Dim hard_disk_list As New List(Of String)
        Dim pass_name As Boolean = False
        Dim pass_disc As Boolean = False

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")      'Decimal separator "."
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")    'Decimal separator "."
        Dim words() As String
        Dim separators() As String = {";"}


        '------ allowed users with hard disc id's -----
        user_list.Add("GP")
        hard_disk_list.Add("3879121263")        'Privee PC, graslaan25

        user_list.Add("GerritP")
        hard_disk_list.Add("2012062914345300")  'VTK PC, GP

        user_list.Add("GerritP")
        hard_disk_list.Add("14290CEE95FC")       'VTK laptop, GP

        user_list.Add("KarelB")
        hard_disk_list.Add("165214800214")      'VTK PC, Karel Bakker

        user_list.Add("JeroenA")
        hard_disk_list.Add("171095402070")      'VTK desktop, Jeroen

        Pro_user = Environment.UserName     'User name on the screen
        HD_number = HardDisc_Id()           'Harddisk identification
        Me.Text &= "  (" & Pro_user & ")"

        'Check user name and disc_id
        For i = 0 To user_list.Count - 1
            If StrComp(LCase(Pro_user), LCase(user_list.Item(i))) = 0 Then pass_name = True
            If CBool(HD_number = Trim(hard_disk_list(i))) Then pass_disc = True
        Next

        'If pass_name = False Or pass_disc = False Then
        '    MessageBox.Show("VTK fan selection program" & vbCrLf & "Access denied, contact GPa" & vbCrLf)
        '    MessageBox.Show("User_name= " & Pro_user & ", Pass name= " & pass_name.ToString)
        '    MessageBox.Show("HD_id= *" & HD_number & "*" & ", Pass disc= " & pass_disc.ToString)
        '    Environment.Exit(0)
        'End If

        '-------Fill combobox1, Surface with RA------------------
        For hh = 0 To RA_surf.Length - 1               'Fill combobox 1,2
            ComboBox1.Items.Add(RA_surf(hh))
            ComboBox2.Items.Add(RA_surf(hh))
        Next hh

        '-------Fill combobox3, Metal expansion------------------
        For hh = 0 To metal_expansion.Length - 1       'Fill combobox 3
            words = metal_expansion(hh).Split(separators, StringSplitOptions.None)
            ComboBox3.Items.Add(words(0))
            ComboBox4.Items.Add(words(0))   'Shaft
            ComboBox5.Items.Add(words(0))   'Hub
        Next hh

        ComboBox1.SelectedIndex = 2         'Ra 0.8 voor krimp of persvlak
        ComboBox2.SelectedIndex = 2         'Ra 0.8 voor krimp of persvlak
        ComboBox3.SelectedIndex = 85        'Stainless 304
        ComboBox4.SelectedIndex = 85        'SHAFT material
        ComboBox5.SelectedIndex = 78        'HUB material

        TextBox26.Text =
        "Persvlakken" & vbTab & "Ra 0.8 tot 1.6" & vbCrLf &
        "Boren" & vbTab & vbTab & "Ra 3.2 tot 6.3" & vbCrLf &
        "Kotteren" & vbTab & vbTab & "Ra 1.6 tot 3.2" & vbCrLf &
        "Draaien" & vbTab & vbTab & "Ra 1.6 tot 6.3" & vbCrLf &
        "Slijpen" & vbTab & vbTab & "Ra 0.4 tot 3.2" & vbCrLf &
        "Borstelen" & vbTab & vbTab & "Ra 0.4" & vbCrLf &
        "Honen" & vbTab & vbTab & "Ra 0.1 tot 0.4"

        Calc()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, NumericUpDown9.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown1.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown13.ValueChanged, NumericUpDown10.ValueChanged, ComboBox2.SelectedIndexChanged, ComboBox1.SelectedIndexChanged, ComboBox5.SelectedIndexChanged, ComboBox4.SelectedIndexChanged
        Calc()
    End Sub

    Private Sub Calc()
        Dim power, moment_motor, speed, factor As Double
        Dim OD_shaft, OD_hub, ID_hub, ring_dikte As Double
        Dim lengte_as, Coeffie_slip, moment_slip, p_vlaktedruk, trekspanning_ring, combi_spanning As Double
        Dim Elast_mod, sd_verhouding, s_maat_mu As Double
        Dim Therm_uitzetting_hub, Therm_uitzetting_shaft, delta_temp As Double
        Dim Therm_uitz_hub_bedrijf, Therm_uitz_shaft_bedrijf As Double
        Dim Coef_exp_shaft, Coef_exp_hub, Bedrijfs_temp, exp_verschil As Double
        Dim F_pers, S_verlies, actual_hot_s, production_s As Double
        Dim ra1, ra2 As Double
        Dim separators() As String = {";"}
        Dim words() As String

        If (ComboBox3.SelectedIndex > -1) And ComboBox4.SelectedIndex > -1 And ComboBox5.SelectedIndex > -1 Then      'Prevent exceptions
            words = metal_expansion(ComboBox4.SelectedIndex).Split(separators, StringSplitOptions.None)
            TextBox18.Text = (Convert.ToDouble(words(1)) / 10 ^ 6).ToString    'Expansion coefficient
            words = metal_expansion(ComboBox5.SelectedIndex).Split(separators, StringSplitOptions.None)
            TextBox19.Text = (Convert.ToDouble(words(1)) / 10 ^ 6).ToString    'Expansion coefficient
        End If

        Double.TryParse(TextBox18.Text, Coef_exp_shaft)
        Double.TryParse(TextBox19.Text, Coef_exp_hub)
        Double.TryParse(CType(ComboBox1.SelectedItem, String), ra1)
        Double.TryParse(CType(ComboBox2.SelectedItem, String), ra2)

        power = NumericUpDown6.Value * 1000
        speed = NumericUpDown5.Value
        factor = NumericUpDown7.Value

        OD_shaft = NumericUpDown1.Value
        OD_hub = NumericUpDown4.Value
        ID_hub = NumericUpDown1.Value                      'ID_hub is praktisch OD_shaft
        lengte_as = NumericUpDown2.Value                    'Insteeklengte as
        Coeffie_slip = NumericUpDown3.Value                 'Coefficient slip
        ring_dikte = (OD_hub - ID_hub) / 2
        Elast_mod = NumericUpDown8.Value

        delta_temp = NumericUpDown9.Value                   'Montage temp
        Bedrijfs_temp = NumericUpDown13.Value - NumericUpDown11.Value   'Bedrijfs temp
        production_s = NumericUpDown10.Value                'koude productie S maat

        '----------- rekenen -----------
        Try
            moment_motor = power * factor / (speed * PI / 30)

            '----- as -----------------------
            p_vlaktedruk = 2 * moment_motor * 1000 / (PI * OD_shaft ^ 2 * lengte_as * Coeffie_slip)    '[N/mm]
            moment_slip = p_vlaktedruk * PI * OD_shaft * lengte_as * Coeffie_slip * 0.5 * OD_shaft / 1000 '[N.m]

            '----- ring ---------------------
            trekspanning_ring = p_vlaktedruk * (OD_hub ^ 2 + ID_hub ^ 2) / (OD_hub ^ 2 - ID_hub ^ 2)

            '----- s/d ---------------------
            sd_verhouding = 2 * p_vlaktedruk * OD_hub ^ 2 / (Elast_mod * (OD_hub ^ 2 - ID_hub ^ 2))
            s_maat_mu = sd_verhouding * OD_shaft * 1000    'Maat noodzakelijk voor het overbrengen Slip moment

            '----- Therm_uitzetting montage --
            Therm_uitzetting_hub = delta_temp * Coef_exp_hub * ID_hub * 1000           '[mu]
            Therm_uitzetting_shaft = delta_temp * Coef_exp_shaft * ID_hub * 1000       '[mu]

            '------ Gecombineerde spanning ------------------ 
            combi_spanning = Sqrt(trekspanning_ring ^ 2 + p_vlaktedruk ^ 2 + trekspanning_ring * p_vlaktedruk)

            '----- Perskracht --
            F_pers = p_vlaktedruk * Coeffie_slip * PI * OD_shaft * lengte_as   '[N]
            F_pers /= 10000                                                 '[N-> ton

            '----- Oppervlakte ruwheid --------
            S_verlies = Round(1.2 * (ra1 + ra2), 0)        '60% verlies 

            '----- hub en as bij bedrijfstemperatuur -----------------------
            Therm_uitz_hub_bedrijf = Bedrijfs_temp * OD_shaft * Coef_exp_hub * 1000
            Therm_uitz_shaft_bedrijf = Bedrijfs_temp * OD_shaft * Coef_exp_shaft * 1000
            exp_verschil = Therm_uitz_hub_bedrijf - Therm_uitz_shaft_bedrijf

            TextBox38.Text = Round(Therm_uitz_shaft_bedrijf, 0).ToString    'Thermische expansie shaft
            TextBox36.Text = Round(Therm_uitz_hub_bedrijf, 0).ToString      'Thermische expansie hub

            '----------------- actual_hot_s ----------------------
            actual_hot_s = production_s - S_verlies - exp_verschil

            '----- Presenteren --------------
            TextBox1.Text = Round(moment_motor, 0).ToString
            TextBox35.Text = Round(Therm_uitzetting_hub, 0).ToString     'Thermische expansie HUB
            TextBox34.Text = Round(Therm_uitzetting_shaft, 0).ToString  'Thermische expansie SHAFT

            TextBox21.Text = Therm_uitzetting_hub.ToString("0")  'Thermische expansie HUB

            TextBox3.Text = Round(ring_dikte, 1).ToString
            TextBox4.Text = Round(p_vlaktedruk, 1).ToString             'Radiale spanning = vlaktedrukColor.Red
            TextBox5.Text = Round(p_vlaktedruk, 1).ToString             'Vlaktedruk as
            TextBox7.Text = Round(trekspanning_ring, 1).ToString        'Trekspanning ring
            TextBox6.Text = Round(sd_verhouding, 4).ToString            's/d 
            TextBox8.Text = Round(1 / sd_verhouding, 0).ToString        'd/s
            TextBox9.Text = Round(s_maat_mu, 0).ToString                's_maat
            TextBox10.Text = Round(combi_spanning, 0).ToString          'gecombineerde spanning hub

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

            TextBox6.Enabled = True
            TextBox8.Enabled = True
            TextBox9.Enabled = True
            TextBox6.ReadOnly = True
            TextBox8.ReadOnly = True
            TextBox9.ReadOnly = True

            If 1 / sd_verhouding > 750 Then     'Check krimpmaat op over-stressed
                TextBox6.BackColor = Color.LightGreen
                TextBox8.BackColor = Color.LightGreen
                TextBox9.BackColor = Color.LightGreen
                TextBox6.ForeColor = Color.Black
                TextBox8.ForeColor = Color.Black
                TextBox9.ForeColor = Color.Black
                Label64.Visible = False
            Else
                TextBox6.BackColor = Color.Red
                TextBox8.BackColor = Color.Red
                TextBox9.BackColor = Color.Red
                TextBox6.ForeColor = Color.White
                TextBox8.ForeColor = Color.White
                TextBox9.ForeColor = Color.Yellow
                Label64.Visible = True
            End If

            If actual_hot_s > s_maat_mu Then     'Slip at operating temperatuur
                TextBox24.BackColor = Color.LightGreen
                NumericUpDown13.BackColor = Color.Yellow
                Label37.Visible = False
            Else
                TextBox24.BackColor = Color.Red
                NumericUpDown13.BackColor = Color.Red
                Label37.Visible = True
            End If

            If NumericUpDown1.Value < NumericUpDown4.Value - 14 Then      'Onmogelijke hub dimensie
                NumericUpDown4.BackColor = Color.LightGreen
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
        oWord = CType(CreateObject("Word.Application"), Word.Application)
        oWord.Visible = True
        oDoc = oWord.Documents.Add

        'Insert a paragraph at the beginning of the document. 
        oPara1 = oDoc.Content.Paragraphs.Add
        oPara1.Range.Text = "VTK Engineering"
        oPara1.Range.Font.Name = "Arial"
        oPara1.Range.Font.Size = font_sizze + 2
        oPara1.Range.Font.Bold = CInt(True)
        oPara1.Format.SpaceAfter = 1                '24 pt spacing after paragraph. 
        oPara1.Range.InsertParagraphAfter()

        oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
        oPara2.Range.Font.Size = font_sizze
        oPara2.Format.SpaceAfter = 1
        oPara2.Range.Font.Bold = CInt(False)
        oPara2.Range.Text = "Berekening krimpen en persen van as en hub" & vbCrLf
        oPara2.Range.InsertParagraphAfter()

        '----------------------------------------------
        'Insert a table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 2)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = CInt(False)
        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)

        row = 1
        oTable.Cell(row, 1).Range.Text = "Project Name"
        oTable.Cell(row, 2).Range.Text = TextBox16.Text
        row += 1
        oTable.Cell(row, 1).Range.Text = "Project number "
        oTable.Cell(row, 2).Range.Text = TextBox17.Text
        row += 1
        oTable.Cell(row, 1).Range.Text = "hub nummer "
        oTable.Cell(row, 2).Range.Text = TextBox22.Text
        row += 1
        oTable.Cell(row, 1).Range.Text = "Author "
        oTable.Cell(row, 2).Range.Text = Environment.UserName
        row += 1
        oTable.Cell(row, 1).Range.Text = "Date "
        oTable.Cell(row, 2).Range.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        oTable.Columns.Item(1).Width = oWord.InchesToPoints(2.5)   'Change width of columns 1 & 2.
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(2)
        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        '----------------------------------------------
        'Insert a 14 x 5 table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 13, 3)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = CInt(False)
        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
        row = 1
        oTable.Cell(row, 1).Range.Text = "Input Data"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Diameter as (d_a)"
        oTable.Cell(row, 2).Range.Text = CType(NumericUpDown1.Value, String)
        oTable.Cell(row, 3).Range.Text = "[mm]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Insteek lengte as (l)"
        oTable.Cell(row, 2).Range.Text = CType(NumericUpDown2.Value, String)
        oTable.Cell(row, 3).Range.Text = "[mm]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Buiten diameter hub (D)"
        oTable.Cell(row, 2).Range.Text = CType(NumericUpDown4.Value, String)
        oTable.Cell(row, 3).Range.Text = "[mm]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Motor vermogen"
        oTable.Cell(row, 2).Range.Text = CType(NumericUpDown6.Value, String)
        oTable.Cell(row, 3).Range.Text = "[Kw]"
        row += 1

        oTable.Cell(row, 1).Range.Text = "Toerental"
        oTable.Cell(row, 2).Range.Text = CType(NumericUpDown5.Value, String)
        oTable.Cell(row, 3).Range.Text = "[rpm]"

        '---- -----
        row += 1
        oTable.Cell(row, 1).Range.Text = "Bedrijfstoeslag factor"
        oTable.Cell(row, 2).Range.Text = CType(NumericUpDown7.Value, String)
        oTable.Cell(row, 3).Range.Text = "[-]"

        row += 1
        oTable.Cell(row, 1).Range.Text = "Frictie coefficient"
        oTable.Cell(row, 2).Range.Text = CType(NumericUpDown3.Value, String)
        oTable.Cell(row, 3).Range.Text = "[-]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "E modulus"
        oTable.Cell(row, 2).Range.Text = CType(NumericUpDown8.Value, String)
        oTable.Cell(row, 3).Range.Text = "[N/mm2]"
        row += 1
        '---- -----
        oTable.Cell(row, 1).Range.Text = "Thermal expansion coefficient rvs"
        oTable.Cell(row, 2).Range.Text = TextBox19.Text
        oTable.Cell(row, 3).Range.Text = "[mm/mm.K]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Opwarming tbv montage"
        oTable.Cell(row, 2).Range.Text = CType(NumericUpDown9.Value, String)
        oTable.Cell(row, 3).Range.Text = "[C]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Ruwheid as"
        oTable.Cell(row, 2).Range.Text = CType(ComboBox1.SelectedItem, String)
        oTable.Cell(row, 3).Range.Text = "[mu]"
        row += 1
        '---- --------
        oTable.Cell(row, 1).Range.Text = "Ruwheid hub"
        oTable.Cell(row, 2).Range.Text = CType(ComboBox2.SelectedItem, String)
        oTable.Cell(row, 3).Range.Text = "[mu]"

        oTable.Columns.Item(1).Width = oWord.InchesToPoints(2.9)   'Change width of columns 1 & 2.
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(0.8)
        oTable.Columns.Item(3).Width = oWord.InchesToPoints(0.9)
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        'Insert a 16 x 3 table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 16, 3)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = CInt(False)
        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
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
        oTable.Rows.Item(row).Range.Font.Bold = CInt(True)
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
        oTable.Rows.Item(row).Range.Font.Bold = CInt(True)
        oTable.Cell(row, 1).Range.Text = "Warme maat"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Thermische uitzetting AS"
        oTable.Cell(row, 2).Range.Text = TextBox38.Text
        oTable.Cell(row, 3).Range.Text = "[mu]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Thermische uitzetting hub"
        oTable.Cell(row, 2).Range.Text = TextBox36.Text
        oTable.Cell(row, 3).Range.Text = "[mu]"

        oTable.Columns.Item(1).Width = oWord.InchesToPoints(2.9)   'Change width of columns 1 & 2.
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(0.8)
        oTable.Columns.Item(3).Width = oWord.InchesToPoints(0.9)
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()


        'Insert a 7 x 3 table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 7, 3)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = font_sizze
        oTable.Range.Font.Bold = CInt(False)
        oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
        row = 1
        oTable.Rows.Item(row).Range.Font.Bold = CInt(True)
        oTable.Cell(row, 1).Range.Text = "Samenvatting RVS hub op stalen as"
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
            oTable.Rows.Item(row).Range.Font.Bold = CInt(True)
            oTable.Cell(row, 1).Range.Text = "hub zit los !!"
        Else
            oTable.Cell(row, 1).Range.Text = "hub zit vast !!"
        End If
        oTable.Columns.Item(1).Width = oWord.InchesToPoints(2.9)   'Change width of columns 1 & 2.
        oTable.Columns.Item(2).Width = oWord.InchesToPoints(0.8)
        oTable.Columns.Item(3).Width = oWord.InchesToPoints(0.9)
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()


        ' save the image to c:\Temp
        Dim FilePath As String = My.Computer.FileSystem.SpecialDirectories.MyPictures
        If Me.PictureBox1.Image IsNot Nothing Then
            Me.PictureBox1.Image.Save(IO.Path.Combine(FilePath, "TestFile.jpg"))
        End If

        oPara3 = oDoc.Content.Paragraphs.Add
        Try
            With oPara3.Range.InlineShapes.AddPicture(FilePath & "\TestFile.jpg")
                .LockAspectRatio = CType(True, Microsoft.Office.Core.MsoTriState)
                .Width = 250
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message & "Could not Load picture")  ' Show the exception's message.
        End Try
        IO.File.Delete(FilePath & "\TestFile.jpg")

        Try
            ufilename = "C:\temp\hub_krimp_" & DateTime.Now.ToString("yyyy_MM_dd__HH_mm_ss") & ".docx"
            oDoc.SaveAs(ufilename.ToString)
        Catch ex As Exception
            MessageBox.Show(ex.Message)  ' Show the exception's message.
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, NumericUpDown18.ValueChanged, NumericUpDown17.ValueChanged, NumericUpDown16.ValueChanged, NumericUpDown15.ValueChanged, NumericUpDown14.ValueChanged, NumericUpDown12.ValueChanged, TabPage3.Enter, ComboBox3.SelectedIndexChanged, NumericUpDown24.ValueChanged, NumericUpDown23.ValueChanged
        Dim L1, uitz1, Delta1 As Double
        Dim L2, uitz2, Delta2 As Double
        Dim L3, uitz3, Delta3 As Double
        Dim L_as, uitz_as, Delta_as As Double
        Dim expansie_coef, uitz_tot As Double
        Dim separators() As String = {";"}
        Double.TryParse(TextBox33.Text, expansie_coef)


        If (ComboBox3.SelectedIndex > -1) Then      'Prevent exceptions
            Dim words() As String = metal_expansion(ComboBox3.SelectedIndex).Split(separators, StringSplitOptions.None)
            TextBox33.Text = (Convert.ToDouble(words(1)) / 10 ^ 6).ToString    'Expansion coefficient
        End If

        L1 = NumericUpDown12.Value
        L2 = NumericUpDown15.Value
        L3 = NumericUpDown17.Value
        L_as = NumericUpDown24.Value

        Delta1 = NumericUpDown14.Value
        Delta2 = NumericUpDown16.Value
        Delta3 = NumericUpDown18.Value
        Delta_as = NumericUpDown23.Value

        uitz1 = L1 * Delta1 * expansie_coef
        uitz2 = L2 * Delta2 * expansie_coef
        uitz3 = L3 * Delta3 * expansie_coef
        uitz_tot = uitz1 + uitz2 + uitz3
        uitz_as = L_as * Delta_as * expansie_coef

        TextBox29.Text = uitz1.ToString
        TextBox30.Text = uitz2.ToString
        TextBox31.Text = uitz3.ToString
        TextBox32.Text = uitz_tot.ToString
        TextBox37.Text = uitz_as.ToString
    End Sub

    Public Function HardDisc_Id() As String
        'Add system.management as reference !!
        Dim tmpStr2 As String = ""
        Dim myScop As New ManagementScope("\\" & Environment.MachineName & "\root\cimv2")
        Dim oQuer As New SelectQuery("SELECT * FROM WIN32_DiskDrive")

        Dim oResult As New ManagementObjectSearcher(myScop, oQuer)
        Dim oIte As ManagementObject
        Dim oPropert As PropertyData
        For Each oIte In oResult.Get()
            For Each oPropert In oIte.Properties
                If Not oPropert.Value Is Nothing AndAlso oPropert.Name = "SerialNumber" Then
                    tmpStr2 = oPropert.Value.ToString
                    Exit For
                End If
            Next
            Exit For
        Next
        Return (Trim(tmpStr2))         'Harddisk identification
    End Function

End Class
