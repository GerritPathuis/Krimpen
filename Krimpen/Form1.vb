Imports System.Math
Imports System
Imports System.Globalization
Imports System.Threading

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
End Class
