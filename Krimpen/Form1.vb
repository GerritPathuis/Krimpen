Imports System.Math
Imports System
Imports System.Globalization
Imports System.Threading

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, NumericUpDown9.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown1.ValueChanged, MyBase.Load, NumericUpDown3.ValueChanged
        Dim power, moment_motor, speed, factor As Double
        Dim OD_as, OD_naaf, ID_naaf, ring_dikte As Double
        Dim Area_as, lengte_as, Coeffie_slip, p_vlaktedruk, trekspanning_ring As Double
        Dim Elast_mod, sd_verhouding, s_maat_mu As Double
        Dim uitzetting, delta_temp, CoeffI_term As Double

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

            'MessageBox.Show(OD_as.ToString)
            '----- as -----------------------
            Area_as = PI * OD_as * lengte_as
            p_vlaktedruk = 2 * moment_motor * 1000 / (PI * OD_as ^ 2 * lengte_as * Coeffie_slip)

            '----- ring ---------------------
            trekspanning_ring = p_vlaktedruk * (OD_naaf ^ 2 + ID_naaf ^ 2) / (OD_naaf ^ 2 - ID_naaf ^ 2)

            '----- s/d ---------------------
            sd_verhouding = 2 * p_vlaktedruk * OD_naaf ^ 2 / (Elast_mod * (OD_naaf ^ 2 - ID_naaf ^ 2))

            'MessageBox.Show(p_vlaktedruk.ToString)

            '----- Uitzetting --
            uitzetting = delta_temp * CoeffI_term * ID_naaf * 1000   '[mu]
            s_maat_mu = sd_verhouding * OD_as * 1000

            '----- Presenteren --------------
            TextBox1.Text = Round(moment_motor, 1).ToString
            TextBox2.Text = Round(uitzetting, 0).ToString               'Thermische expansie
            TextBox3.Text = Round(ring_dikte, 1).ToString
            TextBox4.Text = Round(Area_as, 1).ToString                  'Oppervlak as
            TextBox5.Text = Round(p_vlaktedruk, 1).ToString             'Vlaktedruk as
            TextBox7.Text = Round(trekspanning_ring, 1).ToString        'Trekspanning ring
            TextBox6.Text = Round(sd_verhouding, 4).ToString            's/d 
            TextBox8.Text = Round(1 / sd_verhouding, 0).ToString        'd/s
            TextBox9.Text = Round(s_maat_mu, 0).ToString        'd/s

            If p_vlaktedruk < 90 Then
                TextBox5.BackColor = Color.LightGreen
            Else
                TextBox5.BackColor = Color.Red
            End If

        Catch
            MessageBox.Show("Exception")
        End Try
    End Sub
End Class
