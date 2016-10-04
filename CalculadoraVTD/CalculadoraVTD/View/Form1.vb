Public Class Form1
    Dim Opcion As Integer
    Dim Vtd As VTD = New VTD

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Select Case (ComboBox1.SelectedIndex)
            Case (0)
                Label2.Text = "N/A"
                Label3.Text = "N/A"
                Label4.Text = "N/A"
                Button1.Enabled = False
            Case (1)
                Label2.Text = "Distancia(Km):"
                Label3.Text = "Tiempo(k):"
                Label4.Text = "Velocidad(Km/h):"
                Button1.Enabled = True
            Case (2)
                Label2.Text = "Velocidad(Km/h):"
                Label3.Text = "Tiempo(h):"
                Label4.Text = "Distancia(Km):"
                Button1.Enabled = True
            Case (3)
                Label2.Text = "Velocidad(Km/h):"
                Label3.Text = "Distancia(Km):"
                Label4.Text = "Tiempo(min):"
                Button1.Enabled = True
        End Select
    End Sub

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        InicializarCamposEditables()
        InicializarListBox()
    End Sub

    Private Sub InicializarCamposEditables()
        TextBox1.Text = 0
        TextBox2.Text = 0
        TextBox3.Text = 0
    End Sub

    Private Sub LimpiarCamposEditables()
        TextBox1.Text = 0
        TextBox2.Text = 0
        TextBox3.Text = 0
    End Sub

    Private Sub InicializarListBox()
        ComboBox1.Items.Add("Seleccione")
        ComboBox1.Items.Add("Velocidad")
        ComboBox1.Items.Add("Distancia")
        ComboBox1.Items.Add("Tiempo")
        ComboBox1.SelectedIndex = 0
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim Resultado As Double
        If (Not ValidateNegative()) Then
            MsgBox("No se permiten números negativos", MsgBoxStyle.Critical, "Datos Inválidos")
            LimpiarCamposEditables()
            Exit Sub
        End If

        Try
            Select Case (ComboBox1.SelectedIndex)
                Case (1)
                    Resultado = Vtd.calculateVelocidad(CDbl(TextBox1.Text), CDbl(TextBox2.Text))
                Case (2)
                    Resultado = Vtd.calculateDistancia(CDbl(TextBox1.Text), CDbl(TextBox2.Text))
                Case (3)
                    Resultado = Vtd.calculateTiempo(CDbl(TextBox1.Text), CDbl(TextBox2.Text))
            End Select
            TextBox3.Text = CStr(Resultado)
        Catch ex As InvalidCastException
            MsgBox("Recuerde que solo debe ingresar números", MsgBoxStyle.Critical, "Datos Inválidos")
            LimpiarCamposEditables()
        End Try

    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        If Char.IsDigit(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If Char.IsDigit(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Function ValidateNegative()
        If (CDbl(TextBox1.Text()) < 0) Then
            Return False
        End If
        If (CDbl(TextBox2.Text()) < 0) Then
            Return False
        End If
        Return True
    End Function

End Class