Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim currentDate As DateTime = DateTime.Now
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        departDate.Value = currentDate
        departDate.Format = DateTimePickerFormat.Custom
        departDate.CustomFormat = "dd MMM yyyy"
        returnDate.Value = currentDate
        returnDate.Format = DateTimePickerFormat.Custom
        returnDate.CustomFormat = "dd MMM yyyy"
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles departDate.ValueChanged
        CalculateDuration()
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles returnDate.ValueChanged
        CalculateDuration()
    End Sub

    Private Sub CalculateDuration()
        Dim departureDate As DateTime = departDate.Value
        Dim returnDate As Date = Me.returnDate.Value

        Dim duration As TimeSpan = returnDate - departureDate

        Dim totalDays As Integer = CInt(Math.Round(duration.TotalDays))
        Dim nights As Integer = totalDays - 1 ' Subtract 1 day to get the number of nights

        ' Check if the totalDays and nights are negative
        If totalDays <= 0 AndAlso nights <= 0 Then
            durationTextBox.Text = $"Check tour date!"
        Else
            durationTextBox.Text = $"{totalDays}D {nights}N"
        End If
    End Sub
    Private Sub TravelNec_TextChanged(sender As Object, e As EventArgs) Handles visaPrice.TextChanged, bulletTrain.TextChanged, luggagePrice.TextChanged, bukuUmrah.TextChanged, tagLanyard.TextChanged, mutawifMy.TextChanged, tourLeaderMy.TextChanged, zamZam.TextChanged, agentComm.TextChanged, officePrice.TextChanged, miscPrice.TextChanged
        ' Calculate the total price
        Dim totalPrice As Double = 0

        If Not String.IsNullOrEmpty(visaPrice.Text) Then
            totalPrice += Double.Parse(visaPrice.Text)
        End If

        If Not String.IsNullOrEmpty(bulletTrain.Text) Then
            totalPrice += Double.Parse(bulletTrain.Text)
        End If

        If Not String.IsNullOrEmpty(luggagePrice.Text) Then
            totalPrice += Double.Parse(luggagePrice.Text)
        End If

        If Not String.IsNullOrEmpty(bukuUmrah.Text) Then
            totalPrice += Double.Parse(bukuUmrah.Text)
        End If

        If Not String.IsNullOrEmpty(tagLanyard.Text) Then
            totalPrice += Double.Parse(tagLanyard.Text)
        End If

        If Not String.IsNullOrEmpty(mutawifMy.Text) Then
            totalPrice += Double.Parse(mutawifMy.Text)
        End If

        If Not String.IsNullOrEmpty(tourLeaderMy.Text) Then
            totalPrice += Double.Parse(tourLeaderMy.Text)
        End If

        If Not String.IsNullOrEmpty(zamZam.Text) Then
            totalPrice += Double.Parse(zamZam.Text)
        End If

        If Not String.IsNullOrEmpty(agentComm.Text) Then
            totalPrice += Double.Parse(agentComm.Text)
        End If

        If Not String.IsNullOrEmpty(officePrice.Text) Then
            totalPrice += Double.Parse(officePrice.Text)
        End If

        If Not String.IsNullOrEmpty(miscPrice.Text) Then
            totalPrice += Double.Parse(miscPrice.Text)
        End If

        ' Display the total price
        necessTotalPrice.Text = totalPrice.ToString()
    End Sub
End Class

