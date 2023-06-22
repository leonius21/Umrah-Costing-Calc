Imports System.Runtime.Intrinsics.X86

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

        Dim textboxes() As TextBox = {visaPrice, bulletTrain, luggagePrice, bukuUmrah, tagLanyard, mutawifMy, tourLeaderMy, zamZam, agentComm, officePrice, miscPrice}

        For Each textbox As TextBox In textboxes
            If Not String.IsNullOrEmpty(textbox.Text) Then
                totalPrice += Double.Parse(textbox.Text)
            End If
        Next

        ' Display the total price
        necessTotalPrice.Text = totalPrice.ToString()

    End Sub

    Private Sub SaudiArr_TextChanged(sender As Object, e As EventArgs) Handles saudiArrPerPaxSar.TextChanged, saudiArrPax.TextChanged
        Dim ovrPrice As Double
        Dim paxValue As Integer
        Dim pricePerPaxValue As Double

        If Integer.TryParse(saudiArrPax.Text, paxValue) AndAlso Double.TryParse(saudiArrPerPaxSar.Text, pricePerPaxValue) Then
            ovrPrice = paxValue * pricePerPaxValue
            saudiArrPriceSar.Text = ovrPrice.ToString("0.00")
        Else
            saudiArrPriceSar.Text = ""
        End If

    End Sub

    Private Sub TourLead_TextChanged(sender As Object, e As EventArgs) Handles tourLeadPax.TextChanged, tourLeadPriceSar.TextChanged
        Dim perPaxPrice As Double
        Dim paxValue As Integer
        Dim ovrValue As Double

        If Integer.TryParse(tourLeadPax.Text, paxValue) AndAlso Double.TryParse(tourLeadPriceSar.Text, ovrValue) Then
            perPaxPrice = paxValue * ovrValue
            tourLeadPerPax.Text = perPaxPrice.ToString("0.00")
        Else
            tourLeadPerPax.Text = ""
        End If

    End Sub

    Private Sub ExchgRate_TextChanged(sender As Object, e As EventArgs) Handles exchgRate.TextChanged
        Dim excRate As Double
        Dim saudiArrSar As Double
        Dim tourLeadSar As Double
        If Double.TryParse(exchgRate.Text, excRate) AndAlso Double.TryParse(saudiArrPriceSar.Text, saudiArrSar) Then
            Dim saudiArrRmPrice As Double = excRate * saudiArrSar
            saudiArrPriceRm.Text = saudiArrRmPrice.ToString("0.00")
            If Double.TryParse(exchgRate.Text, excRate) AndAlso Double.TryParse(tourLeadPriceSar.Text, tourLeadSar) Then
                Dim tourLeadRmPrice As Double = excRate * tourLeadSar
                tourLeadPriceRm.Text = tourLeadRmPrice.ToString("0.00")
            Else
                tourLeadPriceRm.Text = ""
            End If
        Else
            saudiArrPriceRm.Text = ""
        End If

    End Sub
End Class

