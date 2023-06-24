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
        Dim nights As Integer = totalDays - 1

        If totalDays <= 0 AndAlso nights <= 0 Then
            durationTextBox.Text = $"Check tour date!"
            durationTextBox.BackColor = Color.Pink
        Else
            durationTextBox.Text = $"{totalDays}D {nights}N"
            durationTextBox.BackColor = SystemColors.Window
        End If

    End Sub

    Private Sub Any_TextChanged(sender As Object, e As EventArgs) Handles saudiArrPerPaxSar.TextChanged, exchgRate.TextChanged, tourLeadPax.TextChanged, tourLeadPriceSar.TextChanged,
        saudiArrPax.TextChanged, ovrTotalRm5.TextChanged, ovrTotalRm4.TextChanged, ovrTotalRm3.TextChanged, ovrTotalRm2.TextChanged, ovrTotalSar5.TextChanged, ovrTotalSar4.TextChanged,
        ovrTotalSar3.TextChanged, ovrTotalSar2.TextChanged, hotelNights5.TextChanged, hotelNights4.TextChanged, hotelNights3.TextChanged, hotelNights2.TextChanged,
        accTotalPrice5.TextChanged, accTotalPrice4.TextChanged, accTotalPrice3.TextChanged, accTotalPrice2.TextChanged, mealPrice5.TextChanged, mealPrice4.TextChanged,
        mealPrice3.TextChanged, mealPrice2.TextChanged, paxPrice5.TextChanged, paxPrice4.TextChanged, paxPrice3.TextChanged, paxPrice2.TextChanged, roomPrice5.TextChanged,
        roomPrice4.TextChanged, roomPrice3.TextChanged, roomPrice2.TextChanged, ovrTotalMadRm5.TextChanged, ovrTotalMadRm4.TextChanged, ovrTotalMadRm3.TextChanged,
        ovrTotalMadRm2.TextChanged, ovrTotalMadSar5.TextChanged, ovrTotalMadSar4.TextChanged, ovrTotalMadSar3.TextChanged, ovrTotalMadSar2.TextChanged, hotelNightsMad5.TextChanged,
        hotelNightsMad4.TextChanged, hotelNightsMad3.TextChanged, hotelNightsMad2.TextChanged, accTotalPriceMad5.TextChanged, accTotalPriceMad4.TextChanged,
        accTotalPriceMad3.TextChanged, accTotalPriceMad2.TextChanged, mealPriceMad5.TextChanged, mealPriceMad4.TextChanged, mealPriceMad3.TextChanged, mealPriceMad2.TextChanged,
        paxPriceMad5.TextChanged, paxPriceMad4.TextChanged, paxPriceMad3.TextChanged, paxPriceMad2.TextChanged, roomPriceMad5.TextChanged, roomPriceMad4.TextChanged,
        roomPriceMad3.TextChanged, roomPriceMad2.TextChanged, tourLeadPriceRm.TextChanged, saudiArrPriceRm.TextChanged, tourLeadPerPaxSar.TextChanged, saudiArrPriceSar.TextChanged,
        agentComm.TextChanged, officePrice.TextChanged, miscPrice1.TextChanged, tagLanyard.TextChanged, mutawifMy.TextChanged, tourLeaderMy.TextChanged, zamZam.TextChanged, visaPrice.TextChanged, bulletTrain.TextChanged, luggagePrice.TextChanged, bukuUmrah.TextChanged, peakEco5.TextChanged, peakBus5.TextChanged,
        peakEco4.TextChanged, peakBus4.TextChanged, peakEco3.TextChanged, peakBus3.TextChanged, peakEco2.TextChanged, peakBus2.TextChanged, babyPax.TextChanged, babyPolicy.TextChanged,
        childwPolicy.TextChanged, childwoPolicy.TextChanged, ticketPriceBus.TextChanged, ticketPriceEco.TextChanged, childwoBedPriceBus2.TextChanged, childwBedPriceBus2.TextChanged,
        childwoBedPaxBus2.TextChanged, childwBedPaxBus2.TextChanged, adultPaxBus2.TextChanged, adultPriceBus2.TextChanged, childwoBedPriceEco2.TextChanged,
        childwBedPriceEco2.TextChanged, childwoBedPaxEco2.TextChanged, childwBedPaxEco2.TextChanged, adultPaxEco2.TextChanged, adultPriceEco2.TextChanged, TabPage4.TextChanged,
        childwoBedPriceBus3.TextChanged, childwBedPriceBus3.TextChanged, childwoBedPaxBus3.TextChanged, childwBedPaxBus3.TextChanged, adultPaxBus3.TextChanged,
        adultPriceBus3.TextChanged, childwoBedPriceEco3.TextChanged, childwBedPriceEco3.TextChanged, childwoBedPaxEco3.TextChanged, childwBedPaxEco3.TextChanged,
        adultPaxEco3.TextChanged, adultPriceEco3.TextChanged, childwoBedPriceBus4.TextChanged, childwBedPriceBus4.TextChanged, childwoBedPaxBus4.TextChanged,
        childwBedPaxBus4.TextChanged, adultPaxBus4.TextChanged, adultPriceBus4.TextChanged, childwoBedPriceEco4.TextChanged, childwBedPriceEco4.TextChanged,
        childwoBedPaxEco4.TextChanged, childwBedPaxEco4.TextChanged, adultPaxEco4.TextChanged, adultPriceEco4.TextChanged, childwoBedPriceBus5.TextChanged,
        childwBedPriceBus5.TextChanged, childwoBedPaxBus5.TextChanged, childwBedPaxBus5.TextChanged, adultPaxBus5.TextChanged, adultPriceBus5.TextChanged,
        childwoBedPriceEco5.TextChanged, childwBedPriceEco5.TextChanged, childwoBedPaxEco5.TextChanged, childwBedPaxEco5.TextChanged, adultPaxEco5.TextChanged,
        adultPriceEco5.TextChanged, miscPrice2.TextChanged

        Update_SaudiArrangements()
        Update_Accommodations()
        Update_Costing()
        Update_Package_Price()
    End Sub

    Private Sub Update_SaudiArrangements()
        Dim saudiArrPerPaxPrice, tourLeadPerPaxPrice, tourLeadOvrValue, excRate, SaudiArrOvrValue As Double
        Dim tourLeadPaxValue, saudiArrPaxValue As Integer


        If Integer.TryParse(saudiArrPax.Text, saudiArrPaxValue) AndAlso Double.TryParse(saudiArrPerPaxSar.Text, saudiArrPerPaxPrice) AndAlso saudiArrPaxValue > 0 Then
            SaudiArrOvrValue = saudiArrPaxValue * saudiArrPerPaxPrice
            saudiArrPriceSar.Text = SaudiArrOvrValue.ToString("0.00")
        Else
            saudiArrPriceSar.Text = String.Empty
        End If

        If Integer.TryParse(tourLeadPax.Text, tourLeadPaxValue) AndAlso Double.TryParse(tourLeadPriceSar.Text, tourLeadOvrValue) AndAlso tourLeadPaxValue > 0 Then
            tourLeadPerPaxPrice = tourLeadOvrValue / tourLeadPaxValue
            tourLeadPerPaxSar.Text = tourLeadPerPaxPrice.ToString("0.00")
        Else
            tourLeadPerPaxSar.Text = String.Empty
        End If

        If Double.TryParse(exchgRate.Text, excRate) AndAlso Double.TryParse(saudiArrPerPaxSar.Text, saudiArrPerPaxPrice) Then
            Dim saudiArrRmPrice As Double = excRate * saudiArrPerPaxPrice
            saudiArrPriceRm.Text = saudiArrRmPrice.ToString("0.00")
            If Double.TryParse(exchgRate.Text, excRate) AndAlso Double.TryParse(tourLeadPerPaxSar.Text, tourLeadPerPaxPrice) Then
                Dim tourLeadRmPrice As Double = excRate * tourLeadPerPaxPrice
                tourLeadPriceRm.Text = tourLeadRmPrice.ToString("0.00")
            Else
                tourLeadPriceRm.Text = String.Empty
            End If
        Else
            saudiArrPriceRm.Text = String.Empty
        End If
    End Sub

    Private Sub Update_Accommodations()
        Dim roomPriceTextbox() As TextBox = {roomPrice2, roomPrice3, roomPrice4, roomPrice5, roomPriceMad2, roomPriceMad3, roomPriceMad4, roomPriceMad5}
        Dim paxPriceTextbox() As TextBox = {paxPrice2, paxPrice3, paxPrice4, paxPrice5, paxPriceMad2, paxPriceMad3, paxPriceMad4, paxPriceMad5}
        Dim mealPriceTextbox() As TextBox = {mealPrice2, mealPrice3, mealPrice4, mealPrice5, mealPriceMad2, mealPriceMad3, mealPriceMad4, mealPriceMad5}
        Dim totalPriceTextbox() As TextBox = {accTotalPrice2, accTotalPrice3, accTotalPrice4, accTotalPrice5, accTotalPriceMad2, accTotalPriceMad3, accTotalPriceMad4, accTotalPriceMad5}
        Dim accNightsTextbox() As TextBox = {hotelNights2, hotelNights3, hotelNights4, hotelNights5, hotelNightsMad2, hotelNightsMad3, hotelNightsMad4, hotelNightsMad5}
        Dim accOvrPriceSarTextbox() As TextBox = {ovrTotalSar2, ovrTotalSar3, ovrTotalSar4, ovrTotalSar5, ovrTotalMadSar2, ovrTotalMadSar3, ovrTotalMadSar4, ovrTotalMadSar5}
        Dim accOvrPriceRmTextbox() As TextBox = {ovrTotalRm2, ovrTotalRm3, ovrTotalRm4, ovrTotalRm5, ovrTotalMadRm2, ovrTotalMadRm3, ovrTotalMadRm4, ovrTotalMadRm5}
        Dim doubleExcRate, doubleRoomPriceSar, doublePaxPriceSar, doubleMealPriceSar, doubleTotalPriceSar As Double
        Dim intAccNights As Integer

        For Each textbox As TextBox In roomPriceTextbox
            Dim indexOfRoomPrice As Integer = Array.IndexOf(roomPriceTextbox, textbox)

            If Double.TryParse(textbox.Text, doubleRoomPriceSar) AndAlso doubleRoomPriceSar >= 0 Then
                doublePaxPriceSar = doubleRoomPriceSar / 2
                paxPriceTextbox(indexOfRoomPrice).Text = doublePaxPriceSar.ToString("0.00")
            Else
                paxPriceTextbox(indexOfRoomPrice).Text = ""
            End If

        Next

        For Each textbox As TextBox In mealPriceTextbox
            Dim indexOfMealPrice As Integer = Array.IndexOf(mealPriceTextbox, textbox)

            If Double.TryParse(textbox.Text, doubleMealPriceSar) AndAlso Double.TryParse(paxPriceTextbox(indexOfMealPrice).Text, doublePaxPriceSar) AndAlso doubleMealPriceSar >= 0 Then
                doubleTotalPriceSar = doubleMealPriceSar + doublePaxPriceSar
                totalPriceTextbox(indexOfMealPrice).Text = doubleTotalPriceSar.ToString("0.00")
            Else
                totalPriceTextbox(indexOfMealPrice).Text = ""
            End If
        Next

        For Each textbox As TextBox In accNightsTextbox
            Dim indexOfAccNights As Integer = Array.IndexOf(accNightsTextbox, textbox)
            Dim accNightsText As String = textbox.Text.Trim()
            Double.TryParse(exchgRate.Text, doubleExcRate)

            If Integer.TryParse(textbox.Text, intAccNights) AndAlso Double.TryParse(totalPriceTextbox(indexOfAccNights).Text, doubleTotalPriceSar) AndAlso intAccNights >= 0 Then
                Dim accOvrPriceSar As Double = intAccNights * doubleTotalPriceSar
                Dim accOvrPriceRm As Double = accOvrPriceSar * doubleExcRate
                accOvrPriceSarTextbox(indexOfAccNights).Text = accOvrPriceSar.ToString("0.00")
                accOvrPriceRmTextbox(indexOfAccNights).Text = accOvrPriceRm.ToString("0.00")
            Else
                accOvrPriceSarTextbox(indexOfAccNights).Text = ""
                accOvrPriceRmTextbox(indexOfAccNights).Text = ""
            End If
        Next

    End Sub

    Private Sub Update_Costing()
        Dim travelNecessitiesTextBox() As TextBox = {visaPrice, bulletTrain, luggagePrice, bukuUmrah, tagLanyard, mutawifMy, tourLeaderMy,
            zamZam, agentComm, officePrice, miscPrice1, miscPrice2}
        Dim makAccOvrPriceRmTextbox() As TextBox = {ovrTotalRm2, ovrTotalRm3, ovrTotalRm4, ovrTotalRm5}
        Dim madAccOvrPriceRmTextbox() As TextBox = {ovrTotalMadRm2, ovrTotalMadRm3, ovrTotalMadRm4, ovrTotalMadRm5}
        Dim paxEcoTextBox() As TextBox = {adultPaxEco2, childwBedPaxEco2, childwoBedPaxEco2, adultPaxEco3, childwBedPaxEco3, childwoBedPaxEco3, adultPaxEco4, childwBedPaxEco4,
            childwoBedPaxEco4, adultPaxEco5, childwBedPaxEco5, childwoBedPaxEco5}
        Dim paxBusTextBox() As TextBox = {adultPaxBus2, childwBedPaxBus2, childwoBedPaxBus2, adultPaxBus3, childwBedPaxBus3, childwoBedPaxBus3, adultPaxBus4, childwBedPaxBus4,
            childwoBedPaxBus4, adultPaxBus5, childwBedPaxBus5, childwoBedPaxBus5}
        Dim totalCostEcoTextBox() As TextBox = {}
        Dim totalCostBusTextBox() As TextBox = {}
        Dim peakSeasonEcoTextBox() As TextBox = {peakEco2, peakEco3, peakEco4, peakEco5}
        Dim peakSeasonBusTextBox() As TextBox = {peakBus2, peakBus3, peakBus4, peakBus5}
        Dim doubleTicketPriceEco, doubleTicketPriceBus, doubleSaudiArrRm, doubleTourLeadRm As Double
        Dim totalNecessPrice As Double = 0
        Dim costingTotal As Double = 0
        Dim costingTotalEco As Double = 0
        Dim costingTotalBus As Double = 0

        For Each textbox As TextBox In travelNecessitiesTextBox
            Dim strTravelNec As String = textbox.Text.Trim()

            If Not String.IsNullOrEmpty(strTravelNec) Then
                Dim doubleTravelNec As Double

                If Double.TryParse(textbox.Text, doubleTravelNec) Then
                    totalNecessPrice += doubleTravelNec
                End If
            End If
        Next

        Double.TryParse(ovrTotalRm2.Text, doubleSaudiArrRm)
        Double.TryParse(saudiArrPriceRm.Text, doubleSaudiArrRm)
        Double.TryParse(saudiArrPriceRm.Text, doubleSaudiArrRm)
        Double.TryParse(tourLeadPriceRm.Text, doubleTourLeadRm)
        Double.TryParse(ticketPriceEco.Text, doubleTicketPriceEco)
        Double.TryParse(ticketPriceBus.Text, doubleTicketPriceBus)

        costingTotal = totalNecessPrice + doubleSaudiArrRm + doubleTourLeadRm
        costingTotalEco = costingTotal + doubleTicketPriceEco
        costingTotalBus = costingTotal + doubleTicketPriceBus

        For Each textbox As TextBox In paxEcoTextBox
            Dim indexOfPaxEco As Integer = Array.IndexOf(paxEcoTextBox, textbox)
            Dim strPaxEco As String = textbox.Text.Trim()
            Dim doublePaxEco As Integer
            Dim doublemakAccPrice, doublemadAccPrice As Double
            Dim arrayCoord As Integer = indexOfPaxEco \ 3

            If Not String.IsNullOrEmpty(strPaxEco) Then
                If indexOfPaxEco >= 0 AndAlso indexOfPaxEco <= 2 Then
                    If Integer.TryParse(textbox.Text, doublePaxEco) AndAlso Double.TryParse(makAccOvrPriceRmTextbox(arrayCoord).Text, doublemakAccPrice) AndAlso Double.TryParse(madAccOvrPriceRmTextbox(arrayCoord).Text, doublemadAccPrice) Then
                        costingTotalEco += doublePaxEco * costingTotalEco + (doublemadAccPrice + doublemakAccPrice)
                        totalCostEcoTextBox(indexOfPaxEco).Text = costingTotalEco.ToString("0.00")
                    End If
                End If

            End If
        Next

        For Each textbox As TextBox In paxBusTextBox
            Dim indexOfPaxBus As Integer = Array.IndexOf(paxBusTextBox, textbox)
            Dim strPaxBus As String = textbox.Text.Trim()
            Dim doublePaxBus As Integer
            Dim doublemakAccPrice, doublemadAccPrice As Double
            Dim arrayCoord As Integer = indexOfPaxBus \ 3

            If Not String.IsNullOrEmpty(strPaxBus) Then
                If Integer.TryParse(textbox.Text, doublePaxBus) AndAlso Double.TryParse(makAccOvrPriceRmTextbox(arrayCoord).Text, doublemakAccPrice) AndAlso Double.TryParse(madAccOvrPriceRmTextbox(arrayCoord).Text, doublemadAccPrice) Then
                    costingTotalBus += doublePaxBus * costingTotalBus + (doublemadAccPrice + doublemakAccPrice)
                    totalCostBusTextBox(indexOfPaxBus).Text = costingTotalBus.ToString("0.00")
                End If
            End If
        Next

        For Each textbox As TextBox In totalCostEcoTextBox
            Dim strTotalCostEco As String = textbox.Text.Trim()

            If Not String.IsNullOrEmpty(strTotalCostEco) Then
                Dim doubleTotalCostEco As Double

                If Double.TryParse(textbox.Text, doubleTotalCostEco) Then
                    costingTotalEco += doubleTotalCostEco
                End If
            End If
        Next

    End Sub

    Private Sub Update_Package_Price()
        Dim adultPriceEcoTextBox() As TextBox = {adultPriceEco2, adultPriceEco3, adultPriceEco4, adultPriceEco5}
        Dim adultPriceBusTextBox() As TextBox = {adultPriceBus2, adultPriceBus3, adultPriceBus4, adultPriceBus5}
        Dim childwBedPriceEcoTextBox() As TextBox = {childwBedPriceEco2, childwBedPriceEco3, childwBedPriceEco4, childwBedPriceEco5}
        Dim childwBedPriceBusTextBox() As TextBox = {childwBedPriceBus2, childwBedPriceBus3, childwBedPriceBus4, childwBedPriceBus5}
        Dim childwoBedPriceEcoTextBox() As TextBox = {childwoBedPriceEco2, childwoBedPriceEco3, childwoBedPriceEco4, childwoBedPriceEco5}
        Dim childwoBedPriceBusTextBox() As TextBox = {childwoBedPriceBus2, childwoBedPriceBus3, childwoBedPriceBus4, childwoBedPriceBus5}
        Dim doubleChildwPolicy, doubleChildwoPolicy, doubleAdultPriceEco, doubleAdultPriceBus As Double

        For Each textbox As TextBox In adultPriceEcoTextBox
            Dim indexOfadultPriceEco As Integer = Array.IndexOf(adultPriceEcoTextBox, textbox)
            Dim strAdultPriceEco As String = textbox.Text.Trim()

            If Not String.IsNullOrEmpty(strAdultPriceEco) Then
                If Double.TryParse(textbox.Text, doubleAdultPriceEco) AndAlso doubleAdultPriceEco > 0 Then
                    textbox.BackColor = SystemColors.Window
                    If Double.TryParse(childwPolicy.Text, doubleChildwPolicy) AndAlso doubleChildwPolicy < 0 Then
                        Dim doubleChildwPolicyPrice As Double = doubleAdultPriceEco + doubleChildwPolicy
                        childwBedPriceEcoTextBox(indexOfadultPriceEco).Text = doubleChildwPolicyPrice.ToString("0.00")
                        If Double.TryParse(childwoPolicy.Text, doubleChildwoPolicy) AndAlso doubleChildwoPolicy < 0 Then
                            Dim doubleChildwoPolicyPrice As Double = doubleAdultPriceEco + doubleChildwoPolicy
                            childwoBedPriceEcoTextBox(indexOfadultPriceEco).Text = doubleChildwoPolicyPrice.ToString("0.00")
                        Else
                            childwoBedPriceEcoTextBox(indexOfadultPriceEco).Text = "0"
                        End If
                    Else
                        childwBedPriceEcoTextBox(indexOfadultPriceEco).Text = "0"
                    End If
                End If
            Else
                textbox.BackColor = Color.Pink
                textbox.Text = "0"
            End If
        Next

        For Each textbox As TextBox In adultPriceBusTextBox
            Dim indexOfadultPriceBus As Integer = Array.IndexOf(adultPriceBusTextBox, textbox)
            Dim strAdultPriceBus As String = textbox.Text.Trim()

            If Not String.IsNullOrEmpty(strAdultPriceBus) Then
                If Double.TryParse(textbox.Text, doubleAdultPriceBus) AndAlso doubleAdultPriceBus > 0 Then
                    textbox.BackColor = SystemColors.Window
                    If Double.TryParse(childwPolicy.Text, doubleChildwPolicy) AndAlso doubleChildwPolicy < 0 Then
                        Dim doubleChildwPolicyPrice As Double = doubleAdultPriceBus + doubleChildwPolicy
                        childwBedPriceBusTextBox(indexOfadultPriceBus).Text = doubleChildwPolicyPrice.ToString("0.00")
                        If Double.TryParse(childwoPolicy.Text, doubleChildwoPolicy) AndAlso doubleChildwoPolicy < 0 Then
                            Dim doubleChildwoPolicyPrice As Double = doubleAdultPriceBus + doubleChildwoPolicy
                            childwoBedPriceBusTextBox(indexOfadultPriceBus).Text = doubleChildwoPolicyPrice.ToString("0.00")
                        Else
                            childwoBedPriceBusTextBox(indexOfadultPriceBus).Text = "0"
                        End If
                    Else
                        childwBedPriceEcoTextBox(indexOfadultPriceBus).Text = "0"
                    End If
                End If
            Else
                textbox.BackColor = Color.Pink
                textbox.Text = "0"
            End If
        Next

    End Sub

End Class

