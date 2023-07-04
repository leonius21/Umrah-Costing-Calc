Imports System.Drawing.Imaging
Imports System.Reflection.Emit
Imports System.Runtime.InteropServices.JavaScript.JSType
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
        Update_Package_Price()
        Update_Profit_Costing()
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
    Private Sub Update_Package_Price()
        Dim adultPriceEcoTextBox() As TextBox = {adultPriceEco2, adultPriceEco3, adultPriceEco4, adultPriceEco5}
        Dim adultPriceBusTextBox() As TextBox = {adultPriceBus2, adultPriceBus3, adultPriceBus4, adultPriceBus5}
        Dim childwBedPriceEcoTextBox() As TextBox = {childwBedPriceEco2, childwBedPriceEco3, childwBedPriceEco4, childwBedPriceEco5}
        Dim childwBedPriceBusTextBox() As TextBox = {childwBedPriceBus2, childwBedPriceBus3, childwBedPriceBus4, childwBedPriceBus5}
        Dim childwoBedPriceEcoTextBox() As TextBox = {childwoBedPriceEco2, childwoBedPriceEco3, childwoBedPriceEco4, childwoBedPriceEco5}
        Dim childwoBedPriceBusTextBox() As TextBox = {childwoBedPriceBus2, childwoBedPriceBus3, childwoBedPriceBus4, childwoBedPriceBus5}
        Dim doubleChildwPolicy, doubleChildwoPolicy, doubleAdultPriceEco, doubleAdultPriceBus As Double

        For Each textbox As TextBox In adultPriceEcoTextBox
            Dim strAdultPriceEco As String = textbox.Text.Trim()

            If String.IsNullOrEmpty(strAdultPriceEco) Then
                textbox.Text = "0"
                Continue For
            End If

            If Not Double.TryParse(textbox.Text, doubleAdultPriceEco) Then
                Continue For
            End If

            Dim indexOfadultPriceEco As Integer = Array.IndexOf(adultPriceEcoTextBox, textbox)

            If Double.TryParse(childwPolicy.Text, doubleChildwPolicy) Then
                If doubleChildwPolicy < 0 AndAlso Double.TryParse(textbox.Text, doubleAdultPriceEco) AndAlso doubleAdultPriceEco > 0 Then
                    Dim doubleChildwPolicyPrice As Double = doubleAdultPriceEco + doubleChildwPolicy
                    childwBedPriceEcoTextBox(indexOfadultPriceEco).Text = doubleChildwPolicyPrice.ToString("0.00")
                ElseIf Double.TryParse(textbox.Text, doubleAdultPriceEco) AndAlso doubleAdultPriceEco = 0 Then
                    childwBedPriceEcoTextBox(indexOfadultPriceEco).Text = ""
                End If
            End If

            If Double.TryParse(childwoPolicy.Text, doubleChildwoPolicy) Then
                If doubleChildwoPolicy < 0 AndAlso Double.TryParse(textbox.Text, doubleAdultPriceEco) AndAlso doubleAdultPriceEco > 0 Then
                    Dim doubleChildwoPolicyPrice As Double = doubleAdultPriceEco + doubleChildwoPolicy
                    childwoBedPriceEcoTextBox(indexOfadultPriceEco).Text = doubleChildwoPolicyPrice.ToString("0.00")
                ElseIf Double.TryParse(textbox.Text, doubleAdultPriceEco) AndAlso doubleAdultPriceEco = 0 Then
                    childwoBedPriceEcoTextBox(indexOfadultPriceEco).Text = ""
                End If
            End If
        Next

        For Each textbox As TextBox In adultPriceBusTextBox
            Dim strAdultPriceBus As String = textbox.Text.Trim()

            If String.IsNullOrEmpty(strAdultPriceBus) Then
                textbox.Text = "0"
                Continue For
            End If

            If Not Double.TryParse(textbox.Text, doubleAdultPriceBus) AndAlso doubleAdultPriceBus > 0 Then
                Continue For
            End If

            Dim indexOfadultPriceBus As Integer = Array.IndexOf(adultPriceBusTextBox, textbox)

            If Double.TryParse(childwPolicy.Text, doubleChildwPolicy) Then
                If doubleChildwPolicy < 0 AndAlso Double.TryParse(textbox.Text, doubleAdultPriceBus) AndAlso doubleAdultPriceBus > 0 Then
                    Dim doubleChildwPolicyPrice As Double = doubleAdultPriceBus + doubleChildwPolicy
                    childwBedPriceBusTextBox(indexOfadultPriceBus).Text = doubleChildwPolicyPrice.ToString("0.00")
                ElseIf Double.TryParse(textbox.Text, doubleAdultPriceBus) AndAlso doubleAdultPriceBus = 0 Then
                    childwBedPriceBusTextBox(indexOfadultPriceBus).Text = ""
                End If
            End If

            If Double.TryParse(childwoPolicy.Text, doubleChildwoPolicy) Then
                If doubleChildwoPolicy < 0 AndAlso Double.TryParse(textbox.Text, doubleAdultPriceBus) AndAlso doubleAdultPriceBus > 0 Then
                    Dim doubleChildwoPolicyPrice As Double = doubleAdultPriceBus + doubleChildwoPolicy
                    childwoBedPriceBusTextBox(indexOfadultPriceBus).Text = doubleChildwoPolicyPrice.ToString("0.00")
                ElseIf Double.TryParse(textbox.Text, doubleAdultPriceBus) AndAlso doubleAdultPriceBus = 0 Then
                    childwoBedPriceBusTextBox(indexOfadultPriceBus).Text = ""
                End If
            End If
        Next
    End Sub
    Private Sub Update_Profit_Costing()
        Dim travelNecessitiesTextBox() As TextBox = {visaPrice, bulletTrain, luggagePrice, bukuUmrah, tagLanyard, mutawifMy, tourLeaderMy, zamZam, agentComm, officePrice, miscPrice1, miscPrice2}
        Dim makAccOvrPriceRmTextbox() As TextBox = {ovrTotalRm2, ovrTotalRm3, ovrTotalRm4, ovrTotalRm5}
        Dim madAccOvrPriceRmTextbox() As TextBox = {ovrTotalMadRm2, ovrTotalMadRm3, ovrTotalMadRm4, ovrTotalMadRm5}
        Dim paxEcoTextBox() As TextBox = {adultPaxEco2, childwBedPaxEco2, childwoBedPaxEco2, adultPaxEco3, childwBedPaxEco3, childwoBedPaxEco3, adultPaxEco4, childwBedPaxEco4, childwoBedPaxEco4, adultPaxEco5, childwBedPaxEco5, childwoBedPaxEco5}
        Dim paxBusTextBox() As TextBox = {adultPaxBus2, childwBedPaxBus2, childwoBedPaxBus2, adultPaxBus3, childwBedPaxBus3, childwoBedPaxBus3, adultPaxBus4, childwBedPaxBus4, childwoBedPaxBus4, adultPaxBus5, childwBedPaxBus5, childwoBedPaxBus5}
        Dim priceEcoTextBox() As TextBox = {adultPriceEco2, childwBedPriceEco2, childwoBedPriceEco2, adultPriceEco3, childwBedPriceEco3, childwoBedPriceEco3, adultPriceEco4, childwBedPriceEco4, childwoBedPriceEco4, adultPriceEco5, childwBedPriceEco5, childwoBedPriceEco5}
        Dim priceBusTextBox() As TextBox = {adultPriceBus2, childwBedPriceBus2, childwoBedPriceBus2, adultPriceBus3, childwBedPriceBus3, childwoBedPriceBus3, adultPriceBus4, childwBedPriceBus4, childwoBedPriceBus4, adultPriceBus5, childwBedPriceBus5, childwoBedPriceBus5}
        Dim peakSeasonEcoTextBox() As TextBox = {peakEco2, peakEco3, peakEco4, peakEco5}
        Dim peakSeasonBusTextBox() As TextBox = {peakBus2, peakBus3, peakBus4, peakBus5}
        Dim bufferCosting(3) As Double
        Dim totalEcoCost As Double = 0
        Dim totalBusCost As Double = 0
        Dim totalEcoSale As Double = 0
        Dim totalBusSale As Double = 0
        Dim totalEcoProfit As Double = 0
        Dim totalBusProfit As Double = 0
        Dim grandTotalCost As Double = 0
        Dim grandTotalSale As Double = 0
        Dim grandTotalProfit As Double = 0

        For Each bufferCost As Double In bufferCosting
            Dim indexOfBufferCost As Integer = Array.IndexOf(bufferCosting, bufferCost)
            Dim doubleSaudiArrPriceRm, doubleTourLeadPriceRm, doubleMakPriceRm, doubleMadPriceRm As Double
            Dim singlePaxCost As Double = 0

            For Each textbox As TextBox In travelNecessitiesTextBox
                Dim strTravelNec As String = textbox.Text.Trim()

                If String.IsNullOrEmpty(strTravelNec) Then
                    Continue For
                End If

                Dim doubleTravelNec As Double

                If Double.TryParse(textbox.Text, doubleTravelNec) AndAlso doubleTravelNec >= 0 Then
                    singlePaxCost += doubleTravelNec
                End If
            Next

            If Double.TryParse(saudiArrPriceRm.Text, doubleSaudiArrPriceRm) Then
                singlePaxCost += doubleSaudiArrPriceRm
            End If

            If Double.TryParse(tourLeadPriceRm.Text, doubleTourLeadPriceRm) Then
                singlePaxCost += doubleTourLeadPriceRm
            End If

            If Double.TryParse(makAccOvrPriceRmTextbox(indexOfBufferCost).Text, doubleMakPriceRm) Then
                singlePaxCost += doubleMakPriceRm
            End If

            If Double.TryParse(madAccOvrPriceRmTextbox(indexOfBufferCost).Text, doubleMadPriceRm) Then
                singlePaxCost += doubleMadPriceRm
            End If

            bufferCosting(indexOfBufferCost) = singlePaxCost
        Next

        For Each textbox As TextBox In paxEcoTextBox
            Dim strPaxEco As String = textbox.Text.Trim()
            Dim indexOfPaxEco As Integer = Array.IndexOf(paxEcoTextBox, textbox)
            Dim intPaxEco As Integer

            If String.IsNullOrEmpty(strPaxEco) Or Not Integer.TryParse(textbox.Text, intPaxEco) Then
                Continue For
            End If

            Dim arrayQuotient As Integer = indexOfPaxEco \ 3
            Dim doublePriceEco, doubleTicketEco, doublePeakSeasonEco As Double
            Dim strTicketEco As String = ticketPriceEco.Text.Trim()
            Dim strPriceEco As String = priceEcoTextBox(indexOfPaxEco).Text.Trim()

            If Double.TryParse(priceEcoTextBox(indexOfPaxEco).Text, doublePriceEco) AndAlso doublePriceEco >= 0 Then
                totalEcoSale += doublePriceEco * intPaxEco
                totalEcoCost += bufferCosting(arrayQuotient) * intPaxEco

                If Double.TryParse(ticketPriceEco.Text, doubleTicketEco) AndAlso doubleTicketEco >= 0 Then
                    totalEcoCost += doubleTicketEco * intPaxEco
                End If

                If Double.TryParse(peakSeasonEcoTextBox(arrayQuotient).Text, doublePeakSeasonEco) AndAlso doublePeakSeasonEco >= 0 Then
                    totalEcoCost += doublePeakSeasonEco * intPaxEco
                End If
            End If
        Next

        For Each textbox As TextBox In paxBusTextBox
            Dim strPaxBus As String = textbox.Text.Trim()
            Dim indexOfPaxBus As Integer = Array.IndexOf(paxBusTextBox, textbox)
            Dim intPaxBus As Integer

            If String.IsNullOrEmpty(strPaxBus) Or Not Integer.TryParse(textbox.Text, intPaxBus) Or intPaxBus < 0 Then
                Continue For
            End If

            Dim arrayQuotient As Integer = indexOfPaxBus \ 3
            Dim doublePriceBus, doubleTicketBus, doublePeakSeasonBus As Double
            Dim strTicketBus As String = ticketPriceBus.Text.Trim()
            Dim strPriceBus As String = priceBusTextBox(indexOfPaxBus).Text.Trim()

            If Double.TryParse(priceBusTextBox(indexOfPaxBus).Text, doublePriceBus) AndAlso doublePriceBus >= 0 Then
                totalBusSale += doublePriceBus * intPaxBus
                totalBusCost += bufferCosting(arrayQuotient) * intPaxBus

                If Double.TryParse(ticketPriceBus.Text, doubleTicketBus) AndAlso doubleTicketBus >= 0 Then
                    totalBusCost += doubleTicketBus * intPaxBus
                End If

                If Double.TryParse(peakSeasonBusTextBox(arrayQuotient).Text, doublePeakSeasonBus) AndAlso doublePeakSeasonBus >= 0 Then
                    totalBusCost += doublePeakSeasonBus * intPaxBus
                End If
            End If
        Next

        totalBusProfit = totalBusSale - totalBusCost
        totalEcoProfit = totalEcoSale - totalEcoCost
        grandTotalCost = totalBusCost + totalEcoCost
        grandTotalSale = totalBusSale + totalEcoSale

        Dim doubleBabyPrice As Double
        Dim intBabyPax As Integer


        If Double.TryParse(babyPolicy.Text, doubleBabyPrice) AndAlso doubleBabyPrice >= 0 AndAlso Integer.TryParse(babyPax.Text, intBabyPax) AndAlso intBabyPax >= 0 Then
            grandTotalProfit = totalBusProfit + totalEcoProfit + (doubleBabyPrice * intBabyPax)
        Else
            grandTotalProfit = totalBusProfit + totalEcoProfit
        End If

        RichTextBox1.Text = ""
        RichTextBox1.SelectionColor = Color.Red
        RichTextBox1.SelectionAlignment = HorizontalAlignment.Center
        RichTextBox1.AppendText(totalEcoCost.ToString("N2").TrimEnd("0"c).TrimEnd("."c))
        RichTextBox1.SelectionColor = Color.Black
        RichTextBox1.AppendText(" / ")
        RichTextBox1.SelectionColor = Color.Blue
        RichTextBox1.AppendText(totalEcoSale.ToString("N2").TrimEnd("0"c).TrimEnd("."c))
        RichTextBox1.SelectionColor = Color.Black
        RichTextBox1.AppendText(" / ")
        RichTextBox1.SelectionColor = Color.Green
        RichTextBox1.AppendText(totalEcoProfit.ToString("N2").TrimEnd("0"c).TrimEnd("."c))

        RichTextBox2.Text = ""
        RichTextBox2.SelectionColor = Color.Red
        RichTextBox2.SelectionAlignment = HorizontalAlignment.Center
        RichTextBox2.AppendText(totalBusCost.ToString("N2").TrimEnd("0"c).TrimEnd("."c))
        RichTextBox2.SelectionColor = Color.Black
        RichTextBox2.AppendText(" / ")
        RichTextBox2.SelectionColor = Color.Blue
        RichTextBox2.AppendText(totalBusSale.ToString("N2").TrimEnd("0"c).TrimEnd("."c))
        RichTextBox2.SelectionColor = Color.Black
        RichTextBox2.AppendText(" / ")
        RichTextBox2.SelectionColor = Color.Green
        RichTextBox2.AppendText(totalBusProfit.ToString("N2").TrimEnd("0"c).TrimEnd("."c))

        RichTextBox3.Text = ""
        RichTextBox3.SelectionColor = Color.Red
        RichTextBox3.SelectionAlignment = HorizontalAlignment.Center
        RichTextBox3.AppendText(grandTotalCost.ToString("N2").TrimEnd("0"c).TrimEnd("."c))
        RichTextBox3.SelectionColor = Color.Black
        RichTextBox3.AppendText(" / ")
        RichTextBox3.SelectionColor = Color.Blue
        RichTextBox3.AppendText(grandTotalSale.ToString("N2").TrimEnd("0"c).TrimEnd("."c))
        RichTextBox3.SelectionColor = Color.Black
        RichTextBox3.AppendText(" / ")
        RichTextBox3.SelectionColor = Color.Green
        RichTextBox3.AppendText(grandTotalProfit.ToString("N2").TrimEnd("0"c).TrimEnd("."c))
    End Sub

End Class

