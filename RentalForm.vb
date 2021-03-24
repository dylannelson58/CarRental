Option Explicit On
Option Strict On
Option Compare Binary

'Dylan Nelson
'RCET0265 
'Spring 2021
'Math Contest
'https://github.com/dylannelson58/CarRental
Public Class RentalForm

    Dim totalCharge As Decimal
    Dim numberOfDays As Integer
    Dim milesOrKilometers As Decimal
    Dim milesDriven As Decimal
    Dim mileageCharge As Decimal
    Dim daysCharge As Integer
    Dim discountV As String


    'Dim numberOfDays As Integer
    Private Sub CalculateButton_Click(sender As Object, e As EventArgs) Handles CalculateButton.Click
        InputValidation()
        Calculations()
        TotalChargeTextBox.Text = CStr($"${totalCharge}")
        TotalMilesTextBox.Text = CStr($"{milesDriven} mi")
        MileageChargeTextBox.Text = CStr($"${mileageCharge}")
        DayChargeTextBox.Text = CStr($"${daysCharge}")
        TotalDiscountTextBox.Text = discountV

    End Sub

    Public Function InputValidation() As String
        Dim problemMessage As String
        Dim numberOfDays As Integer
        Dim endOdometerReading As Integer
        Dim beginOdometerReading As Integer
        Dim zipCode As Integer
        'Dim state As Integer
        'Dim city As String
        'Dim address As String
        'Dim customerName As String

        Try
            numberOfDays = CInt(DaysTextBox.Text)
        Catch ex As Exception
            problemMessage &= "Days must be a number" & vbNewLine
            'Exit Function
        End Try

        Try
            endOdometerReading = CInt(EndOdometerTextBox.Text)
        Catch ex As Exception
            problemMessage &= "End odometer must be a number" & vbNewLine
        End Try

        Try
            beginOdometerReading = CInt(BeginOdometerTextBox.Text)
        Catch ex As Exception
            problemMessage &= "Begin odometer must be a number" & vbNewLine
        End Try

        Try
            zipCode = CInt(ZipCodeTextBox.Text)
        Catch ex As Exception
            problemMessage &= "Zip code must be a number" & vbNewLine
        End Try

        If StateTextBox.Text = "" Then
            problemMessage &= "State box cannot be empty" & vbNewLine
        End If

        If AddressTextBox.Text = "" Then
            problemMessage &= "Address cannot be empty" & vbNewLine
        End If

        If NameTextBox.Text = "" Then
            problemMessage &= "Name cannot be empty" & vbNewLine
        End If

        If CityTextBox.Text = "" Then
            problemMessage &= "City box cannot be empty" & vbNewLine
        End If

        If DaysTextBox.Text = "" Then
            problemMessage &= "Number of days box cannot be empty" & vbNewLine
        End If

        'If beginOdometerReading = CInt("") Then
        '    problemMessage &= "Beginning odometer reading cannot be empty" & vbNewLine
        'End If

        'If endOdometerReading = CInt("") Then
        '    problemMessage &= "Ending odometer reading cannot be empty" & vbNewLine
        'End If

        'If beginOdometerReading > endOdometerReading Then
        '    problemMessage &= "Beginning odometer reading cannot be less than ending odometer reading" & vbNewLine
        'End If



        If problemMessage = "" Then
        Else MsgBox(problemMessage)
            Exit Function

        End If

    End Function

    Public Function Calculations() As Decimal
        'Dim numberOfDays As Integer
        'Dim milesOrKilometers As Decimal
        'Dim milesDriven As Decimal
        'Dim mileageCharge As Decimal

        If MilesradioButton.Checked Then
            milesOrKilometers = 1
        ElseIf KilometersradioButton.Checked Then
            milesOrKilometers = CDec(1.60934)
        End If

        milesDriven = CDec(CDec(EndOdometerTextBox.Text) - CDec(BeginOdometerTextBox.Text)) * CDec(milesOrKilometers)

        numberOfDays = CInt(DaysTextBox.Text)
        daysCharge = numberOfDays * 15






        If milesDriven = 0 Then
            Exit Function
        ElseIf milesDriven < 200 Then
            mileageCharge = CDec((milesDriven) * 0)
        ElseIf milesDriven > 200 Then
            mileageCharge = CDec((milesDriven) * 1.12)
        ElseIf milesDriven > 500 Then
            mileageCharge = CDec((milesDriven) * 1.1)
        End If



        'numberOfDays = CInt(DaysTextBox.Text)
        If AAAcheckbox.Checked Then
            totalCharge = CDec((daysCharge + mileageCharge) * 0.95)
            discountV = CStr("5%")
        ElseIf Seniorcheckbox.Checked Then
            totalCharge = CDec((daysCharge + mileageCharge) * 0.97)
            discountV = CStr("3%")
        Else
            totalCharge = CDec(daysCharge + mileageCharge)
            discountV = CStr("0%")
        End If

    End Function

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click

        NameTextBox.Text = ""
        AddressTextBox.Text = ""
        CityTextBox.Text = ""
        StateTextBox.Text = ""
        ZipCodeTextBox.Text = ""
        BeginOdometerTextBox.Text = ""
        EndOdometerTextBox.Text = ""
        DaysTextBox.Text = ""
        TotalMilesTextBox.Text = ""
        MileageChargeTextBox.Text = ""
        DayChargeTextBox.Text = ""
        TotalDiscountTextBox.Text = ""
        TotalChargeTextBox.Text = ""



    End Sub

    Private Sub ExitButton_Click(sender As Object, e As EventArgs) Handles ExitButton.Click
        MsgBox("Are  you sure you want to exit?", CType(MessageBoxButtons.YesNo, MsgBoxStyle))
        If CBool(DialogResult.Yes) Then
            Me.Close()
        ElseIf CBool(DialogResult.no) Then

        End If
    End Sub

    Private Sub SummaryButton_Click(sender As Object, e As EventArgs) Handles SummaryButton.Click

    End Sub
End Class
