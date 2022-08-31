Public Class ClsConversion
    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public Overrides Function ToString() As String
        Return MyBase.ToString()
    End Function

    Public Overrides Function Equals(obj As Object) As Boolean
        Return MyBase.Equals(obj)
    End Function

    Public Overrides Function GetHashCode() As Integer
        Return MyBase.GetHashCode()
    End Function

    Private workMinutes As Integer
    Public WriteOnly Property NewWorkMinutes() As Integer
        Set(value As Integer)
            workMinutes = value
        End Set
    End Property

    Private carryCriteria As Integer
    Public WriteOnly Property NewCarryOverCriteria() As Integer
        Set(value As Integer)
            carryCriteria = value
        End Set
    End Property

    Public Sub GetClassSettings()
        With My.Settings
            workMinutes = .TotalWorkMinutes
            carryCriteria = .CarryOverCriteria
        End With
    End Sub

    Public Sub SaveClassSettings()
        With My.Settings
            .TotalWorkMinutes = workMinutes
            .CarryOverCriteria = carryCriteria
            .Save()
        End With
    End Sub

    Public Function FractionsOfADay(ByVal Minutes As Integer) As Decimal
        FractionsOfADay = 0
        Try
            If workMinutes > 0 AndAlso carryCriteria > 0 Then
                ' Changes Minutes into decimal number
                Dim result As Decimal = Minutes / workMinutes * 10000
                ' Retains the whole number part
                Dim wholeNumberPart As Integer = Int(result)
                ' Converts result to string
                Dim numberAsText As String = Format(result, "#,##0.00")
                ' Gets the decimal part of numberAsText
                Dim decimalPart As String = Right(numberAsText, 2)
                ' Splits the decimal part
                Dim decimalLeft As String = Left(decimalPart, 1)
                Dim decimalRight As String = Right(decimalPart, 1)
                ' If decimalRight is greater than carryOver, adds 1 to decimalLeft
                Dim carryOver As Integer = IIf(Int(decimalRight) > carryCriteria, Int(decimalLeft) + 1, Int(decimalLeft))
                ' If carryOver is greater than carryOver, adds 1 to wholeNumberPart
                wholeNumberPart = IIf(carryOver > carryCriteria, wholeNumberPart + 1, wholeNumberPart)
                ' Second iteration (converts wholeNumberPart to number with one decimal place)
                result = wholeNumberPart / 10
                ' Retains the whole number part
                wholeNumberPart = Int(result)
                ' Converts result to string
                numberAsText = Format(result, "#,##0.0")
                ' Gets the decimal part of numberAsText
                decimalPart = Right(numberAsText, 1)
                ' If decimalPart is Greater than carryOver, adds one to wholeNumberPart
                wholeNumberPart = IIf(Int(decimalPart) > carryCriteria, wholeNumberPart + 1, wholeNumberPart)
                ' Returns wholeNumberPart as decimal number by dividing it by 1000
                FractionsOfADay = wholeNumberPart / 1000
            End If
        Catch ex As Exception

        End Try
        Return FractionsOfADay
    End Function
End Class
