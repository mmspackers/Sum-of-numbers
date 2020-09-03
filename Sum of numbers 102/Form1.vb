'*******************************************************
'Matthew Smith
'FA/20 IT 102 Programming 2
'Assignment 1 "Sum of Numbers"
'9/1/2020
'*******************************************************
Option Strict On


Public Class Form1

    Private Sub btnEnterNumbers_Click(sender As Object, e As EventArgs) Handles btnEnterNumbers.Click

        Dim strPositiveInteger As String
        Dim intCounter As Integer
        Dim intNumber As Integer
        Dim intSum As Integer


        strPositiveInteger = InputBox("Enter a positive integer value", "Input Needed", "10")


        If ValidInput(strPositiveInteger, intNumber) = True Then

            intSum = 0
            intCounter = 0

            Do Until intCounter > intNumber

                intSum += intCounter

                    intCounter += 1

            Loop


            MessageBox.Show("The Sum of 1 to " & intNumber & " is " & intSum)

        End If

    End Sub

    Function ValidInput(ByRef PositiveInteger As String, ByRef Number As Integer) As Boolean

        If ValidateInputBox(PositiveInteger) = True Then
            If ValidatePositiveInteger(Number, PositiveInteger) = True Then
                If ValidateNumber(Number) = True Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Else
            Return False
        End If

    End Function


    Function ValidateInputBox(ByRef PositiveInteger As String) As Boolean

        If PositiveInteger = "" Then

            MessageBox.Show("Please enter a value.")
            Return False

        End If

        Return True

    End Function


    Function ValidatePositiveInteger(ByRef Number As Integer, ByVal PositiveInteger As String) As Boolean

        If IsNumeric(PositiveInteger) Then

            Number = CInt(PositiveInteger)

        Else

            MessageBox.Show("Please enter a numeric Value.")

            Return False

        End If

        Return True
    End Function

    Function ValidateNumber(ByVal Number As Integer) As Boolean

        If Number < 1 Then

            MessageBox.Show("Please enter in a positive integer.")
            Return False

        End If

        Return True

    End Function


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btnExit.Click

        ' Closes the program 
        Close()

    End Sub

End Class
