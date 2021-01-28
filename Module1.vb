Module Module1

    Sub Main()
        'cz-Zadejte ve smyčce do pole úspory několika osob. Poté se program zeptá, jaká je roční úroková míra. Pole předejte funkci,
        'která vrátí pole po započítání úroků.

        'en-Type in loop to the array savings of few people. After that program will ask you for "annual interest rate".
        'The array pass on to the function, which returns array increased by the interests.

        Dim SavingsArray(10) As Single
        Dim i As Integer
        Dim rpsn As Single
        Dim ret As String
        Dim j As Integer
        Dim SavingsWithInterestsArray() As Single

        i = 0
        Do
            SavingsArray(i) = InputBox("Type amount of savings of  some people.( End typing with type -1")
            i += 1
        Loop Until SavingsArray(i - 1) = -1
        rpsn = InputBox("Type annual interest rate in %.")

        SavingsWithInterestsArray = withInterestsF(i, rpsn, SavingsArray)
        'here is array without brackets

        ret = ""
        For j = 0 To i - 2
            ret = ret + Str(SavingsWithInterestsArray(j)) + Chr(10)
        Next
        MsgBox("Amouts of savings with interests of particular people are: " + Chr(10) + ret)
    End Sub

    Function withInterestsF(iiS As Integer, rpsnF As Integer, SavingsArrayF() As Single) As Single() 'these brackets on the end means that function returns array
        Dim k As Integer

        Dim SavingsWithInterestsArrayF(10) As Single

        For k = 0 To iiS - 2
            SavingsWithInterestsArrayF(k) = (SavingsArrayF(k) * (rpsnF / 100)) + SavingsArrayF(k)
        Next
        withInterestsF = SavingsWithInterestsArrayF
        'here are array without bracket too
    End Function

End Module
