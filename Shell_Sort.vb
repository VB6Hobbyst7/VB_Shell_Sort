'Justin Stachofsky
'Shell sort written for MIS 350
Module shellSort

    'Main subroutine for program
    Sub Main()

        Dim numberArray(100) 'Stores values to be sorted
        Dim userInput 'Stores user action
        Call generateValues(numberArray)

        Console.WriteLine("Unsorted Values:")
        Call printValues(numberArray)

        userInput = "a"
        While (userInput <> "q" Or userInput <> "s") 'Loop used to continually prompt menu if invalid command is given
            Console.WriteLine("Press 's' to sort or 'q' to exit program")
            userInput = Console.ReadLine()

            If userInput = "q" Then
                End
            ElseIf userInput = "s" Then
                Call sortValues(numberArray)
                Call printValues(numberArray)
                Console.ReadLine()
                End
            Else
                Console.WriteLine("Invalid command")
            End If
        End While

    End Sub

    'Subroutine used to fill array with values
    Sub generateValues(ByRef emptyArray)

        For i = 0 To UBound(emptyArray)
            emptyArray(i) = Int(Rnd() * 100) + 0 'Generates number between 0 and 100 and stores in array index
        Next

    End Sub

    'Subroutine used to print array
    Sub printValues(ByRef printArray)

        For i = 0 To UBound(printArray)
            Console.WriteLine(printArray(i))
        Next

    End Sub

    'Subroutine used to sort array, uses shell sort method
    Sub sortValues(ByRef unsortedArray)
        Dim swapValue 'Holds temporary value during sort
        Dim gapSize 'Gap between comparison items
        Dim recalculateGap 'Bool used to determine when it is time to recalculate gap

        gapSize = Int(UBound(unsortedArray) / 2) 'Force integer value

        Do While gapSize >= 1
            Do
                recalculateGap = True
                For i = 0 To (UBound(unsortedArray) - gapSize)
                    If unsortedArray(i) > unsortedArray(i + gapSize) Then
                        swapValue = unsortedArray(i)
                        unsortedArray(i) = unsortedArray(i + gapSize)
                        unsortedArray(i + gapSize) = swapValue
                        recalculateGap = False 'Set to false to make another pass with current gap size
                    End If
                Next
            Loop Until recalculateGap = True 'When condition is true gap is recalculated before next pass
            gapSize = Int(gapSize / 2) 'Force integer value
        Loop

    End Sub

End Module
