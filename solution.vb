' you can write to stdout for debugging purposes, e.g.
' Console.WriteLine("this is a debug message")
Imports System

Module VBModule
    Sub Main()
        Dim A = New Integer() {9, 3, 9, 3, 9, 7, 9}
        Console.WriteLine("Odd Man Out: " & solution(A))
        Console.Read()
    End Sub

    Private Function solution(A As Integer()) As Integer
        Dim numElements As Integer

        Dim alValue As New ArrayList
        Dim alCount As New ArrayList
        Dim i As Integer
        Dim j As Integer

        Dim num1 As Integer
        Dim num2 As Integer

        Dim indexNum As Integer

        numElements = A.Length

        j = numElements - 1

        For i = 0 To j / 2
            'If current element is the middle element
            'A[i] will equal A[j]
            If i = j Then
                num1 = A(i)

                If alValue.Contains(num1) Then
                    indexNum = alValue.IndexOf(num1)
                    alCount.Item(indexNum) = alCount.Item(indexNum) + 1
                Else
                    alValue.Add(num1)
                    alCount.Add(1)
                End If

            Else
                'get elements for first and last of A[]
                num1 = A(i)
                num2 = A(j)

                'if numbers match only search once
                'and increment count by 2
                If num1 = num2 Then

                    'if the arraylist is empty
                    'add the first number to the list
                    If alValue Is Nothing Then
                        alValue.Add(num1)
                        alCount.Add(2)
                    ElseIf alValue.Contains(num1) Then
                        indexNum = alValue.IndexOf(num1)
                        alCount.Item(indexNum) = alCount.Item(indexNum) + 2
                    Else
                        alValue.Add(num1)
                        alCount.Add(2)
                    End If
                Else 'num1 != num2
                    'num1

                    'if the arraylist is empty
                    'add the first number to the list
                    If alValue Is Nothing Then
                        alValue.Add(num1)
                        alCount.Add(1)
                    ElseIf alValue.Contains(num1) Then
                        indexNum = alValue.IndexOf(num1)
                        alCount.Item(indexNum) = alCount.Item(indexNum) + 1
                    Else
                        alValue.Add(num1)
                        alCount.Add(1)
                    End If

                    'num2

                    'if the arraylist is empty
                    'add the first number to the list
                    If alValue Is Nothing Then
                        alValue.Add(num1)
                        alCount.Add(1)
                    ElseIf alValue.Contains(num2) Then
                        indexNum = alValue.IndexOf(num2)
                        alCount.Item(indexNum) = alCount.Item(indexNum) + 1
                    Else
                        alValue.Add(num2)
                        alCount.Add(1)
                    End If
                End If

                ' Decrease index J
                j = j - 1

            End If

        Next i

        indexNum = alCount.IndexOf(1)

        solution = alValue.Item(indexNum)

    End Function
End Module
