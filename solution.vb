' you can write to stdout for debugging purposes, e.g.
' Console.WriteLine("this is a debug message")
Module VBModule
    Sub Main()
        DIm A = New Integer() {9,3,9,3,9,7,9}
        Console.WriteLine("Odd Man Out: " & solution(A))
    End Sub
        
    private function solution(A as Integer()) as Integer
        private dim numElements as Integer

        numElements = A.Length

        Dim alValue as ArrayList
        Dim alCount as ArrayList
        Dim i as Integer
        Dim j as Integer

        dim num1 as Integer
        dim num2 as Integer

        dim indexNum as Integer

        j = numElements

        for i = 1 to numElements/2 + 1
            'If current element is the middle element
            'A[i] will equal A[j]
            if i = j Then 
                num1 = A[i]

                if alValue.contains (num1) Then 
                    indexNum = alValue.IndexOf(num1)
                    alCount.Item(indexNum) = alCount.Item(indexNum) + 1
                Else
                    alValue.Add(num1)
                    alCount.Add(1)
                End If
            
            else 
                'get elements for first and last of A[]
                num1 = A[i]
                num2 = A[j]

                'if numbers match only search once
                'and increment count by 2
                if num1 = num2 Then

                    if alValue.contains (num1) Then 
                        indexNum = alValue.IndexOf(num1)
                        alCount.Item(indexNum) = alCount.Item(indexNum) + 2
                    Else
                        alValue.Add(num1)
                        alCount.Add(2)
                    End If
                Else 'num1 != num2
                    'num1
                    if alValue.contains (num1) Then 
                        indexNum = alValue.IndexOf(num1)
                        alCount.Item(indexNum) = alCount.Item(indexNum) + 1
                    Else
                        alValue.Add(num1)
                        alCount.Add(1)
                    End If
                    'num2
                    if alValue.contains (num2) Then 
                        indexNum = alValue.IndexOf(num2)
                        alCount.Item(indexNum) = alCount.Item(indexNum) + 1
                    Else
                        alValue.Add(num2)
                        alCount.Add(1)
                    End If
                End If
                
                ' Decrease index J
                j = j - 1

            end if

        next i

        indexNum = alCount.IndexOf(1)

        solution = alValue.Item(indexNum)

    end function
end Module
