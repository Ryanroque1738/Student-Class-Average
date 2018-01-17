Module Module1
    'Program:   Student Grade Average Chapter 5
    'Programmer:    Steven Gramlich
    'Date:  4/5/42017
    'Purpose:   Program requests midterm score and final score and computes average for each student and class.

    Sub Main()
        'Obtain Upperbound of arrays
        Dim intClassSize As Integer = 0
        getClassSize(intClassSize)
        intClassSize = intClassSize - 1 'Adjust class size to upperbound of array

        'Declerations
        Dim intMidterm() As Integer = New Integer(intClassSize) {}
        Dim intFinal() As Integer = New Integer(intClassSize) {}
        Dim strName() As String = New String(intClassSize) {}
        Dim dblGradeAverage() As Double = New Double(intClassSize) {}
        Dim strEvaluation() As String = New String(intClassSize) {}

        Dim dblClassAverage As Double = 0
        Dim strClassEvaluation As String = ""


        'Get Grades

        For intCount As Integer = 0 To intClassSize
            getStudentInfo(strName, intCount)
            getMidterm(intMidterm, intCount)
            getFinal(intFinal, intCount)
        Next

        'Calculate Averages andClass Average

        For IntCount As Integer = 0 To intClassSize
            calculateAverage(intMidterm, intFinal, dblGradeAverage, dblClassAverage, IntCount)
            standardPassFail(dblGradeAverage, strEvaluation, IntCount)
        Next

        calculateClassAverage(dblClassAverage, intClassSize)
        classPassFail(dblClassAverage, strClassEvaluation)

        'Display averages and Class Average

        displayTitles()

        For intCount As Integer = 0 To intClassSize
            displayAverage(strName, dblGradeAverage, strEvaluation, intCount)
        Next

        displayClassAverage(dblClassAverage, strClassEvaluation)
        terminateProgram()

    End Sub

    Private Sub getClassSize(ByRef intSize As Integer)

        Console.Write("Enter the number of students om the class. ")
        intSize = CInt(Console.ReadLine())
        Console.Clear()

    End Sub

    Private Sub getStudentInfo(ByRef strStudent() As String, ByVal intIndex As Integer)

        Console.Write("Enter the name of the student. ")
        strStudent(intIndex) = Console.ReadLine()

    End Sub

    Private Sub getMidterm(ByRef intgrade() As Integer, ByVal intIndex As Integer)

        Console.Write("What is the student's midterm score? ")
        intgrade(intIndex) = CInt(Console.ReadLine())

    End Sub

    Private Sub getFinal(ByRef intGrade() As Integer, ByVal intIndex As Integer)

        Console.Write("What is the student's final score? ")
        intGrade(intIndex) = CInt(Console.ReadLine())
        Console.Clear()

    End Sub

    Private Sub calculateAverage(ByRef intGrade1() As Integer, ByRef intGrade2() As Integer, ByRef dblAverage() As Double,
                                 ByRef dblClassTotal As Double, ByVal intIndex As Integer)

        dblAverage(intIndex) = CDbl((intGrade1(intIndex) + intGrade2(intIndex)) / 2)
        dblClassTotal = dblClassTotal + dblAverage(intIndex)

    End Sub

    Private Sub standardPassFail(ByRef dblAverage() As Double, ByRef strEval() As String, ByVal intIndex As Integer)

        If dblAverage(intIndex) >= 90 Then
            strEval(intIndex) = "A"
        ElseIf dblAverage(intIndex) >= 80 Then
            strEval(intIndex) = "B"
        ElseIf dblAverage(intIndex) >= 70 Then
            strEval(intIndex) = "C"
        ElseIf dblAverage(intIndex) >= 60 Then
            strEval(intIndex) = "D"
        Else
            strEval(intIndex) = "F"
        End If

    End Sub

    Private Sub calculateClassAverage(ByRef dblClassTotal As Double, ByVal intSize As Integer)

        dblClassTotal = dblClassTotal / (intSize + 1) ' Adjust upperbound of arrayu back to class size

    End Sub

    Private Sub classPassFail(ByVal dblAverage As Double, ByRef strEval As String)

        If dblAverage >= 90 Then
            strEval = "A"
        ElseIf dblAverage >= 80 Then
            strEval = "B"
        ElseIf dblAverage >= 70 Then
            strEval = "C"
        ElseIf dblAverage >= 60 Then
            strEval = "D"
        Else
            strEval = "F"
        End If

    End Sub

    Private Sub displayTitles()

        Console.WriteLine("Polk County Community College English Department")
        Console.WriteLine()
        Console.Write("Student".PadRight(25))
        Console.Write("Average".PadLeft(10))
        Console.WriteLine("Grade".PadLeft(10))
        Console.WriteLine("------------------------------------------------------")

    End Sub

    Private Sub displayAverage(ByRef strStudent() As String, ByRef dblAverage() As Double, ByRef strEval() As String,
                               ByVal intIndex As Integer)


        Console.Write(strStudent(intIndex).PadRight(20))
        Console.Write(" " & dblAverage(intIndex).ToString.PadLeft(10) & "%")
        Console.WriteLine(" " & strEval(intIndex).ToString.PadLeft(10))

    End Sub

    Private Sub displayClassAverage(ByVal dblAverage As Double, ByVal strEval As String)

        Console.WriteLine()
        Console.WriteLine("The class average is " & dblAverage & "%")
        Console.WriteLine("The class evaulation is " & strEval)

    End Sub

    Private Sub terminateProgram()

        Console.WriteLine()
        Console.WriteLine("Press the eneter key to terminate the program.")
        Console.WriteLine()
    End Sub

End Module
