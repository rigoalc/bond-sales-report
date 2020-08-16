
Module AlcoverRBondsReport
    '                                                  START OF PROGRAM


    Private BondReportFile As New Microsoft.
        VisualBasic.FileIO.TextFieldParser("BONDFILE19.TXT") 'FILE NAME FOR THE PRINCIPAL RECORDS
    Private CityNamesFile As New Microsoft.
        VisualBasic.FileIO.TextFieldParser("CITYNAMES19.TXT") 'FILE NAME FOR THE CITY NAMES
    Private AccountNumberFile As New Microsoft.
        VisualBasic.FileIO.TextFieldParser("ACCOUNTS19.TXT") 'FILE NAME FOR THE ACCOUNTS #

    Private CurrentRecord() As String  ' CURRENT RECORDS
    '                                             NOW WE'LL DECLARE THE FILE 
    '                                       WE USE IN THE PROGRAM AND ASSOICIATE IT 
    '                          WITH THE ACTUAL FILE NAME, WHERE THE DATA IS STORED
    '                                                    INPUT VARIABLES/FIELDS:
    Private FederalIDInteger As Integer
    Private StateIDInteger As Integer
    Private CityIDInteger As Integer
    Private AccountString As String
    Private QuantitySoldInteger As Integer
    Private PrincipleValueDecimal As Decimal


    '                                                         CALCULATED FIELDS:
    Private TotalValueDecimal As Decimal  '                  


    '                                                         ACUMULATED FIELDS
    Private AccumCityTotalDecimal As Decimal = 0
    Private AccumStateTotalDecimal As Decimal = 0
    Private AccumFederalTotalDecimal As Decimal = 0
    Private AccumFinalTotalBondsDecimal As Decimal = 0

    '                                                        HOLD FIELDS:
    Private CityHoldInteger As Integer
    Private StateHoldInteger As Integer
    Private FederalHoldInteger As Integer

    '                                             PAGINATION VARIABLES:

    Private LineCounterInteger As Integer = 99         '      99 FOR HEADINGS ON FIRST PAGE
    '                                                              
    Private Const PAGE_SIZE_INTEGER As Integer = 30
    '                                                              
    Private PageNumberInteger As Integer = 1 '             PAGE #'S FOR HEADINGS            
    '                                                      FILE RECORD AND FILE NAME DECLARATIONS:
    '                                                      WHEN THE FILE IS READ, 
    '                                                      THE RECORD IS PLACED IN THIS VARIABLE
    '                                                      ASSIGNED FIELDS FROM DESICIONS

    '                                           CREATE A VARIABLE FOR THE SUBSCRIPT TO BE USED FOR LOOPING:

    '                                          HOLD ACCOUNT#'S FOR SEARCHING TO VALIDATE #'S FROM FILE
    Private MaxNumberofAccountNumberInteger As Integer = 20 '         WORKING VARIABLES
    Private SubscriptInteger As Integer
    Private SearchAccountBoolean As Boolean = True
    Private FirstRecordBoolean As Boolean = True
    Private AccountErrorString As String

    '                                                  CREATE A TABLE TO HOLD THE PAY RATES AND
    '                                                    ONE FOR THE EMPLOYEES SSN'S:
    'NAME OF THE TABLE TO HOLD THE 10 'STRING' CITY NAMES
    Dim CityNamesTable(9) As String
    Dim AccountNumberTable(MaxNumberofAccountNumberInteger - 1) As String

    Sub Main()   '                                         PROGRAM EXECUTION LOGIC STARTS.
        Call HouseKeeping()
        Do While Not BondReportFile.EndOfData
            Call ProcessRecords()
        Loop
        Call EndOfJob()
    End Sub

    Private Sub HouseKeeping()  '                          LEVEL 2 CONTROL MODULES
        Call SetFileDelimiter()
        Call LoadCityNameTable()
        Call LoadAcountNumberTable()
    End Sub

    Private Sub ProcessRecords() '  PROCESSRECORD, HERE WE ADD CONTROL BREAK                        
        Call ReadFile()
        Call ControlBreak()
        Call DetailCalculations()
        Call AccumulationTotals()
        Call SearchForAccountNumber()
        Call WriteDetailLine()
    End Sub


    Private Sub EndOfJob()
        Call CityTotalOutput()
        Call StateTotalOutput()
        Call FederalTotalOutput()
        Call FinalTotalOutput()
        Call CloseFile()
    End Sub


    'HOUSEKEEPLING MODULES
    'LOADS ARRAY / TABLES BY READING EACH RECORD (VALUE) FROM FILE AND MOVING THE DATA TO THE TABLES
    'LOADS THE PAY RATES INTO PayRateTable USING A 'DEFINITE ITERATION' FOR LOOP
    'USED WHEN # OF VALUES READ IS ALWAYS THE SAME 
    Private Sub SetFileDelimiter()    '           DEFINES FILES AS A DELIMITER
        '                                         DEFINES DELIMITER AS A COMMA        
        BondReportFile.TextFieldType = FileIO.FieldType.Delimited

        BondReportFile.SetDelimiters(",")

    End Sub

    Private Sub LoadCityNameTable()
        SubscriptInteger = 0
        ' DEFINITE ITERATION LOOP OF 10 TIMES FROM FILE DIRECT TO THE ARRAY 
        For SubscriptInteger = 0 To 9 'PREPARATION
            CityNamesTable(SubscriptInteger) = CityNamesFile.ReadLine() 'CREATE
            ' INPUT CITYNAME...CURRENT RECORD IS N/A SINCE ONLY 1 FIELD ALTHOUGH
        Next
        'For SubscriptInteger = 0 To 9 ' TEST CODE TO VALIDATE
        'THIS TEST Is Not IN THE FLOWCHART
        'Console.WriteLine(SubscriptInteger & " " & CityNamesTable(SubscriptInteger))
        'Next
    End Sub

    Private Sub LoadAcountNumberTable()
        SubscriptInteger = 0
        'INDEFINITE ITERATION LOOP 
        Do While Not AccountNumberFile.EndOfData
            AccountString = AccountNumberFile.ReadLine() 'READ THE FILE
            AccountNumberTable(SubscriptInteger) = AccountString 'MOVE TO TABLE
            SubscriptInteger += 1 'ADD ONE 
        Loop
        ' TEST THE TABLE
        MaxNumberofAccountNumberInteger = SubscriptInteger - 1 'PAIR SUBSCRIPT SUBSTRACTING 1
        'For SubscriptInteger = 0 To 9 ' TEST THE TABLE NOT IN FLOWCHART :)
        ' Console.WriteLine(SubscriptInteger & " " & AccountNumberTable(SubscriptInteger))
        'Next
    End Sub

    Private Sub ReadFile() '  READ WHOLE RECORD AND ASSIGN TO THE CURRENT RECORD VARIABLE
        CurrentRecord = BondReportFile.ReadFields()
        FederalIDInteger = CurrentRecord(0)
        StateIDInteger = CurrentRecord(1) '            PLACE CURRENT RECORDS FIELDS 
        CityIDInteger = CurrentRecord(2)
        AccountString = CurrentRecord(3)
        QuantitySoldInteger = CurrentRecord(4)
        PrincipleValueDecimal = CurrentRecord(5)
        If FirstRecordBoolean = True Then   'RECORD BOOLEAN FOR ASIGN THE HOLD VARIABLES
            Call HoldCitySetUp()
            Call HoldStateSetUp()
            Call HoldFederalSetUp()
            FirstRecordBoolean = False
        End If
    End Sub


    Private Sub HoldCitySetUp() ' COMPARE FOR CITY VARIABLE
        CityHoldInteger = CityIDInteger
    End Sub


    Private Sub HoldStateSetUp() ' COMPARE FOR STATE VARIABLE
        StateHoldInteger = StateIDInteger
    End Sub


    Private Sub HoldFederalSetUp() ' COMPARE FOR FEDERAL VARIABLE
        FederalHoldInteger = FederalIDInteger
    End Sub


    Private Sub ControlBreak() ' NEW MODULE CONTROLE BREAK COMPARE AND ORGANIZE THE CITYS, STATES OR FEDERAL
        If FederalIDInteger <> FederalHoldInteger Then
            Call CityTotalOutput()
            Call StateTotalOutput()
            Call FederalTotalOutput()
            LineCounterInteger = LineCounterInteger + 8 'ADDING LINES FOR MORE SPACES UTILIZED
        Else
            If StateIDInteger <> StateHoldInteger Then
                Call CityTotalOutput()
                Call StateTotalOutput()
                LineCounterInteger = LineCounterInteger + 6 'SPACES IN THIS CASE

            Else
                If CityIDInteger <> CityHoldInteger Then
                    Call CityTotalOutput()
                    LineCounterInteger = LineCounterInteger + 4 'FOUR HERE
                End If
            End If
        End If
    End Sub

    Private Sub DetailCalculations()   '                   CALCULATED TOTAL VALUE
        TotalValueDecimal = QuantitySoldInteger * PrincipleValueDecimal
    End Sub

    Private Sub AccumulationTotals()  '                      ACCUMULATE  TOTALS 
        AccumCityTotalDecimal = AccumCityTotalDecimal + TotalValueDecimal
    End Sub

    Private Sub SearchForAccountNumber()
        SearchAccountBoolean = False ' SET BOOLEAN FALSE
        SubscriptInteger = 0 'START IN LOCATION 0
        Do While SubscriptInteger <= MaxNumberofAccountNumberInteger 'LOOP
            Call AccountNumberTest() 'CALL THE ACTUAL TEST
            SubscriptInteger = SubscriptInteger + 1 'ADD 1 TO SUBSCRIPT
        Loop
    End Sub


    Private Sub AccountNumberTest()
        If AccountString = AccountNumberTable(SubscriptInteger) Then 'COMPARE THE TABLE WITH THE STRING
            SearchAccountBoolean = True
            SubscriptInteger = MaxNumberofAccountNumberInteger
        End If
    End Sub


    Private Sub WriteDetailLine()
        '                                          WRITE DETAIL LINE
        If LineCounterInteger >= PAGE_SIZE_INTEGER Then 'LINE COUNTERCOMPARE TO PAGE SIZE
            Call WriteHeadings()
        End If

        If SearchAccountBoolean = False Then

            AccountString = "ERROR*"

        End If
        Console.WriteLine(Space(5) & FederalIDInteger.ToString.PadLeft(3) &
                          Space(6) & StateIDInteger.ToString.PadLeft(2) &
                          Space(2) & CityNamesTable(CityHoldInteger - 1).ToString.PadRight(9) &
                          Space(2) & AccountString.ToString.PadRight(6) &
                          Space(4) & PrincipleValueDecimal.ToString("N").PadLeft(8) &
                          Space(8) & QuantitySoldInteger.ToString.PadLeft(4) &
                          Space(6) & TotalValueDecimal.ToString("C").PadLeft(14))
        '                                     LineCounterInteger = LineCounterInteger +1    
        '                                     COUNT THE LINE PRINTED
        LineCounterInteger += 1 '             +=  IS A ' COMBINED OPERATOR'
        '                                     SHORTCUT FOR ACCUMULATION
        '                                     OUTPUT 1 LINE FOR EACH PERSON PROCESSED 
        '                                     TEST FOR PAGINATION


    End Sub


    Private Sub WriteHeadings()
        '                             WRITE HEADINGS MODULE IS PART OF PROCESS RECORD MODULES
        '                             AND IS CALL BY WRITE DETAILLINE WEN THE LINE COUNTER 
        '                             IS GREATER OR EQUAL TO 25.
        '                             WRITE REPORTHEADLINES
        Console.WriteLine(Space(18) & "Federal, State and City Bond Value Report" &
                          Space(12) & "Page " & PageNumberInteger.ToString("n0").PadLeft(3))
        Console.WriteLine(Space(33) & "Rodrigo Alcover")
        Console.WriteLine()                          'WRITE COLUMN LEADER LINES
        Console.WriteLine(Space(1) &
                          "Federal" & Space(3) &
                          "State" & Space(2) &
                          "City" & Space(5) &
                          "Account" & Space(4) &
                          "Principle" & Space(4) &
                          "Quantity" & Space(15) &
                          "Total")
        Console.WriteLine(Space(4) &
                          "ID #" & Space(4) &
                          "ID #" & Space(2) &
                          "ID #" & Space(6) &
                          "Number" & Space(8) &
                          "Value" & Space(8) &
                          "Sold" & Space(15) &
                          "Value" & Space(1))
        Console.WriteLine()
        LineCounterInteger = 7 '               RESET LINE COUNTER &
        PageNumberInteger += 1       '   ADD TO PAGE#     +=  IS CALLED A  COBINED OPERATOR
    End Sub


    Private Sub CityTotalOutput() ' WHRITE TOTALS END OF REPORT
        Console.WriteLine() 'AND ROLL THE TOTALS , AND SET TO 0 AFTER, AND CALL THE HOLD MODULE
        Console.WriteLine(Space(30) & "City " & CityNamesTable(CityHoldInteger - 1).ToString.PadLeft(9) &
                          Space(6) & "Total:" & Space(8) & AccumCityTotalDecimal.ToString("C").PadLeft(15) &
                          Space(1) & "*")
        Console.WriteLine()
        AccumStateTotalDecimal = AccumStateTotalDecimal + AccumCityTotalDecimal 'ROLL
        AccumCityTotalDecimal = 0 ' SET TO 0
        Call HoldCitySetUp()
    End Sub


    Private Sub StateTotalOutput()
        Console.WriteLine(Space(30) & "State ID " & StateHoldInteger.ToString.PadLeft(2) &
                          Space(9) & "Total:" & Space(8) & AccumStateTotalDecimal.ToString("C").PadLeft(15) &
                          Space(1) & "**")
        Console.WriteLine()
        AccumFederalTotalDecimal = AccumFederalTotalDecimal + AccumStateTotalDecimal
        AccumStateTotalDecimal = 0
        Call HoldStateSetUp() 'CALL HOLD MODULE
    End Sub


    Private Sub FederalTotalOutput()
        Console.WriteLine(Space(30) & "Federal ID " & FederalHoldInteger.ToString.PadLeft(3) &
                          Space(6) & "Total" & Space(8) & AccumFederalTotalDecimal.ToString("C").PadLeft(16) &
                          Space(1) & "***")
        AccumFinalTotalBondsDecimal = AccumFinalTotalBondsDecimal + AccumFederalTotalDecimal
        AccumFederalTotalDecimal = 0
        Call HoldFederalSetUp()
    End Sub


    Private Sub FinalTotalOutput() 'WRITE TOTAL FINAL
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine(Space(30) & "Final Total of All Bonds:" &
                          Space(7) & AccumFinalTotalBondsDecimal.ToString("C").PadLeft(17) &
                          Space(1) & "****")
    End Sub


    Private Sub CloseFile()                        ' END OF JOB MODULES
        BondReportFile.Close() '             CLOSING THE FILE
        CityNamesFile.Close()
        AccountNumberFile.Close()
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine()
        Console.WriteLine("Click ENTER Close Output Window")
        Console.ReadKey() '  WRITE MESSAGE FOR PRESS ENTER AND
        '                    CLOSE THE WINDOW PROMPT 
    End Sub
End Module

