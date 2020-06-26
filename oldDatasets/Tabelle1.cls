VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub get_CityTime()
    Dim maxRowCount As Integer
    Dim maxColCount As Integer
    Dim localDate As Date
    Dim localTime As Date
    Dim homeTeam As String
    Dim homeCityArray(30) As String
    Dim i As Integer


    
    'Init Array with Cities
    homeCityArray(0) = "MŸnchen"
    homeCityArray(1) = "Berlin"
    homeCityArray(2) = "Frankfurt am Main"
    homeCityArray(3) = "Dortmund"
    homeCityArray(4) = "Gelsenkirchen"
    homeCityArray(5) = "Mainz"
    homeCityArray(6) = "Wolfsburg"
    homeCityArray(7) = "Leverkusen"
    homeCityArray(8) = "Freiburg"
    homeCityArray(9) = "Paderborn"
    homeCityArray(10) = "Sinsheim"
    homeCityArray(11) = "Augsburg"
    homeCityArray(12) = "Kšln"
    homeCityArray(13) = "DŸsseldorf"
    homeCityArray(14) = "Bremen"
    homeCityArray(15) = "Mšnchengladbach"
    homeCityArray(16) = "Leipzig"
    homeCityArray(17) = "Hamburg"
    homeCityArray(18) = "Hannover"
    homeCityArray(19) = "Ingolstadt"
    homeCityArray(20) = "Stuttgart"
    homeCityArray(21) = "NŸrnberg"
    homeCityArray(22) = "Braunschweig"
    homeCityArray(23) = "Bochum"
    homeCityArray(24) = "Kaiserslautern"
    homeCityArray(25) = "Darmstadt"
    homeCityArray(26) = "FŸrth"
    homeCityArray(27) = "Bielefeld"
    homeCityArray(28) = "Karlsruhe"
    homeCityArray(29) = "Cottbus"
    
    maxRowCount = Cells(Rows.count, 4).End(xlUp).Row
    maxColCount = Cells(1, Columns.count).End(xlToLeft).Column
    
    
    For i = 2 To maxRowCount
        Cells(i, 24) = mapClubToCity(Cells(i, 4), homeCityArray)
        localDate = Cells(i, 2)
        localTime = DateAdd("h", 1, Cells(i, 3))
        Debug.Print localTime
        Cells(i, 25) = roundTime(CDate(localTime))
    Next i
    
    

End Sub

Sub get_WeatherData()
    Dim WB1 As Workbook
    Dim WB2 As Workbook
    Set WB2 = Workbooks("allCities2.xlsm")
    Set WB1 = Workbooks("Merged_Data_Bundesliga 2009_2020.xlsm")
    Dim i As Integer
    Dim maxRow As Integer
    Dim d_Date As Date
    Dim d_time As Date
    Dim d_city As String
    Dim maxRowCity As Long
    Dim tmp As Integer
    Dim counter As Long
    Dim j As Integer
    
    maxRow = WB1.Worksheets("Data 2009_2020_Football").Cells(Rows.count, 4).End(xlUp).Row
    counter = 0
    For i = 2 To maxRow
        d_Date = WB1.Worksheets("Data 2009_2020_Football").Cells(i, 2)
        'Debug.Print d_Date
        d_city = WB1.Worksheets("Data 2009_2020_Football").Cells(i, 24)
        'Debug.Print d_city
        
        With WB2.Worksheets(d_city)
            maxRowCity = .Cells(Rows.count, 1).End(xlUp).Row
            'Debug.Print maxRowCity
            tmp = CInt(Right(CStr(d_Date), 4))
            'Debug.Print tmp
            If (tmp <= 2020 And (tmp > 13)) Then
                counter = maxRowCity
                While (d_Date <> .Cells(counter, 2))
                    counter = counter - 1
                Wend
                    
            Else
                counter = 2
                While (d_Date <> .Cells(counter, 2))
                    counter = counter + 1
                Wend

            End If
            
            For j = 4 To 27
                WB1.Worksheets("Data 2009_2020_Football").Cells(i, 25 - 4 + j) = .Cells(counter, j)
            Next
            
            
        End With
    Next i
    



End Sub







