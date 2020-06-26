Attribute VB_Name = "helpFunctions"


Function mapClubToCity(ByRef club As String, ByRef cityArray() As String) As String

    Select Case club
    '1
        Case Is = "Bayern Munich"
            mapClubToCity = cityArray(0)
    '2
        Case Is = "Ein Frankfurt"
            mapClubToCity = cityArray(2)
    '3
        Case Is = "Wolfsburg"
            mapClubToCity = cityArray(6)
    '4
        Case Is = "RB Leipzig"
            mapClubToCity = cityArray(16)
    '5
        Case Is = "Mainz"
            mapClubToCity = cityArray(5)
    '6
        Case Is = "FC Koln"
            mapClubToCity = cityArray(12)
    '7
        Case Is = "Hoffenheim"
            mapClubToCity = cityArray(10)
    '8
        Case Is = "Hertha"
            mapClubToCity = cityArray(1)
    '9
        Case Is = "Paderborn"
            mapClubToCity = cityArray(9)
    '10
        Case Is = "Dortmund"
            mapClubToCity = cityArray(3)
    '11
        Case Is = "Augsburg"
            mapClubToCity = cityArray(11)
    '12
        Case Is = "Fortuna Dusseldorf"
            mapClubToCity = cityArray(13)
    '13
        Case Is = "Schalke 04"
            mapClubToCity = cityArray(4)
    '14
        Case Is = "Leverkusen"
            mapClubToCity = cityArray(7)
    '15
        Case Is = "Union Berlin"
            mapClubToCity = cityArray(1)
    '16
        Case Is = "Werder Bremen"
            mapClubToCity = cityArray(14)
    '17
        Case Is = "Freiburg"
            mapClubToCity = cityArray(8)
    '18
        Case Is = "M'gladbach"
            mapClubToCity = cityArray(15)
    '19
        Case Is = "Hannover"
            mapClubToCity = cityArray(18)
    '20
        Case Is = "St Pauli"
            mapClubToCity = cityArray(17)
    '21
        Case Is = "Stuttgart"
            mapClubToCity = cityArray(20)
    '22
        Case Is = "Nurnberg"
            mapClubToCity = cityArray(21)
    '23
        Case Is = "Hamburg"
            mapClubToCity = cityArray(17)
    '24
        Case Is = "Ingolstadt"
            mapClubToCity = cityArray(19)
    '25
        Case Is = "Darmstadt"
            mapClubToCity = cityArray(25)
    '26
        Case Is = "Braunschweig"
            mapClubToCity = cityArray(22)
    '27
        Case Is = "Greuther Furth"
            mapClubToCity = cityArray(26)
    '28
        Case Is = "Bochum"
            mapClubToCity = cityArray(23)
    '29
        Case Is = "Kaiserslautern"
            mapClubToCity = cityArray(24)
    '30
        Case Is = "Bielefeld"
            mapClubToCity = cityArray(27)
    '31
        Case Is = "Karlsruhe"
            mapClubToCity = cityArray(28)
    '32
        Case Is = "Cottbus"
            mapClubToCity = cityArray(29)
            
            
    End Select
    
End Function


Function roundTime(ByRef tmpDate As Date) As Date

    Select Case CStr(tmpDate)
        Case Is = "13:00:00"
            roundTime = CDate("12:00:00")
        
        Case Is = "13:30:00"
            roundTime = CDate("12:00:00")
        
        Case Is = "15:30:00"
            roundTime = CDate("15:00:00")
            
        Case Is = "18:00:00"
            roundTime = CDate("18:00:00")
        
        Case Is = "18:30:00"
            roundTime = CDate("18:00:00")
        
        Case Is = "20:00:00"
            roundTime = CDate("21:00:00")
            
        Case Is = "20:30:00"
            roundTime = CDate("21:00:00")
        
        
       End Select
        



End Function

