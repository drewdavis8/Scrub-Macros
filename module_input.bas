Function AlphaNumeric(pValue) As Boolean

   Dim LPos As Integer
   Dim LChar As String
   Dim LValid_Values As String

   'Start at first character in pValue
   LPos = 1

   'Set up values that are considered to be alphanumeric
   LValid_Values = " abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ+-,./&#0123456789'()\"

   'Test each character in pValue
   While LPos <= Len(pValue)

      'Single character in pValue
      LChar = Mid(pValue, LPos, 1)

      'If character is not alphanumeric, return FALSE
      If InStr(LValid_Values, LChar) = 0 Then
         AlphaNumeric = False
         Exit Function
      End If

      'Increment counter
      LPos = LPos + 1

   Wend

   'Value is alphanumeric, return TRUE
   AlphaNumeric = True

End Function


Sub Validation()
'
''''  Select the number of locations in the worksheet  ''''
   Dim numberOfLocations As Integer
    'Assign the variable
   numberOfLocations = WorksheetFunction.Count(Range("A2:A30000")) + 1
   
    Dim resetRng As Range
    Set resetRng = Range("A2:X" & numberOfLocations)

    For Each cell In resetRng
         cell.Interior.Color = RGB(255, 255, 255)
    Next cell
    
    
    ''''  Check If ZIP is a number  ''''

    Dim zipRng As Range
    Set zipRng = Range("G2:G" & numberOfLocations)
    
    For Each cell In zipRng
        If IsNumeric(cell.Value) = True Then cell.Interior.Color = RGB(255, 255, 255)
        If IsNumeric(cell.Value) = False Then cell.Interior.Color = RGB(200, 200, 200)
        If cell.Value = "" Then cell.Interior.Color = RGB(196, 228, 242)
    Next cell
    
''''  Check If Location Number is blank  ''''

    Dim locRng As Range
    Set locRng = Range("A2:A" & numberOfLocations)
    
    For Each cell In locRng
        If IsNumeric(cell.Value) = True Then cell.Interior.Color = RGB(255, 255, 255)
        If IsNumeric(cell.Value) = False Then cell.Interior.Color = RGB(200, 200, 200)
        If cell.Value = "" Then cell.Interior.Color = RGB(196, 228, 242)
    Next cell
    
''''  Check If Building Number is blank  ''''

    Dim buildRng As Range
    Set buildRng = Range("B2:B" & numberOfLocations)
    
    For Each cell In buildRng
        If IsNumeric(cell.Value) = True Then cell.Interior.Color = RGB(255, 255, 255)
        If IsNumeric(cell.Value) = False Then cell.Interior.Color = RGB(200, 200, 200)
        If cell.Value = "" Then cell.Interior.Color = RGB(196, 228, 242)
    Next cell
    
''''  Check If Num Buildings is blank  ''''

    Dim numRng As Range
    Set numRng = Range("C2:C" & numberOfLocations)
    
    For Each cell In numRng
        If IsNumeric(cell.Value) = True Then cell.Interior.Color = RGB(255, 255, 255)
        If IsNumeric(cell.Value) = False Then cell.Interior.Color = RGB(200, 200, 200)
        If cell.Value = "" Then cell.Interior.Color = RGB(196, 228, 242)
    Next cell

   
''''  Check that there are no blank addresses  ''''
   
       'Define Range
    Dim addressRng As Range
    'Set range to selection
    Set addressRng = Range("D2:D" & numberOfLocations)
    
    For Each cell In addressRng
        'If there is no value in the cell, leave it blank
        If cell.Value <> "" Then cell.Interior.Color = RGB(255, 255, 255)
        If cell.Value = "" Then cell.Interior.Color = RGB(196, 228, 242)
    'Continue to traverse selection
    Next cell
    
''''  Check Occupancy Validation  ''''
   
    'Define Range
    Dim occRng As Range
    'Set range to selection
    Set occRng = Range("I2:I" & numberOfLocations)
    
    For Each cell In occRng
    'Continue to traverse selection
        If cell.Value <> "" Then cell.Interior.Color = RGB(255, 255, 255)
        If cell.Value <> "Agricultural Products Mfg. (Cotton, Grains, Sugar)" And cell.Value <> "Airports" And cell.Value <> "Ammunition or Firearms Manufacturing" And cell.Value <> "Amusement Parks/State Fairgrounds" And cell.Value <> "Animal or Food Processing" And cell.Value <> "Bars/Taverns" And cell.Value <> "Bowling Alleys/Skating Rinks" And cell.Value <> "Bridges/Tunnels" And cell.Value <> "Builders Risk" And cell.Value <> "CBD/Hemp" And cell.Value <> "Camps" And cell.Value <> "Cannabis - Dispensaries" And cell.Value <> "Cannabis - Extraction/Processing" And cell.Value <> "Cannabis - Grow Facilities" And cell.Value <> "Casinos" And cell.Value <> "Chemical Processing/Mfg." And cell.Value <> "Clothing / Garment Mfg." And cell.Value <> "Communication Towers" And cell.Value <> "Dealer's Open Lot" And cell.Value <> "Drug Manufacturing" And cell.Value <> "Explosives or Fireworks Manufacturing or Distribution" And cell.Value <> "Fish Processing" Then cell.Interior.Color = RGB(1, 1, 1)
        If cell.Interior.Color = RGB(1, 1, 1) And cell.Value <> "Foundaries or Metal Manufacturing" And cell.Value <> "Fruit or Vegetable Processing" And cell.Value <> "Garage - Parking" And cell.Value <> "Grade Schools/Day Cares" And cell.Value <> "Heavy Mfg./Assembly" And cell.Value <> "High Tech / Clean Room Mfg." And cell.Value <> "High Value Homes" And cell.Value <> "Hospitals/Clinics" And cell.Value <> "Hotels/Motels" And cell.Value <> "LRO Cannabis" And cell.Value <> "Light Mfg./Assembly" And cell.Value <> "Minerals Processing/Mfg. (Cement, Bricks/Clay)" And cell.Value <> "Movie Theatres/Rec - Sport Clubs" And cell.Value <> "Municipalities / Government Services" And cell.Value <> "Museums / Libraries" And cell.Value <> "Nursing Homes" And cell.Value <> "Paper Products Manufacturing" And cell.Value <> "Penal Institutions/Jails" And cell.Value <> "Personal/Repair Services" Then cell.Interior.Color = RGB(2, 2, 2)
        If cell.Interior.Color = RGB(2, 2, 2) And cell.Value <> "Piers/Wharves/Docks" And cell.Value <> "Plastics Manufacturing" And cell.Value <> "Port Authorities" And cell.Value <> "Poultry/Egg Farm Processing" And cell.Value <> "Professional/Offices" And cell.Value <> "Radio/TV/Telephone Co. excl. Towers" And cell.Value <> "Recycling / Scrap Dealers" And cell.Value <> "Religion/Places of Worship" And cell.Value <> "Residential - Apts/Condos/Assisted Living" And cell.Value <> "Restaurants" And cell.Value <> "Retail Stores / Strip Shopping Centers" And cell.Value <> "Sawmills /Planing/ Chipping Mills" And cell.Value <> "Shopping Malls" And cell.Value <> "Stadiums/Arenas" And cell.Value <> "Theatres - Play/Dance Halls" And cell.Value <> "Universities/Colleges" And cell.Value <> "Vacant Buildings" And cell.Value <> "Wholesale - Cold Storage Facilities" And cell.Value <> "Wholesale - High/Severe Hzd. Goods" And cell.Value <> "Wholesale - Light Hzd. Goods" Then cell.Interior.Color = RGB(3, 3, 3)
        If cell.Interior.Color = RGB(3, 3, 3) And cell.Value <> "Wholesale - Medium Hzd. Goods" And cell.Value <> "Wineries" Then cell.Interior.Color = RGB(200, 200, 200)
        If cell.Interior.Color <> RGB(200, 200, 200) Then cell.Interior.Color = RGB(255, 255, 255)
        If cell.Value = "" Then cell.Interior.Color = RGB(196, 228, 242)
    Next cell
    
''''  Check Construction Validation  ''''
    
    Dim consRng As Range
    Set consRng = Range("J2:J" & numberOfLocations)
    
    For Each cell In consRng
        If cell.Value <> "" Then cell.Interior.Color = RGB(255, 255, 255)
        If cell.Value = "Non-Combustible" Or cell.Value = "Frame" Or cell.Value = "Joisted Masonry" Or cell.Value = "Masonry Non-Combustible" Or cell.Value = "Mod. Fire Res." Or cell.Value = "Fire Resistive" Then cell.Interior.Color = RGB(255, 255, 255)
        If cell.Value <> "Non-Combustible" And cell.Value <> "Frame" And cell.Value <> "Joisted Masonry" And cell.Value <> "Masonry Non-Combustible" And cell.Value <> "Mod. Fire Res." And cell.Value <> "Fire Resistive" Then cell.Interior.Color = RGB(200, 200, 200)
        If cell.Value = "" Then cell.Interior.Color = RGB(196, 228, 242)
    Next cell
    

''''  Check Valuation Validation  ''''
    
    Dim valRng As Range
    Set valRng = Range("K2:K" & numberOfLocations)
    
    For Each cell In valRng
        If cell.Value <> "" Then cell.Interior.Color = RGB(255, 255, 255)
        If cell.Value <> "Replacement Cost Value" And cell.Value <> "Actual Cash Value" And cell.Value <> "Functional Replacement Cost" And cell.Value <> "Indicated on Equipment Schedule" And cell.Value <> "Market Value" And cell.Value <> "Stated Value" Then cell.Interior.Color = RGB(200, 200, 200)
        If cell.Value = "Replacement Cost Value" Or cell.Value = "Actual Cash Value" Or cell.Value = "Functional Replacement Cost" Or cell.Value = "Indicated on Equipment Schedule" Or cell.Value = "Market Value" Or cell.Value = "Stated Value" Then cell.Interior.Color = RGB(255, 255, 255)
        If cell.Value = "" Then cell.Interior.Color = RGB(196, 228, 242)
    Next cell
   
 ''''  Check If Stories Area is blank  ''''

    Dim storiesRng As Range
    Set storiesRng = Range("S2:S" & numberOfLocations)
    
    For Each cell In storiesRng
        If cell.Value <> "" Then cell.Interior.Color = RGB(255, 255, 255)
        If cell.Value = "" Then cell.Interior.Color = RGB(196, 228, 242)
        If IsNumeric(cell.Value) = True And cell.Value <> "" Then cell.Interior.Color = RGB(255, 255, 255)
        If IsNumeric(cell.Value) = False And cell.Value <> "" Then cell.Interior.Color = RGB(200, 200, 200)
    Next cell
    
''''  Check If Floor Area is blank  ''''

    Dim floorRng As Range
    Set floorRng = Range("T2:T" & numberOfLocations)
    
    For Each cell In floorRng
        If cell.Value <> "" Then cell.Interior.Color = RGB(255, 255, 255)
        If cell.Value = "" Then cell.Interior.Color = RGB(196, 228, 242)
        If IsNumeric(cell.Value) = True And cell.Value <> "" Then cell.Interior.Color = RGB(255, 255, 255)
        If IsNumeric(cell.Value) = False And cell.Value <> "" Then cell.Interior.Color = RGB(200, 200, 200)
    Next cell
    
''''  Check Valuation Validation  ''''
    
    Dim roofRng As Range
    Set roofRng = Range("X2:X" & numberOfLocations)
    
    For Each cell In roofRng
        If cell.Value <> "" Then cell.Interior.Color = RGB(255, 255, 255)
        If cell.Value = "" Then cell.Value = "Unknown"
        If cell.Value <> "0-5 years" And cell.Value <> "6-10 years" And cell.Value <> "11+ years" And cell.Value <> "Unknown" Then cell.Interior.Color = RGB(200, 200, 200)
        If cell.Value = "0-5 years" And cell.Value = "6-10 years" And cell.Value = "11+ years" And cell.Value = "Unknown" Then cell.Interior.Color = RGB(255, 255, 255)
    Next cell
    
    
''''  Check for Bad Values  ''''

    'Define Range
    Dim valuationRng As Range
    'Set range to selection
    Set valuationRng = Range("L2:Q" & numberOfLocations)
    
    For Each cell In valuationRng
        If cell.Value <> "" Then cell.Interior.Color = RGB(255, 255, 255)
        'If there is no value in the cell, leave it blank
        If IsNumeric(cell) = False Then cell.Interior.Color = RGB(200, 200, 200)
        
    'Continue to traverse selection
    Next cell
    
    
'''' Trim white characters ''''
    Range("A2:X" & numberOfLocations).NumberFormat = "General"

'''' Set to general format ''''
    Dim trimRng As Range
    Set trimRng = Range("A2:X" & numberOfLocations)
    For Each cell In trimRng
    cell.Value = Trim(cell)
    Next cell
    
'''' Resize ''''

    Cells.Select
    Range("J13").Activate
    Selection.Columns.AutoFit


'''' Delete Conditional Formatting ''''

    Cells.FormatConditions.Delete
    
        ''''
    '''' Add Validation to Roof Year ''''
        Range("X2:X" & numberOfLocations).Select
        
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="0-5 years,6-10 years,11+ years,Unknown"
    
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = False
        End With
    
    ''''
    '''' Add Validation to Valuation Type ''''

        Range("K2:K" & numberOfLocations).Select
        
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="Replacement Cost Value,Actual Cash Value,Functional Replacement Cost,Indicated on Equipment Schedule,Market Value,Stated Value"
    
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = False
        End With
    
    
    ''''
    '''' Add Validation to Construction Type ''''
        Range("J2:j" & numberOfLocations).Select
        
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="Non-Combustible,Frame,Joisted Masonry,Masonry Non-Combustible,Mod. Fire Res.,Fire Resistive"
    
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = False
        End With
    
    ''''
    '''' Add Validation to Occupancy Type ''''
    
        ActiveWorkbook.Names.Add _
        Name:="occRangeWS3", _
        RefersTo:="=MacrosStorage.xlsm!listing"
        
        Range("i2:i" & numberOfLocations).Select
    
        With Selection.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=occRangeWS3"
    
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = False
        End With

''''  Check for Bad Values  ''''

    'Define Range
    Dim charRng As Range
    'Set range to selection
    Set charRng = Range("A2:X" & numberOfLocations)
    
    For Each cell In charRng
        'If there is no value in the cell, leave it blank
        If AlphaNumeric(cell) = False Then cell.Interior.Color = RGB(255, 0, 0)
        
    'Continue to traverse selection
    Next cell

' Make sure states are in proper format
    'Define Range
    Dim stateRng As Range
    'Set range to selection
    Set stateRng = Range("F2:F" & numberOfLocations)
        
    For Each cell In stateRng
        'If there is no value in the cell, leave it blank
        cell.Value = UCase(cell.Value)
        If cell.Value = "ALABAMA" Then cell.Value = "AL"
        If cell.Value = "ALASKA" Then cell.Value = "AK"
        If cell.Value = "ARIZONA" Then cell.Value = "AZ"
        If cell.Value = "ARKANSAS" Then cell.Value = "AR"
        If cell.Value = "CALIFORNIA" Then cell.Value = "CA"
        If cell.Value = "COLORADO" Then cell.Value = "CO"
        If cell.Value = "CONNECTICUT" Then cell.Value = "CT"
        If cell.Value = "DELAWARE" Then cell.Value = "DE"
        If cell.Value = "FLORIDA" Then cell.Value = "FL"
        If cell.Value = "GEORGIA" Then cell.Value = "GA"
        If cell.Value = "HAWAII" Then cell.Value = "HI"
        If cell.Value = "IDAHO" Then cell.Value = "ID"
        If cell.Value = "ILLINOIS" Then cell.Value = "IL"
        If cell.Value = "INDIANA" Then cell.Value = "IN"
        If cell.Value = "IOWA" Then cell.Value = "IA"
        If cell.Value = "KANSAS" Then cell.Value = "KS"
        If cell.Value = "KENTUCKY" Then cell.Value = "KY"
        If cell.Value = "LOUISIANA" Then cell.Value = "LA"
        If cell.Value = "MAINE" Then cell.Value = "ME"
        If cell.Value = "MARYLAND" Then cell.Value = "MD"
        If cell.Value = "MASSACHUSETTS" Then cell.Value = "MA"
        If cell.Value = "MICHIGAN" Then cell.Value = "MI"
        If cell.Value = "MINNESOTA" Then cell.Value = "MN"
        If cell.Value = "MISSISSIPPI" Then cell.Value = "MS"
        If cell.Value = "MISSOURI" Then cell.Value = "MO"
        If cell.Value = "MONTANA" Then cell.Value = "MT"
        If cell.Value = "NEBRASKA" Then cell.Value = "NE"
        If cell.Value = "NEVADA" Then cell.Value = "NV"
        If cell.Value = "NEW HAMPSHIRE" Then cell.Value = "NH"
        If cell.Value = "NEW JERSEY" Then cell.Value = "NJ"
        If cell.Value = "NEW MEXICO" Then cell.Value = "NM"
        If cell.Value = "NEW YORK" Then cell.Value = "NY"
        If cell.Value = "NORTH CAROLINA" Then cell.Value = "NC"
        If cell.Value = "NORTH DAKOTA" Then cell.Value = "ND"
        If cell.Value = "OHIO" Then cell.Value = "OH"
        If cell.Value = "OKLAHOMA" Then cell.Value = "OK"
        If cell.Value = "OREGON" Then cell.Value = "OR"
        If cell.Value = "PENNSYLVANIA" Then cell.Value = "PA"
        If cell.Value = "RHODE ISLAND" Then cell.Value = "RI"
        If cell.Value = "SOUTH CAROLINA" Then cell.Value = "SC"
        If cell.Value = "SOUTH DAKOTA" Then cell.Value = "SD"
        If cell.Value = "TENNESSEE" Then cell.Value = "TN"
        If cell.Value = "TEXAS" Then cell.Value = "TX"
        If cell.Value = "UTAH" Then cell.Value = "UT"
        If cell.Value = "VERMONT" Then cell.Value = "VT"
        If cell.Value = "VIRGINIA" Then cell.Value = "VA"
        If cell.Value = "WASHINGTON" Then cell.Value = "WA"
        If cell.Value = "WEST VIRGINIA" Then cell.Value = "WV"
        If cell.Value = "WISCONSIN" Then cell.Value = "WI"
        If cell.Value = "WYOMING" Then cell.Value = "WY"
        'Continue to traverse selection
    Next cell

End Sub
