Private Sub addressBox_Change()

End Sub

Private Sub cityBox_Change()

End Sub

Private Sub CommandButton1_Click()
    
    ''''
    '''' Transfer Loction Numbers ''''
    If locationBox = "" Then MsgBox "You need to input the number of locations", vbOK + vbExclamation, "Error"
    

    
    ''''
    '''' Rename Existing Sheet ''''
    Dim exists2 As Boolean
    
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = "BaseRefSheet" Then
            exists2 = True
        End If
    Next i
    
    If Not exists2 Then
        ActiveSheet.Name = "BaseRefSheet"
    End If
    
    
    ''''
    '''' Look for Sheet Name and Create New Sheet if Needed ''''
    Dim exists As Boolean
    
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = "Sheet1" Then
            exists = True
        End If
    Next i
    
    If Not exists Then
        Sheets.Add.Name = "Sheet1"
    End If
    
    
    
    ''''
    '''' Look for Sheet Name and Create New Sheet if Needed ''''
    Dim exists3 As Boolean
    
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = "WS3Upload" Then
            exists = True
        End If
    Next i
    
    If Not exists Then
        Sheets.Add.Name = "WS3Upload"
    End If
    
    
    ''''
    '''' Create Header Format ''''
    ' Put headers in the proper location in the upload form
    Sheets("WS3Upload").Select
        Range("a1").Value = "Contract ID"
        Range("b1").Value = "LocationNumber"
        Range("c1").Value = "BuildingNumber"
        Range("d1").Value = "NumberOfBuildings"
        Range("e1").Value = "Address"
        Range("f1").Value = "City"
        Range("g1").Value = "State"
        Range("h1").Value = "ZipCode"
        Range("i1").Value = "BuildingValue"
        Range("j1").Value = "ContentsValue"
        Range("k1").Value = "OtherValue"
        Range("l1").Value = "BiValue"
        Range("m1").Value = "Construction"
        Range("n1").Value = "Occupancy"
        Range("o1").Value = "Year Built"
        Range("p1").Value = "NumberOfStories"
        Range("q1").Value = "FloorArea"
        Range("r1").Value = "LocPerils"
        Range("s1").Value = "DeductType"
        Range("t1").Value = "DeductBldg"
        Range("u1").Value = "DeductOther"
        Range("v1").Value = "DeductContent"
        Range("w1").Value = "DeductTime"
        Range("x1").Value = "LocPerils2"
        Range("y1").Value = "DeductType2"
        Range("z1").Value = "DeductBldg2"
        Range("aa1").Value = "DeductOther2"
        Range("ab1").Value = "DeductContent2"
        Range("ac1").Value = "DeductTime2"
        Range("ad1").Value = "LocPerils3"
        Range("ae1").Value = "DeductType3"
        Range("af1").Value = "DeductBldg3"
        Range("ag1").Value = "DeductOther3"
        Range("ah1").Value = "DeductContent3"
        Range("ai1").Value = "DeductTime3"
        Range("aj1").Value = "LocPerils4"
        Range("ak1").Value = "DeductType4"
        Range("al1").Value = "DeductBldg4"
        Range("am1").Value = "DeductOther4"
        Range("an1").Value = "DeductContent4"
        Range("ao1").Value = "DeductTime4"
        Range("ap1").Value = "LocPerils5"
        Range("aq1").Value = "DeductType5"
        Range("ar1").Value = "DeductBldg5"
        Range("as1").Value = "DeductOther5"
        Range("at1").Value = "DeductContent5"
        Range("au1").Value = "DeductTime5"
        Range("av1").Value = "LocPerils6"
        Range("aw1").Value = "DeductType6"
        Range("ax1").Value = "DeductBldg6"
        Range("ay1").Value = "DeductOther6"
        Range("az1").Value = "DeductContent6"
        Range("ba1").Value = "DeductTime6"
    
    If locationBox <> "" Then
        Sheets("BaseRefSheet").Select
        Range("a1:x" & locationBox.Text + 1).Copy
        Sheets("Sheet1").Select
        Range("a1:a" & locationBox.Text + 1).Select
        Selection.PasteSpecial xlPasteValues
    End If
    
    ''''
    '''' Transfer Location Number ''''
    '
    If locationBox <> "" Then
        Sheets("Sheet1").Select
        Range("a2:a" & locationBox.Text + 1).Copy
        Sheets("WS3Upload").Select
        Range("b2:b" & locationBox.Text + 1).Select
        ActiveSheet.Paste
    End If
    
    
    
    ''''
    '''' Transfer Building Number ''''
    '
    If locationBox <> "" And buildingNumBox <> "" Then
        Sheets("Sheet1").Select
        Range(buildingNumBox.Text & "2:" & buildingNumBox.Text & locationBox.Text + 1).Copy
        Sheets("WS3Upload").Select
        Range("c2:c" & locationBox.Text + 1).Select
        ActiveSheet.Paste
    End If
    
    
    ''''
    '''' Transfer Number of Buildings ''''
    '
    If locationBox <> "" And numBuildingsBox <> "" Then
        Sheets("Sheet1").Select
        Range(numBuildingsBox.Text & "2:" & numBuildingsBox.Text & locationBox.Text + 1).Copy
        Sheets("WS3Upload").Select
        Range("d2:d" & locationBox.Text + 1).Select
        ActiveSheet.Paste
    End If
    
    
    ''''
    '''' Transfer Address ''''
    '
    If locationBox <> "" And addressBox <> "" Then
        Sheets("Sheet1").Select
        Range(addressBox.Text & "2:" & addressBox.Text & locationBox.Text + 1).Copy
        Sheets("WS3Upload").Select
        Range("e2:e" & locationBox.Text + 1).Select
        ActiveSheet.Paste
    End If
    
    
    ''''
    '''' Transfer City ''''
    '
    If locationBox <> "" And cityBox <> "" Then
        Sheets("Sheet1").Select
        Range(cityBox.Text & "2:" & cityBox.Text & locationBox.Text + 1).Copy
        Sheets("WS3Upload").Select
        Range("f2:f" & locationBox.Text + 1).Select
        ActiveSheet.Paste
    End If
    
    
    ''''
    '''' Transfer State ''''
    '
    If locationBox <> "" And stateBox <> "" Then
        Sheets("Sheet1").Select
        Range(stateBox.Text & "2:" & stateBox.Text & locationBox.Text + 1).Copy
        Sheets("WS3Upload").Select
        Range("g2:g" & locationBox.Text + 1).Select
        ActiveSheet.Paste
    End If
    
    
    ''''
    '''' Transfer ZIP Code ''''
    '
    If locationBox <> "" And zipBox <> "" Then
        Sheets("Sheet1").Select
        Range(zipBox.Text & "2:" & zipBox.Text & locationBox.Text + 1).Copy
        Sheets("WS3Upload").Select
        Range("h2:h" & locationBox.Text + 1).Select
        ActiveSheet.Paste
    End If
    
    
    ''''
    '''' Transfer Dedutible Type ''''
    '
    If locationBox <> "" And ratingBox <> "" Then
        Sheets("Sheet1").Select
        Range(ratingBox.Text & "2:" & ratingBox.Text & locationBox.Text + 1).Copy
        Sheets("WS3Upload").Select
        Range("i2:i" & locationBox.Text + 1).Select
        ActiveSheet.Paste
    End If
    
    
    ''''
    '''' Transfer Occupancy ''''
    '
    If locationBox <> "" And occupancyBox <> "" Then
        Sheets("Sheet1").Select
        Range(occupancyBox.Text & "2:" & occupancyBox.Text & locationBox.Text + 1).Copy
        Sheets("WS3Upload").Select
        Range("n2:n" & locationBox.Text + 1).Select
        ActiveSheet.Paste
    End If
    
    
    ''''
    '''' Transfer constructionBox ''''
    '
    If locationBox <> "" And constructionBox <> "" Then
        Sheets("Sheet1").Select
        Range(constructionBox.Text & "2:" & constructionBox.Text & locationBox.Text + 1).Copy
        Sheets("WS3Upload").Select
        Range("m2:m" & locationBox.Text + 1).Select
        ActiveSheet.Paste
    End If
    
    

    
    
    ''''
    '''' Transfer buildingBox ''''
    '
    If locationBox <> "" And buildingBox <> "" Then
        Sheets("Sheet1").Select
        Range(buildingBox.Text & "2:" & buildingBox.Text & locationBox.Text + 1).Copy
        Sheets("WS3Upload").Select
        Range("i2:i" & locationBox.Text + 1).Select
        ActiveSheet.Paste
    End If
    
    
    ''''
    '''' Transfer contentsBox ''''
    '
    If locationBox <> "" And contentsBox <> "" Then
        Sheets("Sheet1").Select
        Range(contentsBox.Text & "2:" & contentsBox.Text & locationBox.Text + 1).Copy
        Sheets("WS3Upload").Select
        Range("j2:j" & locationBox.Text + 1).Select
        ActiveSheet.Paste
    End If
    
    
    ''''
    '''' Transfer ppoBox ''''
    '
    If locationBox <> "" And ppoBox <> "" Then
        Sheets("Sheet1").Select
        Range(ppoBox.Text & "2:" & ppoBox.Text & locationBox.Text + 1).Copy
        Sheets("WS3Upload").Select
        Range("k2:k" & locationBox.Text + 1).Select
        ActiveSheet.Paste
    End If
    
    
    ''''
    '''' Transfer stockBox ''''
    '
    If locationBox <> "" And stockBox <> "" Then
        Sheets("Sheet1").Select
        Range(stockBox.Text & "2:" & stockBox.Text & locationBox.Text + 1).Copy
        Sheets("WS3Upload").Select
        Range("m2:m" & locationBox.Text + 1).Select
        ActiveSheet.Paste
    End If
    
    
    ''''
    '''' Transfer biBox ''''
    '
    If locationBox <> "" And biBox <> "" Then
        Sheets("Sheet1").Select
        Range(biBox.Text & "2:" & biBox.Text & locationBox.Text + 1).Copy
        Sheets("WS3Upload").Select
        Range("l2:l" & locationBox.Text + 1).Select
        ActiveSheet.Paste
    End If
    
    
    ''''
    '''' Transfer yearBox ''''
    '
    If locationBox <> "" And yearBox <> "" Then
        Sheets("Sheet1").Select
        Range(yearBox.Text & "2:" & yearBox.Text & locationBox.Text + 1).Copy
        Sheets("WS3Upload").Select
        Range("o2:o" & locationBox.Text + 1).Select
        ActiveSheet.Paste
    End If
    
    
  
    
    
    ''''
    '''' Transfer storiesBox ''''
    '
    If locationBox <> "" And storiesBox <> "" Then
        Sheets("Sheet1").Select
        Range(storiesBox.Text & "2:" & storiesBox.Text & locationBox.Text + 1).Copy
        Sheets("WS3Upload").Select
        Range("p2:p" & locationBox.Text + 1).Select
        ActiveSheet.Paste
    End If
    
    
    ''''
    '''' Transfer floorBox ''''
    '
    If locationBox <> "" And floorBox <> "" Then
        Sheets("Sheet1").Select
        Range(floorBox.Text & "2:" & floorBox.Text & locationBox.Text + 1).Copy
        Sheets("WS3Upload").Select
        Range("q2:q" & locationBox.Text + 1).Select
        ActiveSheet.Paste
    End If
    
    
    ''''
    '''' Transfer DeductOther ''''
    '
    If locationBox <> "" And sprinklerBox <> "" Then
        Sheets("Sheet1").Select
        Range(sprinklerBox.Text & "2:" & sprinklerBox.Text & locationBox.Text + 1).Copy
        Sheets("WS3Upload").Select
        Range("j2:j" & locationBox.Text + 1).Select
        ActiveSheet.Paste
    End If
    
    
    ''''
    '''' Transfer DeductContent ''''
    '
    If locationBox <> "" And alarmBox <> "" Then
        Sheets("Sheet1").Select
        Range(alarmBox.Text & "2:" & alarmBox.Text & locationBox.Text + 1).Copy
        Sheets("WS3Upload").Select
        Range("l1:l" & locationBox.Text + 1).Select
        ActiveSheet.Paste
    End If
    
    
    ''''
    '''' Transfer DeductTime''''
    '
    If locationBox <> "" And securityBox <> "" Then
        Sheets("Sheet1").Select
        Range(securityBox.Text & "2:" & securityBox.Text & locationBox.Text + 1).Copy
        Sheets("WS3Upload").Select
        Range("m2:m" & locationBox.Text + 1).Select
        ActiveSheet.Paste
    End If
    
    
    ''''
    '''' Transfer roofBox ''''
    '
    If locationBox <> "" And roofBox <> "" Then
        Sheets("Sheet1").Select
        Range(roofBox.Text & "2:" & roofBox.Text & locationBox.Text + 1).Copy
        Sheets("WS3Upload").Select
        Range("i2:i" & locationBox.Text + 1).Select
        ActiveSheet.Paste
    End If

    ''''Transfer DeductBldg'''
     ''''
    '
    If locationBox <> "" And protectionBox <> "" Then
        Sheets("Sheet1").Select
        Range(protectionBox.Text & "2:" & protectionBox.Text & locationBox.Text + 1).Copy
        Sheets("WS3Upload").Select
        Range("k2:k" & locationBox.Text + 1).Select
        ActiveSheet.Paste
    End If

    
    'Call Validation
    
End Sub


Private Sub Frame2_Click()

End Sub

Private Sub Label18_Click()

End Sub

Private Sub Label28_Click()

End Sub

Private Sub Label29_Click()

End Sub

Private Sub Label31_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub UserForm_Click()

End Sub
