Attribute VB_Name = "AddressSplit"
' Address Split v0.1 (Excel macro) - split street addresses into individual
' components using fancy string manipulation and magic numbers.
' Does not check for street types in street names; i.e. Holston Court Ln will
' not split correctly: Holston Ct will be street, Ln will go to the city column.
' Algorithm needs more magic or fancier string manipulation.
' Copyright (C) 2014 David Baker, Milligan College
' Licensed under GPL v2.0 - see license.txt
' David Baker https://github.com/dbaker3

Sub AddressSplit()
Attribute AddressSplit.VB_ProcData.VB_Invoke_Func = " \n14"

    'Column letter assignments in spreadsheet
    Dim fullAddressColumn: fullAddressColumn = "A"
    Dim streetColumn: streetColumn = "B"
    Dim cityColumn: cityColumn = "C"
    Dim stateColumn: stateColumn = "D"
    Dim zipColumn: zipColumn = "E"
    
    Dim curRow
    Dim totalRows: totalRows = InputBox("Number of Rows:")
    Dim addressRange: addressRange = fullAddressColumn + "1:" + fullAddressColumn + totalRows
    
    Range(zipColumn + "1:" + zipColumn + totalRows).NumberFormat = "@" 'format ZIP as text instead of number
    
    'Abbreviate street-types and remove extra punctuation
    For Each cell In Range(addressRange)
        cell.Value = Replace(cell.Value, " drive ", " Dr ", 1, -1, vbTextCompare)
        cell.Value = Replace(cell.Value, " road ", " Rd ", 1, -1, vbTextCompare)
        cell.Value = Replace(cell.Value, " place ", " Pl ", 1, -1, vbTextCompare)
        cell.Value = Replace(cell.Value, " street ", " St ", 1, -1, vbTextCompare)
        cell.Value = Replace(cell.Value, " parkway ", " Pkwy ", 1, -1, vbTextCompare)
        cell.Value = Replace(cell.Value, " lane ", " Ln ", 1, -1, vbTextCompare)
        cell.Value = Replace(cell.Value, " boulevard ", " Blvd ", 1, -1, vbTextCompare)
        cell.Value = Replace(cell.Value, " circle ", " Cir ", 1, -1, vbTextCompare)
        cell.Value = Replace(cell.Value, " avenue ", " Ave ", 1, -1, vbTextCompare)
        cell.Value = Replace(cell.Value, " terrace ", " Ter ", 1, -1, vbTextCompare)
        cell.Value = Replace(cell.Value, " court ", " Ct ", 1, -1, vbTextCompare)
        cell.Value = Replace(cell.Value, ",", "", 1, -1, vbTextCompare)
        cell.Value = Replace(cell.Value, ".", "", 1, -1, vbTextCompare)
    Next
    
    'Street types to detect. Splits based on location of these
    Dim streetTypes As Variant
    streetTypes = Array( _
                    " Dr ", _
                    " Rd ", _
                    " Pl ", _
                    " St ", _
                    " Pkwy ", _
                    " Ln ", _
                    " Blvd ", _
                    " Cir ", _
                    " Ave ", _
                    " Ct ", _
                    " Way ", _
                    " Ter ")
    
    'ZIP Code
    curRow = 0
    Dim zipCell
    For Each cell In Range(addressRange)
        curRow = curRow + 1
        zipCell = CStr(curRow)
        zipCell = zipColumn + zipCell

        Dim zip
        zip = Right(cell, 10)
        
        If IsNumeric(Left(zip, 1)) = True Then  '9 digit zip
            Range(zipCell).Value = zip
        Else                                    '5 digit zip
            Range(zipCell).Value = Right(zip, 5)
        End If
    Next
    
    'State
    curRow = 0
    Dim stateCell
    For Each cell In Range(addressRange)
        curRow = curRow + 1
        stateCell = CStr(curRow)
        stateCell = stateColumn + stateCell

        Dim stateSub
        stateSub = Right(cell, 13)
        
        If IsNumeric(Right(Left(stateSub, 4), 1)) = True Then   '9 digit zip
            stateSub = Left(stateSub, 2)
            Range(stateCell).Value = stateSub
        Else                                                    '5 digit zip
            Range(stateCell).Value = Right(Left(stateSub, 7), 2)
        End If
    Next
    
    curRow = 0

    'Street Address & City
    curRow = 0
    Dim streetCell
    Dim cityCell
    Dim streetString
    Dim cityString
    
    For Each cell In Range(addressRange)
        curRow = curRow + 1
        streetCell = CStr(curRow)
        cityCell = CStr(curRow)
        streetCell = streetColumn + streetCell
        cityCell = cityColumn + cityCell
        
        For Each streetType In streetTypes
            If (InStr(1, cell, streetType, vbTextCompare)) Then
                streetString = Left(cell, InStr(1, cell, streetType, vbTextCompare)) + Trim(streetType)
                Range(streetCell).Value = streetString
                
                'get state and zip so we can remove it from city string
                stateSub = Range(stateColumn + CStr(curRow)).Value
                zip = Range(zipColumn + CStr(curRow)).Value
                
                cityString = Replace(cell, streetString, "", 1, -1, vbTextCompare) 'remove street
                cityString = Replace(cityString, stateSub, "", 1, -1, vbTextCompare) 'remove state
                cityString = Replace(cityString, zip, "", 1, -1, vbTextCompare) 'remove zip
                Range(cityCell).Value = Trim(cityString)
            End If
        Next
        
    Next
    
End Sub
