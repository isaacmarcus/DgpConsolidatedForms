Attribute VB_Name = "DataConsolidationSub"
Option Explicit
Const compSignalList As String = "LLC,LTD,PVT,INC,C/O,M/V,M.V.,C.O.,LIMITED,INCORPORATED"
Const countryList As String = "AUSTRALIA,INDIA,CHINA,INDONESIA,JAPAN,KOREA,UNITED KINGDOM,HONG KONG,PHILIPPINES,THAILAND,GERMANY,VIETNAM,TAIWAN"
Const monthList As String = "JANUARY,FEBRUARY,MARCH,APRIL,MAY,JUNE,JULY,AUGUST,SEPTEMBER,OCTOBER,NOVEMBER,DECEMBER,JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC"

Sub ConsolidateData()
    'Optimize Macro Speed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    'Prompt User to choose folder to get files from
    Dim path As String
    MsgBox ("Please choose folder to get files from")
    path = GetFolder() & "\"
    
    If path = "\" Then
        MsgBox "No folder selected!"
        GoTo cancel_folder_selection
    End If
    
    'Check total number of files in directory
    Dim numFiles As Long
    Dim curFileIndex As Long
    Dim pctDone As Single
    curFileIndex = 1
    numFiles = CountFilesInFolder(path, "*.xls*")
    pctDone = curFileIndex / numFiles
    'Set progress bar width to 0 and show progress bar
    ufProgress.LabelProgress.Width = 0
    ufProgress.Show

    'Create file name variables
    Dim FileName As String
    Dim fileNameElementIndex As Integer
    Dim fileNameArraySize As Integer
    Dim FileNameArray As Variant
    FileName = Dir(path & "*.xls*") 'Call the first excel file in folder
    
    'mWS will be the reference for the ws where we store the data
    Dim mWS As Worksheet
    Set mWS = ThisWorkbook.Sheets(1)
    Dim mWSCountry As Worksheet
    Dim mWSCity As Worksheet
    Set mWSCountry = ThisWorkbook.Sheets("List of Countries") 'sheet in ThisWorkBook with list of countries
    Set mWSCity = ThisWorkbook.Sheets("List of Cities") 'sheet in ThisWorkBook with list of cities
    Dim curWorkBook As Workbook 'object to handle current workbook to be parsed
    Dim curWorkSheet As Worksheet 'object to handle current worksheet to be parsed
    Dim workBookElement As Worksheet ' variable for going through worksheets
     
    'Create variables for data we want to store
    Dim curDate As String
    Dim curCountry As String
    Dim curUNList As Collection
    Dim curUNNumber As String
    Dim curShipper As String
    Dim curCustomer As String
    Dim curState As String
    
    'Create variables for doing calculations etc
    Dim UNListIndex As Variant
    Dim UNTableIndex As Integer
    Dim UNFileIndex As Integer
    Dim curRow As Integer
    curRow = mWS.Range("A1048576").End(xlUp).Row + 1
    
    'Creating variables for COUNTRIES CHECK
    Dim curRowinColumn As Variant
    Dim lastrowCountries As Integer
    Dim lastRowFinderCountry As Integer
    Dim arrayElement As Variant
    Dim tempString As String
    Dim countryFound As Boolean
    Dim shipperArray() As String
    Dim saIndex As Integer
    Dim curStringShipper As String
    Dim arrElem As Variant
    Dim forBool As Boolean
    Dim i As Integer 'i to be used as counter for loops
    Dim j As Integer 'j to be used as counter for inner loops
    lastrowCountries = mWSCountry.Range("A30000").End(xlUp).Row
    
    'Create error variables
    Dim errorNote As String
    Dim errorIndex As Integer
    errorIndex = mWS.Range("N1048576").End(xlUp).Row + 1
    Dim notEmptySheet As Boolean
    
    'Start Loop to go through workbooks in folder
    Do While FileName <> ""
        
        notEmptySheet = False
    
        pctDone = curFileIndex / numFiles
        'Update progress bar for every new file we start processing
        With ufProgress
        .LabelCaption.Caption = "Processing File " & curFileIndex & " of " & numFiles
        .LabelProgress.Width = pctDone * (.FrameProgress.Width)
        End With
        ufProgress.Repaint
    
        On Error GoTo country_error:
        
        'set variable for current workbook
        Set curWorkBook = Workbooks.Open(FileName:=path & FileName, ReadOnly:=True, UpdateLinks:=False, CorruptLoad:=xlRepairFile)
        
        '~~~~~~~~~~~~~~~~~~~~~
        '~check if draft file~
        '~~~~~~~~~~~~~~~~~~~~~
        If InStr(UCase(FileName), "DRAFT") Then
            errorNote = "Draft file"
            errorIndex = errorFileParse(errorNote, FileName, mWS, errorIndex) 'Call the function to input error file, then increment index
            GoTo end_of_while_loop
        End If
        
        '~~~~~~~~~~~~~~~~~~~~~~~
        '~check if revised file~
        '~~~~~~~~~~~~~~~~~~~~~~~
        If InStr(UCase(FileName), "REVISE") Then
            errorNote = "Revised file"
            errorIndex = errorFileParse(errorNote, FileName, mWS, errorIndex) 'Call the function to input error file, then increment index
            GoTo end_of_while_loop
        End If
        
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        '~check if less than 1 sheet~
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        If curWorkBook.Sheets.Count < 2 Then
            errorNote = "File only has 1 sheet"
            errorIndex = errorFileParse(errorNote, FileName, mWS, errorIndex) 'Call the function to input error file, then increment index
            GoTo end_of_while_loop
        End If
        
        'set variable for current sheet, will be getting info fomr first sheet
        Set curWorkSheet = curWorkBook.Sheets(1)
        
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        '~check if file has correct sheet~
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        If curWorkSheet.Range("D1") <> "SHIPPER'S DECLARATION FOR DANGEROUS GOODS" Then
            errorNote = "first sheet has different label"
            errorIndex = errorFileParse(errorNote, FileName, mWS, errorIndex) 'Call the function to input error file, then increment index
            GoTo end_of_while_loop
        End If
        
        'Reset variables that may overlap to next document
        curDate = ""
        curCustomer = ""
        UNFileIndex = 77
        
        '-------------------
        '*FIND CURRENT DATE*
        '-------------------
        FileNameArray = Split(Trim(TrimString(FileName))) 'Split the file name into an array to be parsed
        fileNameArraySize = UBound(FileNameArray) - LBound(FileNameArray) 'get the size of said array
        Dim UNFileIndexFound As Boolean
        UNFileIndexFound = False
        
        'Loop through array of strings and look for month name
        For fileNameElementIndex = 1 To fileNameArraySize
            
            'Find Index of UN Number in file name
            If UCase(FileNameArray(fileNameElementIndex)) = "UN" And UNFileIndexFound = False Then
                UNFileIndex = fileNameElementIndex
                UNFileIndexFound = True
            End If
            
            'if monthname is found, set curDate to date found in file name
            If InStr(monthList, UCase(FileNameArray(fileNameElementIndex))) <> 0 Then
                If Len(FileNameArray(fileNameElementIndex)) > 2 Then
                    curDate = FileNameArray(fileNameElementIndex - 1) & FileNameArray(fileNameElementIndex) & FileNameArray(fileNameElementIndex + 1)
                End If
            End If
        Next fileNameElementIndex
        
        ' Check if there was a date that was found in file name
        If curDate = "" Then
            errorNote = "No proper date found"
            errorIndex = errorFileParse(errorNote, FileName, mWS, errorIndex)
            GoTo end_of_while_loop
        End If
        
        '---------------------------------
        '*FIND CUSTOMER NAME IN FILE NAME*
        '---------------------------------
        'Check if UN file index was found if not, parse error
        If UNFileIndex = 77 Then
            errorNote = "No UN Signifier in filename for Customer"
            errorIndex = errorFileParse(errorNote, FileName, mWS, errorIndex)
            GoTo end_of_while_loop
        End If
        
        'Loop through array and append curCustomer String accordingly
        For fileNameElementIndex = 0 To UNFileIndex - 1
            curCustomer = curCustomer & " " & FileNameArray(fileNameElementIndex)
        Next fileNameElementIndex
        
        
        '----------------------
        '*FIND CURRENT COUNTRY*
        '----------------------
        curShipper = ""
        countryFound = False
        forBool = False
        Dim consigneeIndex As Integer
    
        On Error GoTo country_error:
        
            'Find consignee table and index
            For i = 1 To curWorkSheet.Range("C30000").End(xlUp).Row
            
                '----------------------
                '*FIND CURRENT SHIPPER*
                '----------------------
                If UCase(Trim(curWorkSheet.Range("C" & i).Value)) = "SHIPPER" Then
                    lastRowFinderCountry = curWorkSheet.Range("D" & i).End(xlDown).Row
                    If IsEmpty(curWorkSheet.Range("D" & lastRowFinderCountry + 1)) = False Then
                        lastRowFinderCountry = curWorkSheet.Range("D" & lastRowFinderCountry).End(xlDown).Row
                    End If
                    
                    'loop through shipper details
                    For saIndex = i To lastRowFinderCountry
                        curStringShipper = curWorkSheet.Range("D" & saIndex).Value
                        curStringShipper = Replace(curStringShipper, ":", "")
                        'Check for ON BEHALF OF signifier
                        If InStr(UCase(curStringShipper), "ON BEHALF OF") <> 0 Then
                            forBool = True
                            curShipper = Trim(TrimString(curWorkSheet.Range("D" & saIndex + 1).Value))
                            If curShipper = "" Then
                                curShipper = Trim(TrimString(curWorkSheet.Range("D" & saIndex).Value))
                                curShipper = Trim(Replace(UCase(curShipper), "ON BEHALF OF", ""))
                            End If
                            shipperArray = Split(curShipper)
                            Exit For
                        End If
                        
                        'Check for FOR Signifier
                        shipperArray = Split(TrimString(curStringShipper))
                        For Each arrElem In shipperArray
                            If UCase(arrElem) = "FOR" Then
                                forBool = True
                                curShipper = Trim(TrimString(curWorkSheet.Range("D" & saIndex + 1).Value))
                                'Check if shipper is on same line as for or after
                                If curShipper = "" Then
                                    Dim intShipperArray() As String
                                    Dim intArrElem As Variant
                                    Dim incAsShipper As Boolean
                                    curShipper = Trim(TrimString(curWorkSheet.Range("D" & saIndex).Value))
                                    intShipperArray = Split(curShipper)
                                    curShipper = ""
                                    For Each intArrElem In intShipperArray
                                        If incAsShipper Then
                                            curShipper = curShipper & " " & intArrElem
                                        End If
                                        If UCase(intArrElem) = "FOR" Then
                                            incAsShipper = True
                                        End If
                                    Next intArrElem
                                End If
                                Exit For
                            End If
                        Next arrElem
                        
                        If forBool Then
                            shipperArray = Split(Trim(curShipper))
                            Exit For
                        End If
                    Next saIndex
                    
                    'If shipper is not signed for someone, take first line of shipper area to be shipper
                    If forBool = False Then
                        curShipper = Trim(TrimString(curWorkSheet.Range("D" & curWorkSheet.Range("D" & i).End(xlDown).Row).Value))
                    Else
                        curShipper = ""
                        For saIndex = 0 To UBound(shipperArray) - LBound(shipperArray)
                            curShipper = curShipper & " " & shipperArray(saIndex)
                            If InStr(UCase(shipperArray(saIndex)), "LTD") <> 0 Then
                                Exit For
                            End If
                        Next saIndex
                    End If
                    
                End If
                
                '*END FIND SHIPPER-----
                '----------------------
            
                If UCase(Trim(curWorkSheet.Range("C" & i).Value)) = "CONSIGNEE" Then
                    consigneeIndex = i
                    lastRowFinderCountry = curWorkSheet.Range("D" & i).End(xlDown).Row
                    
                    'check if row below this is empty
                    If IsEmpty(curWorkSheet.Range("D" & lastRowFinderCountry + 1)) = False Then
                        lastRowFinderCountry = curWorkSheet.Range("D" & lastRowFinderCountry).End(xlDown).Row
                    End If
                    '-----------------------------------
                    '***FIRST WAVE OF CHECKING COUNTRIES
                    '-----------------------------------
                    For Each curRowinColumn In curWorkSheet.Range("D" & i + 1 & ":D" & lastRowFinderCountry)
                        'Check if row has company country that would mess up parsing
                        If noCompanyCountry(Trim(curRowinColumn)) Then
                            'for each row, check if any of the countries in the list match
                            'Debug.Print curRowinColumn
                            For j = 1 To lastrowCountries
                                'if it does, set curCountry to that country found
                                If InStr(UCase(curRowinColumn), UCase(Trim(mWSCountry.Range("A" & j).Value))) <> 0 And InStr(UCase(curRowinColumn), "VIRGIN AUSTRALIA") = 0 Then
                                    curCountry = mWSCountry.Range("A" & j).Value
                                    countryFound = True
                                    Exit For
                                End If
                            Next j
                        End If
                        'Check if country has been found
                        If countryFound = True Then
                            Exit For
                        End If
                    Next
                    
                    'If country found after first wave, skip second wave and exit loop
                    If countryFound = True Then
                        Exit For
                    End If
                    
                    '-----------------------------------
                    '**SECOND WAVE OF CHECKING COUNTRIES  Specifically 2-letter/3-letter codes
                    '-----------------------------------
                    For Each curRowinColumn In curWorkSheet.Range("D" & i + 1 & ":D" & lastRowFinderCountry)
                        'Check if current line has no company country that would mess parsing
                        If noCompanyCountry(Trim(curRowinColumn)) Then
                            'Split each line into an array of seperate words, using spit,trim,replace to deal w commas etc
                            For Each arrayElement In Split(TrimString(Replace(curRowinColumn, ",", " ")))
                                For j = 1 To lastrowCountries
                                    'If match found, store match into variable and exit loop
                                    If UCase(Trim(arrayElement)) = UCase(mWSCountry.Range("B" & j).Value) Or UCase(arrayElement) = UCase(mWSCountry.Range("C" & j).Value) Then
                                        curCountry = mWSCountry.Range("A" & j).Value
                                        countryFound = True
                                        Exit For
                                    End If
                                Next j
                                'Check if country has been found, exit if true
                                If countryFound = True Then
                                    Exit For
                                End If
                            Next
                        End If
                        
                        'Check if country has been found, exit if true
                        If countryFound = True Then
                                Exit For
                            End If
                    Next
                    
                    'Check if country has been found, exit if true
                    If countryFound = True Then
                        Exit For
                    Else
                        'If code gets here, means no country was found in consignee
                        curCountry = "No country found"
                        errorNote = curCountry
                        errorIndex = errorFileParse(errorNote, FileName, mWS, errorIndex)
                        Exit For
                    End If
                End If
            Next i
            
        'Check if shipper was found or not
        If curShipper = "" Then
            curShipper = "Error in finding shipper"
            errorNote = curShipper
            errorIndex = errorFileParse(errorNote, FileName, mWS, errorIndex)
            GoTo end_of_while_loop:
        End If
        
        'Check if country still found or not
        If countryFound = False Then
            GoTo end_of_while_loop:
        End If
        
        curCountry = checkCountryIf(curCountry)
        
        'Check if country is in the list of accepted countries
        If InStr(countryList, UCase(curCountry)) <> 0 Then
            GoTo find_state:
        Else
            GoTo end_of_while_loop:
        End If
        
'Error handling for country when if checks still do not work
country_error:
        curCountry = "Error in finding country"
        errorNote = err.Description
        errorIndex = errorFileParse(errorNote, FileName, mWS, errorIndex)
        GoTo end_of_while_loop:
    
find_state:
        '-------------------------------------------------------
        '*GO THROUGH CONSIGNEE AREA TO FIND STATE AFTER COUNTRY*
        '-------------------------------------------------------
        Dim stateFinderLetter As String
        Dim stateCodeLetter As String
        Dim stateFound As Boolean
        Dim stateIndexFound As Boolean
        stateIndexFound = False
        stateFound = False
        
        'Loop through country/state list to find index for state parsing
        For i = 4 To mWSCountry.UsedRange.Columns.Count
            If UCase(curCountry) = UCase(mWSCountry.Cells(1, i).Value) Then
                stateIndexFound = True
                stateFinderLetter = Split(mWSCountry.Cells(1, i).Address, "$")(1)
                stateCodeLetter = Split(mWSCountry.Cells(1, i + 1).Address, "$")(1)
            End If
        Next i
        
        'Handle case where no state index was found
        If stateIndexFound = False Then
            curState = "State Index not Found"
            GoTo parse_unnumber:
        End If

        '----------------------------------------------
        'First round of checking state (no state codes)
        '----------------------------------------------
        For Each curRowinColumn In curWorkSheet.Range("D" & consigneeIndex + 1 & ":D" & lastRowFinderCountry)
            'Check if current line is referencing a company
            'Debug.Print curRowinColumn
            If noCompanyCountry(Trim(curRowinColumn)) Then
                'Check through state list of specific country
                For i = 2 To mWSCountry.Range(stateFinderLetter & "30000").End(xlUp).Row
                    If InStr(UCase(curRowinColumn), UCase(Trim(mWSCountry.Range(stateFinderLetter & i).Value))) <> 0 Then
                        stateFound = True
                        curState = Trim(mWSCountry.Range(stateFinderLetter & i).Value)
                        Exit For
                    End If
                Next i
            End If
            
            'Exit if state has been found
            If stateFound Then
                Exit For
            End If
        Next curRowinColumn
        
        '------------------------------------------------
        'Second round of checking state (for state codes)
        '------------------------------------------------
        Dim stateArrEle As Variant
        For Each curRowinColumn In curWorkSheet.Range("D" & consigneeIndex + 1 & ":D" & lastRowFinderCountry)
            'Check if referencing company in row
            If noCompanyCountry(Trim(curRowinColumn)) Then
                'Check through state list of specific country
                For j = 2 To mWSCountry.Range(stateCodeLetter & "30000").End(xlUp).Row
                    'Split the row into an array to check if it contains the state code
                    For Each stateArrEle In Split(Trim(TrimString(Replace(curRowinColumn, ",", " "))))
                        If UCase(stateArrEle) = UCase(Trim(mWSCountry.Range(stateCodeLetter & j).Value)) Then
                            'if state code is found, set state to corresponding code
                            curState = Trim(mWSCountry.Range(stateFinderLetter & j).Value)
                            stateFound = True
                            Exit For
                        End If
                    Next stateArrEle
                    'Exit for loop if state found
                    If stateFound Then
                        Exit For
                    End If
                Next j
            End If
            'Exit for loop if state found
            If stateFound Then
                Exit For
            End If
        Next curRowinColumn
        
        'If no state was found, set the cur state accordingly
        If stateFound = True Then
            GoTo parse_unnumber:
        End If
        
check_for_cities:
        '-----------------------------------------
        '*GO THROUGH LIST OF CITIES TO FIND STATE*
        '-----------------------------------------
        Dim cityFinderLetter As String
        Dim cityCodeLetter As String
        Dim cityIndexFound As Boolean
        cityIndexFound = False
        
        'Loop through state/city list to find index for city parsing
        For i = 1 To mWSCity.UsedRange.Columns.Count
            If UCase(curCountry) = UCase(mWSCity.Cells(1, i).Value) Then
                cityIndexFound = True
                cityFinderLetter = Split(mWSCity.Cells(1, i).Address, "$")(1)
                cityCodeLetter = Split(mWSCity.Cells(1, i + 1).Address, "$")(1)
            End If
        Next i
        
        'Handle case where no state index was found
        If cityIndexFound = False Then
            curState = "City Index not Found"
            GoTo parse_unnumber:
        End If
        
        For Each curRowinColumn In curWorkSheet.Range("D" & consigneeIndex + 1 & ":D" & lastRowFinderCountry)
            'Check if referencing company in row
            If noCompanyCountry(Trim(curRowinColumn)) Then
                'Check through city list of specific country
                For j = 2 To mWSCity.Range(cityCodeLetter & "30000").End(xlUp).Row
                    'Split the row into an array to check if it contains the state code
                    If InStr(UCase(curRowinColumn), UCase(Trim(mWSCity.Range(cityCodeLetter & j).Value))) <> 0 Then
                        stateFound = True
                        curState = Trim(mWSCity.Range(cityFinderLetter & j).Value)
                        Exit For
                    End If
                Next j
            End If
            'Exit for loop if state found
            If stateFound Then
                Exit For
            End If
        Next curRowinColumn
        
        If stateFound = False Then
            curState = "City & State not Found"
        End If
        
parse_unnumber:
        '--------------------------------------------------
        '*GO THROUGH ALL SHEETS WITH TITLE AS FOLLOWS BELOW
        '--------------------------------------------------
        For Each workBookElement In curWorkBook.Sheets
            If workBookElement.Range("D1") = "SHIPPER'S DECLARATION FOR DANGEROUS GOODS" Then
            
                '----------------------------------------
                '*FIND CURRENT UN Number, Class & Weight*
                '----------------------------------------
                'reset the current UNList before we start finding
                Set curUNList = New Collection
                'Loop through column B to find start of UNTable
                For i = 1 To workBookElement.Range("B30000").End(xlUp).Row
                    If InStr(1, Trim(TrimString(workBookElement.Range("B" & i).Value)), "UN or ID No.") <> 0 Then
                        UNTableIndex = i + 2
                        Exit For 'Once starting index has been found, end loop
                    End If
                Next i
                
                'Loop from start of table to find all UN Numbers
                For i = UNTableIndex To workBookElement.Range("B30000").End(xlUp).Row
                    If InStr(Trim(workBookElement.Range("B" & i).Value), "UN") <> 0 Then
                        If Trim(TrimString(workBookElement.Range("B" & i).Value)) = "UN" Then
                            Debug.Print
                            curUNNumber = Trim(TrimString(workBookElement.Range("B" & i).Value & " " & workBookElement.Range("B" & i + 1).Value))
                        Else
                            curUNNumber = Trim(TrimString(workBookElement.Range("B" & i).Value))
                        End If
                        
                        curUNList.Add ParseUNDetails(curUNNumber, i, workBookElement)
                    End If
                Next i
        
                '----------------------------------
                '*ADD DATA TO CONSOLIDATED WORKBOOK
                '----------------------------------
                
                'Loop through UNList & Add to Workbook
                For Each UNListIndex In curUNList
                    If UNListIndex(3) <> 0 Then
                        notEmptySheet = True
                        mWS.Cells(curRow, 1).Value = FileName         ' file name
                        mWS.Cells(curRow, 2).Value = curDate          ' date
                        curCountry = Trim(checkCountryIf(curCountry))
                        mWS.Cells(curRow, 3).Value = Trim(curCountry) ' country
                        mWS.Cells(curRow, 4).Value = Trim(curState)   ' state
                        mWS.Cells(curRow, 5).Value = UNListIndex(0)   ' un number
                        mWS.Cells(curRow, 6).Value = UNListIndex(1)   ' class
                        mWS.Cells(curRow, 7).Value = UNListIndex(2)   ' no boxes
                        mWS.Cells(curRow, 8).Value = UNListIndex(3)   ' weight per box
                        mWS.Cells(curRow, 9).Value = UNListIndex(4)   ' total weight
                        mWS.Cells(curRow, 10).Value = UNListIndex(5)   ' unit
                        mWS.Cells(curRow, 11).Value = Trim(UCase(curCustomer))   ' customer
                        mWS.Cells(curRow, 12).Value = Trim(UCase(curShipper))    ' shipper
                        curRow = curRow + 1
                    End If
                Next
            
            End If
        'End of For loop for each sheet in workbook
        Next
        
        '-------------------
        '*FORMS PER COUNTRY*
        '-------------------
        Dim lastRowCountryForm As Long
        Dim countryExists As Boolean
        Dim noStateCounterExists As Boolean
        
        If notEmptySheet = True Then
        
            countryExists = False
            noStateCounterExists = False
            lastRowCountryForm = mWS.Range("Q50000").End(xlUp).Row
            'curCountry = checkCountryIf(curCountry)
            
            For i = 1 To lastRowCountryForm
                If InStr(Trim(mWS.Range("Q" & i).Value), Trim(curCountry)) <> 0 Then
                    countryExists = True
                    mWS.Range("R" & i).Value = mWS.Range("R" & i).Value + 1
                    'Check if the state was found for this form
                    If stateFound = False Then
                        If mWS.Range("S" & i).Value = "" Then
                            mWS.Range("S" & i).Value = 1
                        Else
                            mWS.Range("S" & i).Value = mWS.Range("S" & i).Value + 1
                        End If
                    End If
                    Exit For
                End If
            Next i
            
            If countryExists = False Then
            mWS.Range("Q" & lastRowCountryForm + 1).Value = Trim(curCountry)
            mWS.Range("R" & lastRowCountryForm + 1).Value = 1
                'Check if the state was found for this form
                If stateFound = False Then
                    mWS.Range("S" & lastRowCountryForm + 1).Value = 1
                End If
            End If
            
        End If
        
end_of_while_loop:
        
        curWorkBook.Close savechanges:=False 'Close current workbook after done parsing data
        FileName = Dir 'Get next workbook and restart loop
        curFileIndex = curFileIndex + 1
        
        
    Loop 'End of While Loop
    
    'MsgBox "Finished Consolidating Data!"
    
    Unload ufProgress
    
cancel_folder_selection:
    
End Sub
' Function to parse item details based on index of sheet and sheet
Public Function ParseUNDetails(unNumber As String, index As Integer, Worksheet As Worksheet) As Variant
    ' variables for array parsing
    Dim arrayElement As Variant
    Dim arrayIndex As Integer
    Dim arrayLength As Integer
    ' variables for array splitting
    Dim splitArray1() As String
    Dim splitArray2() As String
    Dim arr1Length As Long
    Dim arr2Length As Long
    Dim comArray() As Variant
    ' variables to output
    Dim curClass As String
    Dim totWeight As Single
    Dim weightPerBox As Single
    Dim weightUnit As String
    Dim curNoBoxes As Integer
    
    curClass = Worksheet.Range("U" & index).Value
    splitArray1 = Split(TrimString(Trim(Worksheet.Range("AD" & index).Value)))
    splitArray2 = Split(TrimString(Trim(Worksheet.Range("AD" & index + 1).Value)))
    If IsEmpty(Worksheet.Range("AD" & index).Value) = False Then
        comArray = MergeArrays(splitArray1, splitArray2)
        
        'Check number of boxes listed
        If IsNumeric(comArray(0)) Then
            If InStr(comArray(0), ".") <> 0 Then
                curNoBoxes = 1
            Else
                curNoBoxes = comArray(0)
            End If
        End If
        
        'Check for weight per box
        arrayLength = UBound(comArray) - LBound(comArray)
        For arrayIndex = 0 To arrayLength
            If InStr(1, comArray(arrayIndex), ".") <> 0 Then
                weightPerBox = CSng(comArray(arrayIndex))
                weightUnit = comArray(arrayIndex + 1)
            End If
        Next arrayIndex
    Else
        curNoBoxes = 1
        weightPerBox = 0
        weightUnit = "Error"
    End If
        
    'Calculate total weight
    totWeight = curNoBoxes * weightPerBox
    
    'Return the array of details
    ParseUNDetails = Array(unNumber, curClass, curNoBoxes, weightPerBox, totWeight, weightUnit)
End Function

' Function to merge two arrays together
Public Function MergeArrays(arr1 As Variant, arr2 As Variant) As Variant
    Dim i As Long
    Dim arr As Variant
    ReDim arr(UBound(arr1) + UBound(arr2) + 1)

    For i = LBound(arr1) To UBound(arr1)
        arr(i) = arr1(i)
    Next i
    
    For i = LBound(arr2) To UBound(arr2)
        arr(i + UBound(arr1)) = arr2(i)
    Next i
    MergeArrays = arr
End Function

' Function to trim spaces in string to just one space
Public Function TrimString(strInput As String)
    Dim strTemp As String
 
    strTemp = strInput
 
    Do
        If InStr(1, strTemp, vbLf) <> 0 Then
            strTemp = Replace(strTemp, vbLf, "")
        ElseIf InStr(1, strTemp, "  ") > 0 Then
            strTemp = Replace(strTemp, "  ", " ")
        Else
            Exit Do
        End If
    Loop

    TrimString = strTemp
End Function

Public Function errorFileParse(errorString As String, errorFileName As Variant, mainSheet As Worksheet, rowIndex As Integer) As Integer
    mainSheet.Cells(rowIndex, 14).Value = errorFileName
    mainSheet.Cells(rowIndex, 15).Value = errorString
    errorFileParse = rowIndex + 1
End Function

Function GetFolder() As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function

Private Function CountFilesInFolder(strDir As String, Optional strType As String)
'DEVELOPER: Ryan Wells (wellsr.com)
'DESCRIPTION: This macro counts the files in a folder and retuns the result in a msgbox
'INPUT: Pass the procedure a string with your directory path and an optional
' file extension with the * wildcard
'EXAMPLES: Call CountFilesInFolder("C:\Users\Ryan\")
' Call CountFilesInFolder("C:\Users\Ryan\", "*txt")
    Dim file As Variant, i As Integer
    If Right(strDir, 1) <> "\" Then strDir = strDir & "\"
    file = Dir(strDir & strType)
    While (file <> "")
        i = i + 1
        file = Dir
    Wend
    CountFilesInFolder = i
End Function

'Add in countries that need to be replaced
Private Function checkCountryIf(countryIN As String)
    If countryIN = "United States" Then
        checkCountryIf = "United States of America"
    ElseIf countryIN = "Viet Nam" Then
        checkCountryIf = "Vietnam"
    Else
        checkCountryIf = countryIN
    End If
End Function

'Add in checks for signifiers that the line contains a country
Private Function noCompanyCountry(stringIN As String)
    Dim signalFound As Boolean
    Dim stringArr() As String
    Dim stringArrEle As Variant
    
    stringArr = Split(Trim(TrimString(stringIN)), " ")
    For Each stringArrEle In stringArr
        If InStr(UCase(compSignalList), UCase(stringArrEle)) <> 0 Or InStr(UCase(stringArrEle), "LIMITED") <> 0 Then
            Debug.Print stringArrEle
            signalFound = True
            Exit For
        End If
    Next stringArrEle
    
    If signalFound = True Then
        noCompanyCountry = False
    Else
        noCompanyCountry = True
    End If
    
End Function





