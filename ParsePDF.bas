Attribute VB_Name = "ParsePDF"
Sub ParsePDF()

' REQUIRED REFERENCES:      (Tools > References) _
    Microsoft Scripting Runtime         supports dictionary object



' DECLARATIONS

    '' Variables
    Dim i As Integer
    Dim j As Integer
    Dim NonProductSKUs As Variant
    Dim BrandNames As Dictionary
    
    
    '' Arrays
    NonProductSKUs = Array("CALIBRATE AUDIO", "CALIBRATE VIDEO", "Credit Card Transa...", "CRESTRON", "Discount", "DONATION", "EXTRON", "Hardware", "Item", "LABOR", "LABOR/Materials", "LABOR/TESTING", "Materials/Costs", "RENTAL", "REPAIR", "Restocking Fee", "Retainage", "SEC-ACTIVATE", "SECMON36-SMART", "SECMON60-SMART", "SERVICE CONTRACT", "SHIPPING", "TESTING", "TRAINING", "TRAVEL", "2GIGBASIC", "ACCPOINT", "ACTIVATESERV", "AUDIOCAL", "CALIBPROJ", "CALIBTV", "CAM", "COMPSETUP", "DEMO", "ELECTRICAL", "HDMIBALUN", "LAMP", "LVPANEL14", "LVPANEL28", "LVPANEL42", "MAINT-LEVEL1", "MAINT-LEVEL2", "MAINT-LEVEL3", "MUSICCAST", "OVRC-CONFIG", "POWEREXT", "PROGHARMONY", "PROGREM1", "PROGREM2", "PROGREM3", "PROGREM4", _
                     "PROGREMREV", "PROJ", "PW5.1", "PW7.2", "PWATMOS", "PWCAT5", "PWCAT6", "PWHDMI", "PWPROJ", "PWRG6", "PWRG6CAT5", "PWSPKZONE", "RACKSETUP18", "RACKSETUP44", "RELOAVCOMP", "RELOROUTER", "RETROCAM", "RETROCAT5", "RETROHDMI", "RETRORG6", "RETRORG6CAT5", "RETROSPK", "ROUTER", "SCREEN", "SECDVR", "SECINSPECT", "SECREINSTATE", "SETUPAV", "SETUPAVSTAND", "SHADE", "SONOS", "SONOSBAR", "SOUNDBAR", "SPKFS", "SPKIC", "SPKIW", "SPKOD", "SPKSM", "SWITCH", "TOTALCONTR...", "TOUCHPANEL", "TRIMCAT5", "TROUBLEST", "TV40", "TV41-65", "TV66-84", "TVANTENNA", "TVART-40", "TVART41-65", "TVART66-84", "TVHIDECABLES", "VC", "SWITCH24")

    Set BrandNames = New Dictionary
        BrandNames.Add "Audio", "Audio Technica"
        BrandNames.Add "Beale", "Beale St"
        BrandNames.Add "Definitive", "Definitive Technology"
        BrandNames.Add "Key", "Key Digital"
        BrandNames.Add "Listen", "Listen Technology"
        BrandNames.Add "Middle", "Middle Atlantic"
        BrandNames.Add "Origin", "Origin Acoustics"
        BrandNames.Add "Repair", "Repair Master"
        BrandNames.Add "Screen", "Screen Innovations"
    
    
    
' PREPARATION

    '' Disable Events
    screenUpdateState = Application.ScreenUpdating
    statusBarState = Application.DisplayStatusBar
    calcState = Application.Calculation
    eventsState = Application.EnableEvents
    displayPageBreakState = ActiveSheet.DisplayPageBreaks
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False

    '' Copy Worksheet
    Cells.EntireRow.AutoFit
    Worksheets(1).Copy After:=Worksheets(1)

    '' Remove Formatting
    Cells.Select
    Selection.UnMerge
    With Selection
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .ClearFormats
        .Font.Bold = False
        .Font.Color = vbBlack
    End With


' COLUMN MANAGEMENT

    '' Define Columns
    DefineRange ("QTY")
    DefineRange ("Item")
    DefineRange ("Description")
    
    '' Loop through columns
    i = FindLast("Column")
    Do While i > 0
    
        ''' Preserve Item, Description, QTY Columns
        Select Case (i)
            Case Range("ItemColumn").Column
            
            Case Range("DescriptionColumn").Column
            
            Case Range("QTYColumn").Column
        
        ''' Delete all other columns
            Case Else
                Columns(i).Delete
        End Select
        i = i - 1
    Loop

    '' Move QTY Column                  (First Column)
    Range("QTYColumn").Columns(1).Cut
    Range("A:A").Insert Shift:=xlToRight
    
    '' Create and Name Brand Column     (New First Column)
    Range("A1").EntireColumn.Insert
    Range("A1").Value = "Brand"
    DefineRange ("Brand")
    
    '' Name Project Column              (Last Column)
    Cells(1, Chr(65 + FindLast("Column"))).Value = "Project"
    DefineRange ("Project")
    
    '' Name Project Column              (Last Column)
'    Cells(1, Chr(65 + FindLast("Column"))).Value = "Trello"
'    DefineRange ("Trello")



' ROW MANAGEMENT
    
    '' Loop (A) through all Rows
    i = FindLast("Row")
    Do While i > 0
        

        
        ''' Delete Non-Product rows
        If Range("ItemColumn")(i).Value = 0 Or _
           Range("DescriptionColumn")(i).Value = 0 Or _
           IsInArray(Range("ItemColumn")(i).Value, NonProductSKUs) Then
            Rows(i).Delete

        ''' For Product Rows:
        Else
        
            '''' Populate Brand, Project, and Trello Columns
            Range("BrandColumn")(i).Value = Split(Range("DescriptionColumn")(i).Value)
            Range("ProjectColumn")(i).Value = Left(ActiveWorkbook.Name, (InStrRev(ActiveWorkbook.Name, ".", -1, vbTextCompare) - 1))
'            Range("TrelloColumn")(i).Value = "=QTYColumn & "") "" & BrandColumn & "" "" & ItemColumn"
            
            '''' Flag Truncated Items
            If InStr(1, Range("ItemColumn")(i).Value, "...") > 1 Then
                With Range("ItemColumn")(i)
                    .Font.Color = vbBlue
                End With
            End If
     
            ''' Loop (B) - Combine Identical Product
            j = i - 1
            Do While j > 0
                        
                '''' Rows with Matching Items AND Descriptions
                If Range("ItemColumn")(i).Value = Range("ItemColumn")(j).Value And _
                   Range("DescriptionColumn")(i).Value = Range("DescriptionColumn")(j).Value Then
                   
                    ''''' Add QTY to primary row
                    Range("QTYColumn")(i).Value = Range("QTYColumn")(i).Value + Range("QTYColumn")(j).Value
                   
                    ''''' Delete secondary row
                    Rows(j).Delete
                    
                '''' Matching Items, Different Descriptions
                ElseIf Range("ItemColumn")(i).Value = Range("ItemColumn")(j).Value And _
                       Range("DescriptionColumn")(i).Value <> Range("DescriptionColumn")(j).Value Then
                       
                    ''''' Flag Non-Duplicate Descriptions
                    With Range("ItemColumn")(i)
                        .Font.Color = vbRed
                    End With
                    
                    With Range("ItemColumn")(j)
                        .Font.Color = vbRed
                    End With
                End If
                
                '''' Reduce Loop (B) Counter
                j = j - 1
            Loop

        End If

        ''' Reduce Loop (A) Counter
        i = i - 1
        
    ''Repeat Loop A
    Loop
        
    
' CLEANUP

    '' Sort by Brand, Description
    i = FindLast("Column")
    Columns("A:" & Chr(64 + i)).Sort _
        key1:=Columns(Range("BrandColumn").Column), _
        order1:=xlAscending, _
        key2:=Columns(Range("ItemColumn").Column), _
        order2:=xlAscending, _
        Header:=xlNo
    '' Note: Chr(64 + x) returns the alpha reference for x
     
    '' Formatting
    Cells.EntireRow.AutoFit
    Cells.EntireColumn.AutoFit
    Range("DescriptionColumn").ColumnWidth = 25
    Range("A1").Select
    Application.CutCopyMode = False
    
    '' Resume Events
    Application.ScreenUpdating = screenUpdateState
    Application.DisplayStatusBar = statusBarState
    Application.Calculation = calcState
    Application.EnableEvents = eventsState
    ActiveSheet.DisplayPageBreaks = displayPageBreaksState
    
End Sub


Sub DefineRange(Text As String)

    Dim ColumnNumber As Long
    ColumnNumber = FindText(Text)
    
    ActiveWorkbook.Names.Add _
        Name:=Text & "Column", _
        RefersTo:=Range(Chr(64 + ColumnNumber) & ":" & Chr(64 + ColumnNumber))
        
End Sub

Function FindLast(RowOrColumn As String) As Long

    Select Case (RowOrColumn)
    
        Case "Column"
            FindLast = Cells.Find(What:="*", _
                After:=Range("A1"), _
                LookAt:=xlPart, _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByColumns, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False).Column
        
        Case "Row"
            FindLast = Cells.Find(What:="*", _
                After:=Range("A1"), _
                LookAt:=xlPart, _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False).Row
                    
        Case Else
            FindLast = 0
    End Select
    
End Function

Function FindText(Text As String) As Long

    FindText = Cells.Find(What:=Text, _
        After:=Range("A1"), _
        LookAt:=xlWhole, _
        LookIn:=xlFormulas, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlNext, _
        MatchCase:=False).Column

End Function

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean

  IsInArray = Not IsError(Application.Match(stringToBeFound, arr, 0))

End Function


Sub ReportRunTime()
'PURPOSE: Determine how many seconds it took for code to completely run
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim StartTime As Double
Dim SecondsElapsed As Double
Dim LastRow As Long

LastRow = Cells.Find(What:="*", _
                After:=Range("A1"), _
                LookAt:=xlPart, _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False).Row

'Remember time when macro starts
  StartTime = Timer

ParsePDF

'Determine how many seconds code took to run
  SecondsElapsed = Round(Timer - StartTime, 2)

'Notify user in seconds
  MsgBox "This code processed " & LastRow & " rows in " & SecondsElapsed & " seconds", vbInformation

End Sub


