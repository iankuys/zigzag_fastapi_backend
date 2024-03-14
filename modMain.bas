Attribute VB_Name = "modMain"
Option Explicit

Dim NPsychMaster As String

Const SQLMasterData As String = _
        "Provider=SQLOLEDB;" _
         & "Server=Spinal;" _
         & "Database=IBACohort;" _
         & "Integrated Security=SSPI;" _
         & "DataTypeCompatibility=80;" _
         & "MARS Connection=True;"
         
         
Dim app As Application

Dim YArray(28) As Single
Dim XCenter As Single
Dim XStdDev As Single
Dim XMax As Single
Dim XMin As Single
Dim LabelScores(80, 1) As Variant
Dim TestStdDevs(28) As Variant
''Dim TestScores1(28) As Variant
Public aPID As Integer
Public aSID As Integer
Public aVnum As Integer
Public aColor As Integer

Sub DataEntry(pid)
    aPID = 0
    Load UserForm1
    UserForm1.txtPID = pid
    UserForm1.Show
    
End Sub

'Public Sub SetDBMaster(db As String)
Public Sub SetDBMaster()
    NPsychMaster = SQLMasterData
    
End Sub

Public Sub SetSubject(pid, sid, vnum, color)
    aPID = pid
    aSID = sid
    aVnum = vnum
    aColor = color
    initialize
End Sub

Sub initialize()
Dim i As Integer
Dim SQL As String
    
    
    ''2022-08-11 DKH
    For i = 0 To 28
        YArray(i) = 138.75 + (19.5 * i)
    Next i
    For i = 0 To UBound(LabelScores)
        LabelScores(i, 0) = Null
        LabelScores(i, 1) = Null
    Next i
    
    XCenter = 321.5
    XStdDev = 56.6
    XMax = (XStdDev * 3) 'actually the -5 stddev at right of graph
    XMin = (XStdDev * -3) 'actually the +3 stddev at left of graph
    
    If aPID = 0 Then
        If UserForm1.txtBlack <> "" Then
            SetTestScores UserForm1.txtPID, UserForm1.txtBlack, 3 '2084, 2, 0, 1
            DrawTestScores 0, 0, 0
        End If
        If UserForm1.txtBlue <> "" Then
            SetTestScores UserForm1.txtPID, UserForm1.txtBlue, 2 '2084, 2, 0, 1
            DrawTestScores 0, 0, 255
        End If
        If UserForm1.txtRed <> "" Then
            SetTestScores UserForm1.txtPID, UserForm1.txtRed, 1 '2084, 2, 0, 1
            DrawTestScores 255, 0, 0
            UpdateScoreFields
        End If
    Else
        If aColor = 3 Then
            SetTestScores aPID, aVnum, 3 '2084, 2, 0, 1
            DrawTestScores 0, 0, 0
        End If
        If aColor = 2 Then
            SetTestScores aPID, aVnum, 2 '2084, 2, 0, 1
            DrawTestScores 0, 0, 255
        End If
        If aColor = 1 Then
            SetTestScores aPID, aVnum, 1 '2084, 2, 0, 1
            DrawTestScores 255, 0, 0
            UpdateScoreFields
        End If
    End If
End Sub

Sub CreateLine(x1 As Single, y1 As Single, x2 As Single, y2 As Single, r As Integer, g As Integer, b As Integer)
Dim mydocument As Slide

    Set mydocument = ActivePresentation.Slides(1)
    With mydocument.Shapes.AddLine(x1, y1, x2, y2).Line
        .DashStyle = msoLineSolid
        .Weight = 2
        .ForeColor.RGB = RGB(r, g, b)
    End With
    
End Sub

Sub UpdateScoreFields()
Dim daShape As Shape, i As Integer
Dim mydocument As Slide
    
    Set mydocument = ActivePresentation.Slides(1)
    
    For i = 0 To UBound(LabelScores, 1)
        If Not IsNull(LabelScores(i, 0)) Then
            Set daShape = mydocument.Shapes.Item(LabelScores(i, 0))
            daShape.TextFrame.TextRange.Text = CStr(LabelScores(i, 1))
        End If
    Next i
    
End Sub

Sub SetTestScores(pid As Integer, vnum As Integer, flgLineSet As Integer)
Dim conn As ADODB.Connection
Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset
Dim retval As Long, i As Integer
Dim mydocument As Slide
Dim total As Variant
Dim age As Integer
Dim edu As Integer
Dim sex As Integer
Dim mean As Integer
Dim sd As Integer
Dim connstring As String
Dim varCutOff As Integer

''2020-09-04 DKH
''To test and run without Access DB
''In the Immediate Windows, run the following commands
''SetDBMaster  <-- sets up access to SQL Database
''DataEntry  <--run the 3 lines function...

''SetDBMaster
''SetSubject 4000, 1, 0, 1 <--Run single line chart
''Patient 4000, SID 1, Visit 0, Color Red 1


''2021-01-05  View items labels for VBA
''https://stackoverflow.com/questions/52940674/how-to-find-the-label-of-a-shape-on-a-powerpoint-slide
''On the Home tab in Powerpoint there is an Editing submenu that contains a Select function. As pictured here:
''When you click on Select, another submenu appears that shows a Selection Pane function. If you click on that, the selction pane shows up on the right hand of the screen. In that pane you will see all of the objects on the current slide and the names Powerpoint has given to each of them.


    Set mydocument = ActivePresentation.Slides(1)

    Set conn = New ADODB.Connection
    Set cmd = New ADODB.Command

    conn.Open NPsychMaster
    cmd.ActiveConnection = conn
            
    ''updated to StoreProcedure... quick hack to combine C2 and TCog.
    ''DKH 2021-08-16
    cmd.CommandText = "spNPsychSelectedPatientZigZag_v2022"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("pPatientID", adInteger, adParamInput, , CStr(pid))
    cmd.Parameters.Append cmd.CreateParameter("pVisitNumber", adInteger, adParamInput, , CStr(vnum))
    
    Set rs = cmd.Execute(retval)
    
    age = rs("age")
    edu = rs("edu")
    sex = rs("Sex")
       
    ''TestStdDev(index)
    ''Shape
    ''ScoreValues
    
    ''Don't draw line if SD greate than varCutOff. arbitrarily using 5. DKH 2021-01-05
    ''handling these + data status codes
    ''88/888=Optional item
    ''95/995=Physical problem
    ''96/996=Cognitive/behavior problem
    ''97/997=Other problem
    ''98/998=Verbal refusal
    
    varCutOff = 5
       
    ''Setup Word List items:
    
    Dim varWordlist As Integer
    varWordlist = IIf(IsNull(rs("Wordlist Test")), 0, rs("Wordlist Test"))
    
    ''Word Test Label
    mydocument.Shapes("txtlblWordlist").TextFrame.TextRange.Text = IIf(IsNull(rs("Wordlist Test Label")), "", rs("Wordlist Test Label"))
    
    If varWordlist = 1 Then ''RAVLT
      ''TestItem1
      TestStdDevs(0) = rs("RAVLT Total Learning SD")
      mydocument.Shapes("pp01_Score").TextFrame.TextRange.Text = "(" & rs("RAVLT Total Learning") & ")"
      setGenericScoreValuesAgeTestSection "TotalLearn", age, "01", "[Npsych].[tlkpRAVLTAgeMeanStdDev]"
    
      mydocument.Shapes("pp01_ScoreLabels").TextFrame.TextRange.Text = "(" & rs("RAVLT 1-5 Total Learning Scores") & ")"
    
    
      ''TestItem2
      TestStdDevs(1) = rs("RAVLT Trial6 SD")
      mydocument.Shapes("pp02_Score").TextFrame.TextRange.Text = "(" & rs("RAVLT Trial6") & ")"
      setGenericScoreValuesAgeTestSection "Trial6", age, "02", "[Npsych].[tlkpRAVLTAgeMeanStdDev]"
      
      ''TestItem3
      TestStdDevs(2) = rs("RAVLT Long-Delay Recall SD")
      mydocument.Shapes("pp03_Score").TextFrame.TextRange.Text = "(" & rs("RAVLT Long-Delay Recall") & ")"
      setGenericScoreValuesAgeTestSection "Trial7", age, "03", "[Npsych].[tlkpRAVLTAgeMeanStdDev]"
      
      ''TestItem4
      TestStdDevs(3) = rs("RAVLT Long-Delay Recog True Pos SD")
      mydocument.Shapes("pp04_Score").TextFrame.TextRange.Text = "(" & rs("RAVLT Long-Delay Recog True Pos") & ")"
      setGenericScoreValuesAgeTestSection "DelayedRecog", age, "04", "[Npsych].[tlkpRAVLTAgeMeanStdDev]"
        
      mydocument.Shapes("pp04_ScoreLabels").TextFrame.TextRange.Text = "(false-positive errors=" & rs("RAVLT Long-Delay Recog False Pos") & ")"
      
    ElseIf varWordlist = 2 Then ''CERAD
      ''HardCoded Mean and SD from [IBACohortDE_UDS20].[Npsych].[tlkpCERADWL_MeanStdDev] DKH 2022-08-17
      ''TestItem1
      TestStdDevs(0) = rs("CWLT TT SD")
      mydocument.Shapes("pp01_Score").TextFrame.TextRange.Text = "(" & rs("CWLT TT") & ")"
      setGenericScoreValuesMeanSD 20.9, 3.9, "01"
      
    
      ''TestItem2
      TestStdDevs(1) = rs("CWLT 5mDT SD")
      mydocument.Shapes("pp02_Score").TextFrame.TextRange.Text = "(" & rs("CWLT 5mDT") & ")"
      setGenericScoreValuesMeanSD 7.2, 1.8, "02"
      
      ''TestItem3
      TestStdDevs(2) = rs("CWLT 30mDT SD")
      mydocument.Shapes("pp03_Score").TextFrame.TextRange.Text = "(" & rs("CWLT 30mDT") & ")"
      setGenericScoreValuesMeanSD 7, 1.8, "03"
      
      ''TestItem4
      TestStdDevs(3) = rs("CWLT 30mRT SD")
      mydocument.Shapes("pp04_Score").TextFrame.TextRange.Text = "(" & rs("CWLT 30mRT") & ")"
      setGenericScoreValuesMeanSD 19.6, 0.5, "04"
        
    ElseIf varWordlist = 2 Then ''CVLT
      ''TestItem1
      TestStdDevs(0) = Null
      mydocument.Shapes("pp01_Score").TextFrame.TextRange.Text = "(" & rs("CVLT TrialTS") & ")"
      'setGenericScoreValuesAgeTestSection "TotalLearn", age, "01", "tlkpRAVLTAgeMeanStdDev"
    
      ''TestItem2
      TestStdDevs(1) = Null
      mydocument.Shapes("pp02_Score").TextFrame.TextRange.Text = "(" & rs("CVLT ShortDelayFree") & ")"
      'setGenericScoreValuesAgeTestSection "Trial6", age, "02", "tlkpRAVLTAgeMeanStdDev"
      
      ''TestItem3
      TestStdDevs(2) = Null
      mydocument.Shapes("pp03_Score").TextFrame.TextRange.Text = "(" & rs("CVLT LongDelayFree") & ")"
      'setGenericScoreValuesAgeTestSection "Trial7", age, "03", "tlkpRAVLTAgeMeanStdDev"
      
      ''TestItem4
      TestStdDevs(3) = Null
      mydocument.Shapes("pp04_Score").TextFrame.TextRange.Text = "(" & rs("CVLT LongDelayRecogHits") & ")"
      'setGenericScoreValuesAgeTestSection "DelayedRecog", age, "04", "tlkpRAVLTAgeMeanStdDev"

    Else
    ''Other
      TestStdDevs(0) = Null
      'mydocument.Shapes("pp01_Score").TextFrame.TextRange.Text = "(" & rs("CWLT TT") & ")"
      'setGenericScoreValuesAgeTestSection "TotalLearn", age, "01", "tlkpRAVLTAgeMeanStdDev"
    
      ''TestItem2
      TestStdDevs(1) = Null
      'mydocument.Shapes("pp02_Score").TextFrame.TextRange.Text = "(" & rs("CWLT 5mDT") & ")"
      'setGenericScoreValuesAgeTestSection "Trial6", age, "02", "tlkpRAVLTAgeMeanStdDev"
      
      ''TestItem3
      TestStdDevs(2) = Null
      'mydocument.Shapes("pp03_Score").TextFrame.TextRange.Text = "(" & rs("CWLT 30mDT") & ")"
      'setGenericScoreValuesAgeTestSection "Trial7", age, "03", "tlkpRAVLTAgeMeanStdDev"
      
      ''TestItem4
      TestStdDevs(3) = Null
      'mydocument.Shapes("pp04_Score").TextFrame.TextRange.Text = "(" & rs("CWLT 30mRT") & ")"
      'setGenericScoreValuesAgeTestSection "DelayedRecog", age, "04", "tlkpRAVLTAgeMeanStdDev"
        
    End If
    
    ''TestItem5
    TestStdDevs(4) = rs("CRAFT Immediate Paraphrase ZScore")
    mydocument.Shapes("pp05_Score").TextFrame.TextRange.Text = "(" & rs("CRAFT Immediate Paraphrase") & ")"
    setGenericScoreValuesSexAgeEdu sex, age, edu, "05", "[UDS3].[tlkpCraftStoryRecall-Immediate-Paraphrase_SexAgeEduStdDev]"
    
    ''TestItem6
    TestStdDevs(5) = rs("CRAFT Delayed Paraphrase ZScore")
    mydocument.Shapes("pp06_Score").TextFrame.TextRange.Text = "(" & rs("CRAFT Delayed Paraphrase") & ")"
    setGenericScoreValuesSexAgeEdu sex, age, edu, "06", "[UDS3].[tlkpCraftStoryRecall-Delayed-Paraphrase_SexAgeEduStdDev]"
    
    ''TestItem7; update to DUFF calculations DKH 2023-10-13
    ''TestStdDevs(6) = rs("BVMT TotalRecall ZScore")
    TestStdDevs(6) = rs("BVMT DUFF TotalRecall ZScore")
    ''mydocument.Shapes("pp07_Score").TextFrame.TextRange.Text = "(" & rs("BVMT TotalRecall TScore") & ")"
    mydocument.Shapes("pp07_Score").TextFrame.TextRange.Text = "(" & rs("BVMT DUFF TotalRecall TScore") & ")"
    setGenericScoreValuesMeanSD 50, 10, "07"
    'setGenericScoreValuesAgeTestSection "ListB", age, "07", "tlkpRAVLTAgeMeanStdDev"
    
    mydocument.Shapes("pp07_ScoreLabels").TextFrame.TextRange.Text = "(" & rs("BVMT 1-3 Total Learning Scores") & ") (" & rs("BVMT TotalRecall") & ")"
      
    '
    ''TestItem8
    ''TestStdDevs(7) = rs("BVMT Delayed Recall ZScore")
    TestStdDevs(7) = rs("BVMT DUFF Delayed Recall ZScore")
    ''mydocument.Shapes("pp08_Score").TextFrame.TextRange.Text = "(" & rs("BVMT Delayed Recall TScore") & ")"
    mydocument.Shapes("pp08_Score").TextFrame.TextRange.Text = "(" & rs("BVMT DUFF Delayed Recall TScore") & ")"
    setGenericScoreValuesMeanSD 50, 10, "08"
    'setGenericScoreValuesAgeTestSection "Trial6", age, "08", "tlkpRAVLTAgeMeanStdDev"
    
    mydocument.Shapes("pp08_ScoreLabels").TextFrame.TextRange.Text = "(" & rs("BVMT Delayed Recall") & ")"
    
    ''TestItem9
    ''TestStdDevs(8) = rs("BVMT Hits ZScore")
    If IsNull(rs("BVMT Hits")) Then
        TestStdDevs(8) = Null
        mydocument.Shapes("pp09_Score").TextFrame.TextRange.Text = ""
    Else
        TestStdDevs(8) = setBVMTHits_ZScoreValuesAge(age, "09", rs("BVMT Hits"))
        mydocument.Shapes("pp09_Score").TextFrame.TextRange.Text = "(" & rs("BVMT False Alarm") & "))" & "(" & rs("BVMT Hits") & ")"
        setBVMTHits_ScoreValuesAge age, "09"
    End If
    
    ''TestItem10
    TestStdDevs(9) = rs("BENSON CFT Delayed Recall ZScore")
    mydocument.Shapes("pp10_Score").TextFrame.TextRange.Text = "(" & rs("BENSON CFT Delayed Recall") & ")"
    setGenericScoreValuesSexAgeEdu sex, age, edu, "10", "[UDS3].[tlkpBensonComplexFigure-Recall_SexAgeEduStdDev]"

    ''TestItem11
    TestStdDevs(10) = rs("TOPF StdS SD")
    mydocument.Shapes("pp11_Score").TextFrame.TextRange.Text = "(" & rs("TOPF StdS") & ")"
    ' pass values to standard score (mean:100, std:15)
    setGenericScoreValuesMeanSD 100, 15, "11"
    '    setGenericScoreValuesAgeTestSection "Error", age, 11, "tlkpRAVLTAgeMeanStdDev", True

    'Test Line 12
    TestStdDevs(11) = rs("NumSpan DIGFORSL ZScore")
    mydocument.Shapes("pp12_Score").TextFrame.TextRange.Text = "(" & rs("NumSpan DIGFORSL") & ")"
    setGenericScoreValuesSexAgeEdu sex, age, edu, "12", "[UDS3].[tlkpNumberSpanTest-Forward-Longest-Span_SexAgeEduStdDev]"
    
    'Digit Span Backwards Length
    'Test Line 13
    TestStdDevs(12) = rs("NumSpan DIGBACLS ZScore")
    mydocument.Shapes("pp13_Score").TextFrame.TextRange.Text = "(" & rs("NumSpan DIGBACLS") & ")"
    setGenericScoreValuesSexAgeEdu sex, age, edu, "13", "[UDS3].[tlkpNumberSpanTest-Backward-Longest-Span_SexAgeEduStdDev]"
      
    '    'Test Line 14
    TestStdDevs(13) = rs("SDMT SD")
    mydocument.Shapes("pp14_Score").TextFrame.TextRange.Text = "(" & rs("SDMT #Written") & ")"
    setGenericScoreValuesAgeEdu age, edu, 14, "[Npsych].[tlkpSymbolDigitModalityAgeEduMeanStdDev]"

    '    'Test Line 15
    TestStdDevs(14) = rs("DKEFS C2 WordReading ZScore")
    mydocument.Shapes("pp15_Score").TextFrame.TextRange.Text = "(" & rs("DKEFS C2 WordReading SS") & ")"
    ' pass values to standard score (mean:100, std:15)
    setGenericScoreValuesMeanSD 10, 3, "15"
    mydocument.Shapes("pp15_ScoreLabels").TextFrame.TextRange.Text = "(" & rs("DKEFS C2 WordReading") & ")"
    
    'Test Line 16
    TestStdDevs(15) = rs("DKEFS C1 ColorNaming ZScore")
    mydocument.Shapes("pp16_Score").TextFrame.TextRange.Text = "(" & rs("DKEFS C1 ColorNaming SS") & ")"
    ' pass values to standard score (mean:100, std:15)
    setGenericScoreValuesMeanSD 10, 3, "16"
    mydocument.Shapes("pp16_ScoreLabels").TextFrame.TextRange.Text = "(" & rs("DKEFS C1 ColorNaming") & ")"

    'Test Line 17
    TestStdDevs(16) = rs("TrailsA ZScore")
    mydocument.Shapes("pp17_Score").TextFrame.TextRange.Text = "(" & rs("TrailsA Sec") & ")"
    setGenericScoreValuesSexAgeEdu sex, age, edu, "17", "[UDS3].[tlkpTrailMaking-PartA_SexAgeEduStdDev]", True
    'setGenericScoreValuesMeanSD 10, 3, "17"
    
    'Test Line 18
    TestStdDevs(17) = rs("KDC SecToComp SD")
    mydocument.Shapes("pp18_Score").TextFrame.TextRange.Text = "(" & rs("KDC SecToComp") & ")"
    setGenericScoreValuesAge age, "18", "[Npsych].[tlkpKendrickDigitCopyAgeMeanStdDev]"
    
    ''Test Line 19
    ''Multilingual Naming Test
    TestStdDevs(18) = rs("MINT ZScore")
    mydocument.Shapes("pp19_Score").TextFrame.TextRange.Text = "(" & rs("MINT TS") & ")"
    setGenericScoreValuesSexAgeEdu sex, age, edu, "19", "[UDS3].[tlkpMINT_SexAgeEduStdDev]"

    'Test Line 20
    'Letter Fluency FAS Scaled Score
    TestStdDevs(19) = rs("FAS SD")
    
    If IsNumeric(rs("FAS SD")) = True Then
        mydocument.Shapes("pp20_Score").TextFrame.TextRange.Text = "(" & CInt(rs("FAS SS")) & ")"
    Else
        mydocument.Shapes("pp20_Score").TextFrame.TextRange.Text = "(" & rs("FAS SS") & ")"
    End If
    'setGenericScoreValuesMeanSD 10, 3, "20"
    mydocument.Shapes("pp20_ScoreLabels").TextFrame.TextRange.Text = "(" & rs("FAS 1-3 Total Learning Scores") & ")  (" & rs("FAS TS") & ")"
    
    
    'Test Line 21
    'Setup Category Fluency
    TestStdDevs(20) = rs("CCF ZScore")
    mydocument.Shapes("pp21_Score").TextFrame.TextRange.Text = "(" & rs("CCF TS") & ")"
    setGenericScoreValuesSexAgeEdu sex, age, edu, "21", "[UDS3].[tlkpCategoryFluency-Animals_SexAgeEduStdDev]"
    
    ''Test Line 22
    TestStdDevs(21) = rs("BENSON CFT Drawing ZScore")
    mydocument.Shapes("pp22_Score").TextFrame.TextRange.Text = "(" & rs("BENSON CFT Drawing") & ")"
    setGenericScoreValuesSexAgeEdu sex, age, edu, "22", "[UDS3].[tlkpBensonComplexFigure-Copy_SexAgeEduStdDev]"
    
    ''Test Line 23
    TestStdDevs(22) = rs("BVMT Copy ZScore")
    mydocument.Shapes("pp23_Score").TextFrame.TextRange.Text = "(" & rs("BVMT Copy") & ")"
    setGenericScoreValuesMeanSD 11, 1.175, "23"
    mydocument.Shapes("pp23_99").TextFrame.TextRange.Text = ""
    mydocument.Shapes("pp23_95").TextFrame.TextRange.Text = ""
    'mydocument.Shapes("pp23_85").TextFrame.TextRange.Text = ""
    'mydocument.Shapes("pp23_50").TextFrame.TextRange.Text = ""
    'mydocument.Shapes("pp23_15").TextFrame.TextRange.Text = ""
    'mydocument.Shapes("pp23_05").TextFrame.TextRange.Text = ""
    'mydocument.Shapes("pp23_01").TextFrame.TextRange.Text = ""
    
    
    ''Test Line 24
    TestStdDevs(23) = rs("WAIS4 BD SD")
    mydocument.Shapes("pp24_Score").TextFrame.TextRange.Text = "(" & rs("WAIS4 BD SS") & ")"
    setGenericScoreValuesMeanSD 10, 3, "24"
    
    mydocument.Shapes("pp24_ScoreLabels").TextFrame.TextRange.Text = "(JLO=" & rs("jlo_raw") & "/30) (Read/Set Time=" & rs("Clock ReadSetTime") & "/6)  (" & rs("WAIS4 BD") & ")"
        
    ''Test Line 25
    TestStdDevs(24) = rs("TrailsB ZScore")
    mydocument.Shapes("pp25_Score").TextFrame.TextRange.Text = "(" & rs("TrailsB Sec") & ")"
    setGenericScoreValuesSexAgeEdu sex, age, edu, "25", "[UDS3].[tlkpTrailMaking-PartB_SexAgeEduStdDev]", True
    ''setGenericScoreValuesMeanSD 10, 3, "25"
    
    mydocument.Shapes("pp25_ScoreLabels").TextFrame.TextRange.Text = "(" & rs("TrailsB Errs") & " errors)"
    
    
    ''Test Line 26
    TestStdDevs(25) = rs("DKEFS C3 Inhibition ZScore")
    mydocument.Shapes("pp26_Score").TextFrame.TextRange.Text = "(" & rs("DKEFS C3 Inhibition SS") & ")"
    setGenericScoreValuesMeanSD 10, 3, "26"
    
    mydocument.Shapes("pp26_ScoreLabels").TextFrame.TextRange.Text = "(" & rs("DKEFS C3 Inhibition Uncorrected Errors") & " uncorrected errors) (" & rs("DKEFS C3 Inhibition") & ")"
    
    ''Test Line 27
    TestStdDevs(26) = rs("DKEFS C4 InhibitionSwitching ZScore")
    mydocument.Shapes("pp27_Score").TextFrame.TextRange.Text = "(" & rs("DKEFS C4 InhibitionSwitching SS") & ")"
    setGenericScoreValuesMeanSD 10, 3, "27"
    
    mydocument.Shapes("pp27_ScoreLabels").TextFrame.TextRange.Text = "(" & rs("DKEFS C4 InhibitionSwitching Uncorrected Errors") & " uncorrected errors) (" & rs("DKEFS C4 InhibitionSwitching") & ")"
    
    ''Test Line 28
    If IsNull(rs("WCST64 Categories Complete Percentile")) Then
        TestStdDevs(27) = Null
        mydocument.Shapes("pp28_Score").TextFrame.TextRange.Text = ""
    Else
        TestStdDevs(27) = setWCST64_ZScoreValues_AgeEdu(age, edu, rs("WCST64 Categories Complete Percentile"), "NumberOfCategories", "[Npsych].[tlkpWCST64_AgeEduTScore]")
        mydocument.Shapes("pp28_Score").TextFrame.TextRange.Text = "(" & rs("WCST64 Categories Complete") & ")"
        setWCST64ScoreValuesAgeEdu age, edu, "28", "NumberOfCategories", "[Npsych].[tlkpWCST64_AgeEduTScore]"
    End If
    
    ''Test Line 29
    If IsNull(rs("WCST64 Categories Complete Percentile")) Then
        TestStdDevs(28) = Null
        mydocument.Shapes("pp29_Score").TextFrame.TextRange.Text = ""
        setGenericScoreValuesMeanSD 50, 10, "29"
    Else
        TestStdDevs(28) = rs("WCST64 Perseverative Errors ZScore")
        mydocument.Shapes("pp29_Score").TextFrame.TextRange.Text = "(" & rs("WCST64 Perseverative Errors TScore") & ")"
        setGenericScoreValuesMeanSD 50, 10, "29" ''generic for T-Score
        
        mydocument.Shapes("pp29_ScoreLabels").TextFrame.TextRange.Text = "(" & rs("WCST64 Perseverative Errors") & ")"
        
    End If
    

    'added GDS 2011-05-31 DKH
    mydocument.Shapes("txtScoreGDS").TextFrame.TextRange.Text = "(" & rs("GDS GDS") & "/15)"

    'moved insight rating 2017-09-14 DKH
    mydocument.Shapes("txtScoreInsight").TextFrame.TextRange.Text = "(" & rs("Insight Rating") & ")"

    
    mydocument.Shapes("txtPatientName").TextFrame.TextRange.Text = CStr(rs("Ptname"))
    mydocument.Shapes("txtPatientID").TextFrame.TextRange.Text = CStr(pid)
    
    Set rs = cmd.Execute(retval)
    
    If flgLineSet = 1 Then
        mydocument.Shapes("txtEDate").TextFrame.TextRange.Text = CStr(Format(rs("examdate"), "MM/DD/YYYY"))
        mydocument.Shapes("txtYEAR1").TextFrame.TextRange.Text = CStr(Format(rs("examdate"), "YYYY"))
    ElseIf flgLineSet = 2 Then
        mydocument.Shapes("txtYEAR2").TextFrame.TextRange.Text = CStr(Format(rs("examdate"), "YYYY"))
    Else
        mydocument.Shapes("txtYEAR3").TextFrame.TextRange.Text = CStr(Format(rs("examdate"), "YYYY"))
    End If
    
    rs.Close
    conn.Close
    
    For i = 1 To UBound(TestStdDevs) + 1
        If TestStdDevs(i - 1) < -3# Then
            TestStdDevs(i - 1) = -3#
        End If
        If TestStdDevs(i - 1) > varCutOff Then
            TestStdDevs(i - 1) = Null
        End If
        If TestStdDevs(i - 1) > 3# Then
            TestStdDevs(i - 1) = 3#
        End If
    Next
    
End Sub

Sub CreateBullet(x As Single, y As Single, s As Integer, r As Integer, g As Integer, b As Integer)
Dim daShape As Shape
Dim mydocument As Slide

    Set mydocument = ActivePresentation.Slides(1)
    Set daShape = mydocument.Shapes.AddShape(msoShapeDiamond, x, y, s, s)
    daShape.Fill.ForeColor.RGB = RGB(r, g, b)
    daShape.Line.ForeColor.RGB = RGB(r, g, b)
    
End Sub

Sub DrawTestScores(r As Integer, g As Integer, b As Integer)
Dim i As Integer
Dim x As Single, x1 As Single
Dim y As Single, y1 As Single
    
    'draw lines first
    For i = 1 To UBound(TestStdDevs)
        If Not IsNull(TestStdDevs(i - 1)) Then
            If Not IsNull(TestStdDevs(i)) Then
                y = YArray(i - 1)
                y1 = YArray(i)
                x = TestStdDevs(i - 1) * -1
                x1 = TestStdDevs(i) * -1
                
                If x < -3 Then
                    x = -3
                ElseIf x > 5 Then
                    x = 5
                End If
                
                If x1 < -3 Then
                    x1 = -3
                ElseIf x1 > 5 Then
                    x1 = 5
                End If
                
                CreateLine XCenter + x * XStdDev, y, XCenter + x1 * XStdDev, y1, r, g, b
            End If
        End If
    Next
    
    'draw Bullets last
    For i = 0 To UBound(TestStdDevs)
        If Not IsNull(TestStdDevs(i)) Then
            y = YArray(i)
            x = TestStdDevs(i) * -1
            
            If x < -3 Then
                x = -3
            ElseIf x > 5 Then
                x = 5
            End If
            
            If r = 0 And g = 0 And b = 0 Then
                CreateBullet XCenter - 5 + x * XStdDev, y - 5, 10, r, g, b
            ElseIf r = 0 And g = 0 Then
                CreateBullet XCenter - 4 + x * XStdDev, y - 4, 8, r, g, b
            Else
                CreateBullet XCenter - 3 + x * XStdDev, y - 3, 6, r, g, b
            End If
        End If
    Next
    
End Sub






