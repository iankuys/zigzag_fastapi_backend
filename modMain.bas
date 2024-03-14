Attribute VB_Name = "modMain"
Option Explicit

Dim NPsychMaster As String

Const SQLMasterData As String = _
        "Provider=SQLOLEDB;" _
         & "Server=Spinal;" _
         & "Database=IBACohortReports;" _
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
Dim TestScores1(28) As Variant
Public aPID As Integer
Public aSID As Integer
Public aVnum As Integer
Public aColor As Integer
     
Sub DataEntry()
    aPID = 0
    Load UserForm1
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

    Set mydocument = ActivePresentation.Slides(1)

    Set conn = New ADODB.Connection
    Set cmd = New ADODB.Command

    'conn.Provider = "Microsoft.Jet.OLEDB.4.0"
    'conn.Properties("Jet OLEDB:System Database") = "z:\alzdb.mdw"
    conn.Open NPsychMaster '"Data Source=" & NPsychMaster , "report", "report"

    cmd.ActiveConnection = conn
    
    'Find age at visit date
    If pid >= 30000 And pid < 40000 Then ''REDCap CADC 2017-01-27 DKH.
        cmd.CommandText = "SELECT [Age] AS age, [total_edu_yrs] AS edu FROM [REDCapImports].[pid184].[vwNpsychDemo] WHERE CADCID = " & pid
    Else
        cmd.CommandText = "Select IBAUtilities.dbo.AgeInYears(pt.birthdate, pv.PhysicalExamDate, pv.NeuropsychExamDate) as age, df.EducationYears as edu, df.sex AS Sex from IBACohort.dbo.tblPatients pt Inner Join IBACohort.dbo.tblPatientVisits pv on pt.patientid=pv.patientid left join IBACohort.dbo.tblDemographicsFull df on pt.patientid=df.patientid where pt.PatientID = " & pid & " And pv.VisitNumber = " & vnum
    End If
    
    cmd.CommandType = adCmdText
    
    Set rs = cmd.Execute(retval)
    
    If Not rs.EOF Then
        age = rs("age")
        edu = rs("edu")
        If rs("sex") = "M" Then
            sex = 1
        ElseIf rs("Sex") = "F" Then
            sex = 2
        End If
        rs.Close
    Else
        MsgBox "Error unable to find subject record to compute Age. Exiting..."
        Exit Sub
    End If
    
    'posibly changing vwNPsychStdBattRecords to tblCVLT
    cmd.CommandText = "select * from IBACohort.dbo.vwNPsychStdBattRecords where PatientID=" & pid & " and VisitNumber=" & vnum
    cmd.CommandType = adCmdText
    
    Set rs = cmd.Execute(retval)
    
    
    TestStdDevs(0) = rs("MMSE SD")
    mydocument.Shapes("TxtBoxMMSE").TextFrame.TextRange.Text = "(" & rs("MMSE MMSE") & ")"
    
    TestStdDevs(1) = rs("MOCA SD")
    mydocument.Shapes("sc1").TextFrame.TextRange.Text = "(" & rs("MoCA TS") & ")"
    setGenericScoreValues age, edu, 1, "tlkpMoCAAgeEduMeanStdDev"
    
    
    TestStdDevs(2) = rs("CWLT T1TSD")
    mydocument.Shapes("txtBoxTrial1").TextFrame.TextRange.Text = "(" & rs("CWLT T1T") & ")"
    
    TestStdDevs(3) = rs("CWLT T2TSD")
    mydocument.Shapes("TxtBoxTrial2").TextFrame.TextRange.Text = "(" & rs("CWLT T2T") & ")"
    
    TestStdDevs(4) = rs("CWLT T3TSD")
    mydocument.Shapes("TxtBoxTrial3").TextFrame.TextRange.Text = "(" & rs("CWLT T3T") & ")"
    
    TestStdDevs(5) = rs("CWLT 5mDTSD")
    mydocument.Shapes("TxtBox5MinRecall").TextFrame.TextRange.Text = "(" & rs("CWLT 5mDT") & ")"
    
    TestStdDevs(6) = rs("CWLT 30mDTSD")
    mydocument.Shapes("TxtBox30MinRecall").TextFrame.TextRange.Text = "(" & rs("CWLT 30mDT") & ")"
    
    TestStdDevs(7) = IIf(rs("CWLT 30mRTSD") < -1# And rs("CWLT 30mRTSD") > -2#, -1, rs("CWLT 30mRTSD"))
    mydocument.Shapes("TxtBox30MinRecognition").TextFrame.TextRange.Text = "(" & rs("CWLT 30mRT") & ")"
        
    'Test Line 8
    If rs("CRAFTImmediate SD") = Null Then
        TestStdDevs(8) = ((rs("WMS3LM1 StARaw") - 13.9) / 3.9)
        mydocument.Shapes("sc8").TextFrame.TextRange.Text = "(" & rs("WMS3LM1 StARaw") & ")"
    Else
        TestStdDevs(8) = rs("CRAFTImmediate SD")
        mydocument.Shapes("sc8").TextFrame.TextRange.Text = "(" & rs("CRAFT Immediate Paraphrase") & ")"
    End If
    setGenericScoreValues age, edu, 8, "tlkpCRAFTImmediateAgeEduMeanStdDev"
    
    
    'Test Line 9
    If rs("CRAFTDelayed SD") = Null Then
        TestStdDevs(9) = ((rs("WMS3LM2 StARaw") - 12.6) / 4.3)
        mydocument.Shapes("sc9").TextFrame.TextRange.Text = "(" & rs("WMS3LM2 StARaw") & ")"
    Else
        TestStdDevs(9) = rs("CRAFTDelayed SD")
        mydocument.Shapes("sc9").TextFrame.TextRange.Text = "(" & rs("CRAFT Delayed Paraphrase") & ")"
    End If
    setGenericScoreValues age, edu, 9, "tlkpCRAFTDelayedAgeEduMeanStdDev"
    
    
    'Test Line 10
    TestStdDevs(10) = rs("BENSONDelayRecall SD") ' - Benson CFT: Delayed Recall
    mydocument.Shapes("sc10").TextFrame.TextRange.Text = "(" & rs("BENSON CFT Delayed Recall") & ")"
    setBensonCFTDelay age, edu, sex, 10
    
    'Test Line 11
    TestStdDevs(11) = rs("WMS3F1 SD")
    mydocument.Shapes("txtBoxWMS3Faces1").TextFrame.TextRange.Text = "(" & rs("WMS3F1 SS") & ")"
    
    'Test Line 12
    TestStdDevs(12) = rs("WMS3F2 SD")
    mydocument.Shapes("txtBoxWMS3Faces2").TextFrame.TextRange.Text = "(" & rs("WMS3F2 SS") & ")"
    
    'Test Line 13
    TestStdDevs(13) = rs("WAISRInfo SD")
    
    'Test Line 14
    ''Modified added -19 and check isnull DKH 2017-05-05
    If IsNull(rs("NumSpan DIGFORSL SD")) Or rs("NumSpan DIGFORSL SD") = -19 Then
        TestStdDevs(14) = Round(rs("WAIS3DS Fwd Len SD"), 6)
        mydocument.Shapes("sc14").TextFrame.TextRange.Text = "(" & rs("WAIS3DS Fwd Len") & ")"
    Else
        TestStdDevs(14) = rs("NumSpan DIGFORSL SD")
        mydocument.Shapes("sc14").TextFrame.TextRange.Text = "(" & rs("NumSpan DIGFORSL") & ")"
    End If
    
    setGenericScoreValues age, edu, 14, "tlkpNumSpanForwardAgeEduMeanStdDev"
    
    'Digit Span Backwards Length
    'Test Line 15
    ''Modified added -19 and check isnull DKH 2017-05-05
    If IsNull(rs("NumSpan DIGBACLS SD")) Or rs("NumSpan DIGBACLS SD") = -19 Then
        TestStdDevs(15) = rs("WAIS3DS Bkwd Len SD")
        mydocument.Shapes("sc15").TextFrame.TextRange.Text = "(" & rs("WAIS3DS Bkwd Len") & ")"
    Else
        TestStdDevs(15) = rs("NumSpan DIGBACLS SD")
        mydocument.Shapes("sc15").TextFrame.TextRange.Text = "(" & rs("NumSpan DIGBACLS") & ")"
    End If
    
    setGenericScoreValues age, edu, 15, "tlkpNumSpanBackwardAgeEduMeanStdDev"
       
    
    TestStdDevs(16) = rs("SDMT SD")
    mydocument.Shapes("sc16").TextFrame.TextRange.Text = "(" & rs("SDMT #Written") & ")"
    setGenericScoreValues age, edu, 16, "tlkpSymbolDigitModalityAgeEduMeanStdDev"
    
    'Test Line 17
    If rs("MINT SD") = Null Then
        TestStdDevs(17) = rs("BN30 SD")
        mydocument.Shapes("sc17").TextFrame.TextRange.Text = "(" & rs("BNTScore") & ")"
    Else
        TestStdDevs(17) = rs("MINT SD")
        mydocument.Shapes("sc17").TextFrame.TextRange.Text = "(" & rs("MINT TS") & ")"
    End If
    setGenericScoreValues age, edu, 17, "tlkpMINTAgeEduMeanStdDev"
    
'    'Cindy Tran 2015-08-06
'    'updated to change MINT score to Boston Naming / MINT with bold
'    mydocument.Shapes("sc16").TextFrame.TextRange.Text = "(" & rs("MINT TS") & "/" & rs("BNTScore") & ")"
'
'    Dim StartPos As Integer
'    Dim EndPos As Integer
'    StartPos = InStr(1, mydocument.Shapes("sc16").TextFrame.TextRange, "/")
'    EndPos = InStr(1, mydocument.Shapes("sc16").TextFrame.TextRange, ")")
'    mydocument.Shapes("sc16").TextFrame.TextRange.Characters(StartPos, EndPos - StartPos).Font.Bold = True

    TestStdDevs(18) = rs("FAS SD")
    mydocument.Shapes("sc18").TextFrame.TextRange.Text = "(" & rs("FAS SS") & ")"
    
    'Setup Category Fluency sc19
    TestStdDevs(19) = rs("CCF SD")
    mydocument.Shapes("sc19").TextFrame.TextRange.Text = "(" & rs("CCF TS") & ")"
    setGenericScoreValues age, edu, 19, "tlkpCategoryFluencyAgeEduMeanStdDev"

'Tina - 05/14/2015 - add TestStdDevs(19) = rs("BENSONDraw SD") & rows shift down
    TestStdDevs(20) = rs("BENSONDraw SD")
    mydocument.Shapes("sc20").TextFrame.TextRange.Text = "(" & rs("BENSON CFT Drawing") & ")"
    setGenericScoreValues age, edu, 20, "tlkpBCFTDrawAgeEduStdDev"
    
    TestStdDevs(21) = rs("ConstPrax SD")
    
    TestStdDevs(22) = rs("Clock SD")
    
    If Not IsNull(rs("WAIS3BD SD")) Then
        TestStdDevs(23) = rs("WAIS3BD SD")
    Else
        TestStdDevs(23) = Null
    End If
    
    
    TestStdDevs(24) = rs("Judge SD")
    
    
    'Test Line 25
    'WAISR-NI or WAIS3Sim
    'WAISR-NI first then WAIS3Sim
    
'    If Not IsNull(rs("WAISR NI Raw")) And rs("WAISR NI Raw") >= 0 Then
'        If Not IsNull(rs("WAISR NI SD")) Then
'            TestStdDevs(25) = rs("WAISR NI SD")
'        Else
'            TestStdDevs(25) = Null
'        End If
'    Else
     If Not IsNull(rs("WAIS3Sim SD")) Then
            TestStdDevs(25) = rs("WAIS3Sim SD")
        Else
            TestStdDevs(25) = Null
        End If
'    End If
    

    TestStdDevs(26) = rs("TrailsA SD")
    TestStdDevs(27) = rs("TrailsB SD")
    TestStdDevs(28) = rs("KDC SD")  'Kendrick Digit - Test #28 (lbl28, sc28, pp28*)
    setGenericScoreValuesAge age, 28, "tlkpKendrickDigitCopyAgeMeanStdDev"
    
    
'********************
        
    'added GDS 2011-05-31 DKH
    mydocument.Shapes("txtBoxGDSScore").TextFrame.TextRange.Text = "(" & rs("GDS GDS") & ")"
    
    'moved insight rating 2017-09-14 DKH
    mydocument.Shapes("scInsight").TextFrame.TextRange.Text = "(" & rs("Insight Rating") & ")"
    
    mydocument.Shapes("txtBoxWAISRInformation").TextFrame.TextRange.Text = "(" & rs("WAISRInfo SS") & ")"
    
    mydocument.Shapes("TxtBoxCERADDrawing").TextFrame.TextRange.Text = "(" & rs("ConstPrax TS") & ")"
    
    mydocument.Shapes("TxtBoxReadSetTime").TextFrame.TextRange.Text = "(" & rs("Clock TS") & ")"
    
    If Not IsNull(rs("WAIS3BD SS")) Then
        mydocument.Shapes("TxtBoxWAIS3BlockDesign").TextFrame.TextRange.Text = "(" & rs("WAIS3BD SS") & ")"
    Else
        mydocument.Shapes("TxtBoxWAIS3BlockDesign").TextFrame.TextRange.Text = "(_)"
    End If
    
'    mydocument.Shapes("TxtBoxInsight").TextFrame.TextRange.Text = "(" & rs("Insight Rating") & ")"
    mydocument.Shapes("TxtBoxJudgement").TextFrame.TextRange.Text = "(" & rs("Judge TS") & ")"
    
    
'    If rs("WAISR NI SD") = Null Or rs("WAISR NI SS") = -10 Then ''DKH edited to show only WAIS3 if NI is null or -10  2015-12-09
        If Not IsNull(rs("WAIS3Sim SS")) Then
            mydocument.Shapes("TxtBoxSimilarities").TextFrame.TextRange.Text = "(" & rs("WAIS3Sim SS") & ")"
        Else
            mydocument.Shapes("TxtBoxSimilarities").TextFrame.TextRange.Text = "(_)"
        End If
'    Else
'        If Not IsNull(rs("WAISR NI SS")) Then
'            mydocument.Shapes("TxtBoxSimilarities").TextFrame.TextRange.Text = "(" & rs("WAISR NI SS") & ")"
'        Else
'            mydocument.Shapes("TxtBoxSimilarities").TextFrame.TextRange.Text = "(_)"
'        End If
'    End If
     
    
    mydocument.Shapes("TxtBoxTrailsA").TextFrame.TextRange.Text = "(" & rs("TrailsA SS") & ")"
    mydocument.Shapes("TxtBoxTrailsB").TextFrame.TextRange.Text = "(" & rs("TrailsB SS") & ")"
    mydocument.Shapes("sc28").TextFrame.TextRange.Text = "(" & rs("KDC SecToComp") & "/" & rs("KDC #Comp2m") & ")"
    


'********************
'DKH comment out 2015-05-20
'    If Not IsNull(rs("WAIS3DS SS")) Then
'    Else
'        LabelScores(57, 0) = 45
'        LabelScores(57, 1) = "WAIS     Digit Span"
'    End If
'
'    If Not IsNull(rs("WAIS3Sim SS")) Then
'    Else
'        LabelScores(56, 0) = 55
'        LabelScores(56, 1) = "WAIS-III Similarities"
'    End If
    
    ' next 58
    If Not IsNull(rs("WAIS3BD SS")) Then
    Else
        LabelScores(58, 0) = 53
        LabelScores(58, 1) = "WAIS     Blk Design"
    End If
    
    'rs.Close
    
    'cmd.CommandText = "select ptname from qryNameVisitDate where pid=" & pid & " and sid=" & sid
    'cmd.CommandType = adCmdText
    
    'Set rs = cmd.Execute(retval)
    
    'Tina - 05/18/2015 - ' - Text Box 2 - lbl_pname
'    mydocument.Shapes("txt").TextFrame.TextRange.Text = " "
    'LabelScores(63, 0) = 19 ''- pname
    'LabelScores(63, 1) = CStr(rs("Ptname"))
    mydocument.Shapes("txtPatientName").TextFrame.TextRange.Text = CStr(rs("Ptname"))
    'LabelScores(67, 0) = 18
    'LabelScores(67, 1) = CStr(pid)
    mydocument.Shapes("txtPatientID").TextFrame.TextRange.Text = CStr(pid)
    
    'rs.Close
    
    'cmd.CommandText = "select * from qryNameVisitDate where pid=" & pid & " and sid=" & sid & " and vnum=" & vnum
    'cmd.CommandType = adCmdText
    
    Set rs = cmd.Execute(retval)
    
    
    If flgLineSet = 1 Then
'        LabelScores(65, 0) = 132
'        LabelScores(65, 1) = CStr(Format(rs("examdate"), "YYYY"))
        mydocument.Shapes("txtYEAR1").TextFrame.TextRange.Text = CStr(Format(rs("examdate"), "YYYY"))

'Tina - 06/30/2015
'        LabelScores(64, 0) = 20
'        LabelScores(64, 1) = CStr(rs("examdate"))
         mydocument.Shapes("txtEDate").TextFrame.TextRange.Text = CStr(Format(rs("examdate"), "MM/DD/YYYY"))
    
    ElseIf flgLineSet = 2 Then
'        LabelScores(66, 0) = 133
'        LabelScores(66, 1) = CStr(Format(rs("examdate"), "YYYY"))
        mydocument.Shapes("txtYEAR2").TextFrame.TextRange.Text = CStr(Format(rs("examdate"), "YYYY"))
    Else
'        LabelScores(62, 0) = 134
'        LabelScores(62, 1) = CStr(Format(rs("examdate"), "YYYY"))
        mydocument.Shapes("txtYEAR3").TextFrame.TextRange.Text = CStr(Format(rs("examdate"), "YYYY"))
    End If
    
    rs.Close
    conn.Close
    
    For i = 1 To UBound(TestStdDevs) + 1
        If TestStdDevs(i - 1) < -3# Then
            TestStdDevs(i - 1) = -3#
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
Sub setGenericScoreValues(age As Integer, edu As Integer, scoreitem As Integer, strTable As String)
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim conn As ADODB.Connection
    Dim retval As Long

    Set conn = New ADODB.Connection
    Set cmd = New ADODB.Command

    conn.Open NPsychMaster

    cmd.ActiveConnection = conn
        
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT * FROM [IBACohort].[Npsych].[" & strTable & "] " + _
        "WHERE (min_age <= " + CStr(age) + " AND max_age >= " + CStr(age) + ")" + _
        "AND (min_edu <= " + CStr(edu) + " AND max_edu >= " + CStr(edu) + ");"
        
    Set rs = cmd.Execute(retval)
    
    Dim mydocument As Slide
    Set mydocument = ActivePresentation.Slides(1)
    
    If Not rs.EOF Then
    
    ''Setting 99 as blank, since SD3plus is max 'DKH 2017-09-01
    'mydocument.Shapes("pp" & scoreitem & "_99").TextFrame.TextRange.Text = ""
    mydocument.Shapes("pp" & scoreitem & "_95").TextFrame.TextRange.Text = rs("SD2plus") & ""  ''<== trick to set null string
    mydocument.Shapes("pp" & scoreitem & "_85").TextFrame.TextRange.Text = rs("SD1plus")
    mydocument.Shapes("pp" & scoreitem & "_50").TextFrame.TextRange.Text = Round(rs("Mean"))
    mydocument.Shapes("pp" & scoreitem & "_15").TextFrame.TextRange.Text = rs("SD1minus")
    mydocument.Shapes("pp" & scoreitem & "_05").TextFrame.TextRange.Text = rs("SD2minus")
    mydocument.Shapes("pp" & scoreitem & "_01").TextFrame.TextRange.Text = rs("SD3minus")

    End If

End Sub

Sub setGenericScoreValuesAge(age As Integer, scoreitem As Integer, strTable As String)
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim conn As ADODB.Connection
    Dim retval As Long

    Set conn = New ADODB.Connection
    Set cmd = New ADODB.Command

    conn.Open NPsychMaster

    cmd.ActiveConnection = conn
        
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT * FROM [IBACohort].[Npsych].[" & strTable & "] " + _
        "WHERE (min_age <= " + CStr(age) + " AND max_age >= " + CStr(age) + ");"
        
    Set rs = cmd.Execute(retval)
    
    Dim mydocument As Slide
    Set mydocument = ActivePresentation.Slides(1)
    
    If Not rs.EOF Then
    
    ''Setting 99 as blank, since SD3plus is max 'DKH 2017-09-01
    'mydocument.Shapes("pp" & scoreitem & "_99").TextFrame.TextRange.Text = ""
    mydocument.Shapes("pp" & scoreitem & "_95").TextFrame.TextRange.Text = rs("SD2plus") & ""  ''<== trick to set null string
    mydocument.Shapes("pp" & scoreitem & "_85").TextFrame.TextRange.Text = rs("SD1plus")
    mydocument.Shapes("pp" & scoreitem & "_50").TextFrame.TextRange.Text = Round(rs("Mean"))
    mydocument.Shapes("pp" & scoreitem & "_15").TextFrame.TextRange.Text = rs("SD1minus")
    mydocument.Shapes("pp" & scoreitem & "_05").TextFrame.TextRange.Text = rs("SD2minus")
    mydocument.Shapes("pp" & scoreitem & "_01").TextFrame.TextRange.Text = rs("SD3minus")

    End If

End Sub


Sub setBensonCFTDelay(age As Integer, edu As Integer, sex As Integer, scoreitem As Integer)
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim conn As ADODB.Connection
    Dim retval As Long

    Set conn = New ADODB.Connection
    Set cmd = New ADODB.Command

    conn.Open NPsychMaster

    cmd.ActiveConnection = conn
        
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT * FROM [IBACohort].[Npsych].[tlkpBCFTDelayedSexAgeEduStdDev] " + _
        "WHERE (min_age <= " + CStr(age) + " AND max_age >= " + CStr(age) + ")" + _
        " AND (min_edu <= " + CStr(edu) + " AND max_edu >= " + CStr(edu) + ")" + _
        " AND (sex = " + CStr(sex) + ");"
        
    Set rs = cmd.Execute(retval)
    
    Dim mydocument As Slide
    Set mydocument = ActivePresentation.Slides(1)
    
    If Not rs.EOF Then
    
    ''Setting 99 & 95 as blank, since SD1plus is max 'DKH 2017-09-01
    'mydocument.Shapes("pp" & scoreitem & "_99").TextFrame.TextRange.Text = rs("SD3plus") & ""  ''<== trick to set null string
    mydocument.Shapes("pp" & scoreitem & "_95").TextFrame.TextRange.Text = rs("SD2plus") & ""
    mydocument.Shapes("pp" & scoreitem & "_85").TextFrame.TextRange.Text = rs("SD1plus")
    mydocument.Shapes("pp" & scoreitem & "_50").TextFrame.TextRange.Text = Round(rs("Mean"))
    mydocument.Shapes("pp" & scoreitem & "_15").TextFrame.TextRange.Text = rs("SD1minus")
    mydocument.Shapes("pp" & scoreitem & "_05").TextFrame.TextRange.Text = rs("SD2minus")
    mydocument.Shapes("pp" & scoreitem & "_01").TextFrame.TextRange.Text = rs("SD3minus") & ""

    End If
    
End Sub
