Option Compare Database

Sub saubern()
    Dim s As String
    Dim o As String
    Dim outputtags As String
    Dim a() As String
    Dim aa() As String
    Dim excl1() As Variant
    Dim excl2Instr() As Variant
    Dim i As Integer
    Dim ii As Integer
    Dim rst As DAO.Recordset
    Dim found As Boolean
    Dim taglist_templ As String
    Set form1 = Forms("Database Tools")

    input_table = form1.Text14.Value
    output_table = input_table & "Cl"
    taglist_templ = form1.Text49.Value
    taglistInput_templ = "taglistInput_templ"

    strPath = CurrentProject.FullName

    If Not IsNull(DLookup("Type", "MSYSObjects", "Name='" & output_table & "'")) Then
        MsgBox "Query " & output_table & ": an object already exists with this name, deleting old one."
        DoCmd.DeleteObject acTable, output_table
    End If
    DoCmd.TransferDatabase acImport, "Microsoft Access", strPath, acTable, taglistInput_templ, output_table, True

    Set db = CurrentDb()
    Set rst = db.OpenRecordset("Select * From " & input_table & " Order By UserID Asc, Tags Asc", dbOpenDynaset)
    Set exclude1 = db.OpenRecordset(form1.Text18.Value, dbOpenDynaset)
    Set exclude2instr = db.OpenRecordset(form1.Text21.Value, dbOpenDynaset)
    Set rst2 = db.OpenRecordset(output_table, dbOpenDynaset) 'Zieltabelle

    '''Zeiger an den Anfang der Tabellen setzen'''
If Not (rst.BOF And rst.EOF) Then
        rst.MoveFirst
    End If
    If Not (exclude1.BOF And exclude1.EOF) Then
        exclude1.MoveFirst
    End If
    If Not (exclude2instr.BOF And exclude2instr.EOF) Then
        exclude2instr.MoveFirst
    End If
    If Not (rst2.BOF And rst2.EOF) Then
        rst2.MoveFirst
    End If

    '''Exclude-Arrays einlesen'''

    excl1 = exclude1.GetRows(1000)
    excl2Instr = exclude2instr.GetRows(1000)

    rst.MoveLast
    rst.MoveFirst

    o = rst("UserID").Value
    s = rst("Tags").Value
    a = Split(s, ";")

    Do While Not rst.EOF
        If IsNull(rst("Tags").Value) Or rst("Tags").Value = ";;" Then
            GoTo skip
        End If

        If Not o = rst("UserID").Value Then 'erste Zeile Änderung Owner, tags einlesen und übertragen
            o = rst("UserID").Value
            s = rst("Tags").Value
            a = Split(s, ";")
            output_tags = ""
            ''''''für jedes Tag in Zeile:''''''
            For i = 0 To UBound(a, 1)
                'Prüfen, ob Tag in Exclude-Listen
                found = False
                If Len(a(i)) <= 1 Or IsNumeric(a(i)) = True Then found = True
                If found = False Then
                    For ii = 0 To UBound(excl2Instr, 2)
                        If InStr(1, a(i), excl2Instr(1, ii)) > 0 Then
                            found = True
                            Exit For
                        End If
                    Next ii
                End If
                If found = False Then
                    For ii = 0 To UBound(excl1, 2)
                        If a(i) = excl1(1, ii) Then
                            found = True
                            Exit For
                        End If
                    Next ii
                End If

                'Einzelnse Tags aussortieren und in Tabelle 2 (rst2) untereinander schreiben,
                'falls nicht in einer der exclude-Tabellen
                If Not a(i) = "" And found = False Then
                    output_tags = output_tags & ";" & a(i)
                End If
            Next i
            ''''''Ende für jedes Tag in Zeile''''''
        Else 'wenn derselbe Owner!
            ' process weitere Tags des Users für andere Bilder, übernehme nur neue, noch nicht verwendete
            s = rst("Tags").Value
            aa = Split(s, ";")
            output_tags = ""
            For i = 0 To UBound(aa, 1) 'für jedes tag in aa (aktuelle tagliste)
                found = False
                For ii = 0 To UBound(a, 1) 'prüfe, ob tag schon in ausgangstagliste (a) vorhanden ist
                    If a(ii) = aa(i) Then found = True
                Next ii

                'Prüfen, ob Tag in Exclude-Listen
                If Len(aa(i)) <= 1 Or IsNumeric(aa(i)) = True Then found = True
                If found = False Then
                    For ii = 0 To UBound(excl2Instr, 2)
                        If InStr(1, aa(i), excl2Instr(1, ii)) > 0 Then
                            found = True
                            Exit For
                        End If
                    Next ii
                End If
                If found = False Then
                    For ii = 0 To UBound(excl1, 2)
                        If aa(i) = excl1(1, ii) Then
                            found = True
                            Exit For
                        End If
                    Next ii
                End If

                If found = False Then
                    output_tags = output_tags & ";" & aa(i) 'Füge Tag hinzu wenn es nicht schon vorhanden ist
                    ReDim Preserve a(0 To UBound(a) + 1) As String 'Füge Tag der Ausgangstagliste hinzu!
       a(UBound(a)) = aa(i)
                End If

            Next i

            'rst2.Edit

        End If

        If Not output_tags = "" Then 'Speichern der relevanten Tags, wenn keine Tags mehr f. Foto, dann keine Datenausgabe
            rst2.AddNew
            rst2("Tags").Value = output_tags & ";"
            rst2("UserID").Value = rst("UserID").Value
            rst2("PhotoID").Value = rst("PhotoID").Value
            rst2.Update
        End If

skip:
        rst.MoveNext 'Nächste TagListRow
    Loop

    '''Close all open tables'''
    rst.Close
    exclude1.Close
    exclude2instr.Close
    rst2.Close
    Set rst = Nothing
    Set exclude1 = Nothing
    Set exclude2instr = Nothing
    Set rst2 = Nothing

    form1.progress.Caption = "1/1 Done."
    form1.Text71.Value = output_table
    form1.Text23.Value = "Taglist_" & output_table
End Sub



Public Sub runsplit()
    Dim s As String
    Dim a() As String
    Dim excl1() As Variant
    Dim excl2Instr() As Variant
    Dim i As Integer
    Dim ii As Integer
    Dim rst As DAO.Recordset
    Dim found As Boolean
    Dim taglist_templ As String
    Set form1 = Forms("Database Tools")

    output_table = form1.Text23.Value
    input_table = form1.Text71.Value
    taglist_templ = form1.Text49.Value

    strPath = CurrentProject.FullName

    If Not IsNull(DLookup("Type", "MSYSObjects", "Name='" & output_table & "'")) Then
        MsgBox "Query " & output_table & ": an object already exists with this name, deleting old one."
            DoCmd.DeleteObject acTable, output_table
    End If
    DoCmd.TransferDatabase acImport, "Microsoft Access", strPath, acTable, taglist_templ, output_table, True
 
    Set db = CurrentDb()
    Set rst = db.OpenRecordset(input_table, dbOpenDynaset) 'Ursprungstabelle
    Set rst2 = db.OpenRecordset(output_table, dbOpenDynaset) 'Zieltabelle
    Set exclude1 = db.OpenRecordset(form1.Text18.Value, dbOpenDynaset)
    Set exclude2instr = db.OpenRecordset(form1.Text21.Value, dbOpenDynaset)

    '''Zeiger an den Anfang der Tabellen setzen'''
    If Not (rst.BOF And rst.EOF) Then
        rst.MoveFirst
    End If
    If Not (rst2.BOF And rst2.EOF) Then
        rst2.MoveFirst
    End If
    If Not (exclude1.BOF And exclude1.EOF) Then
        exclude1.MoveFirst
    End If
    If Not (exclude2instr.BOF And exclude2instr.EOF) Then
        exclude2instr.MoveFirst
    End If

    '''Exclude-Arrays einlesen'''
    excl1 = exclude1.GetRows(1000)
    excl2Instr = exclude2instr.GetRows(1000)

    '''''''''''''''''start''''''''''''''''''''
    '''Taglisten aus Originaldatei einlesen'''
    ''''''''''''''''''''''''''''''''''''''''''
    rst.MoveLast
    rst.MoveFirst

    Do While Not rst.EOF
        If IsNull(rst("Tags").Value) Or rst("Tags").Value = ";;" Then
            GoTo skip
        End If
        s = rst("Tags").Value
        a = Split(s, ";")
        'a = Split(RegExprReplace(s, "[^a-zA-Z0-9;]", ""), ";")

        'Pro Zeile für jedes Tag:
        For i = 0 To UBound(a, 1)
            'Prüfen, ob Tag in Exclude-Listen
            found = False
    
     Set objRegEx = CreateObject("VBScript.RegExp")
        objRegEx.Global = True
            objRegEx.Pattern = "[^A-Za-z0-9]"
            a(i) = objRegEx.Replace(a(i), "")
            'RegExprReplace(A2,"[^a-zA-Z0-9,\.;: -]") --That formula will try to remove anything that is not a letter, digit, period, comma, semicolon, colon, space, or hyphen.  To expand to other punctuation marks:
            'a(i) = RegExprReplace(a(i)
            If Len(a(i)) <= 1 Or IsNumeric(Left(a(i), 1)) = True Then found = True
            If found = False Then
                For ii = 0 To UBound(excl2Instr, 2)
                    If InStr(1, a(i), excl2Instr(1, ii)) > 0 Then
                        found = True
                        Exit For
                    End If
                Next ii
            End If
            If found = False Then
                For ii = 0 To UBound(excl1, 2)
                    If a(i) = excl1(1, ii) Then
                        found = True
                        Exit For
                    End If
                Next ii
            End If

            'Einzelnse Tags aussortieren und in Tabelle 2 (rst2) untereinander schreiben,
            'falls nicht in einer der exclude-Tabellen
            If Not a(i) = "" And found = False Then
                rst2.AddNew
                rst2("Tags").Value = a(i)
                rst2.Update
            End If
        Next i
skip:
        rst.MoveNext 'Nächste TagListRow
    Loop

    '''Close all open tables'''
    rst.Close
    rst2.Close
    exclude1.Close
    exclude2instr.Close

    Set rst = Nothing
    Set rst2 = Nothing
    Set exclude1 = Nothing
    Set exclude2instr = Nothing

    form1.Text27.Value = output_table
    form1.Text33.Value = output_table & "_Count"

End Sub

Sub count_duplicates()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim sSQL As String
    Set form1 = Forms("Database Tools")
    Set db = CurrentDb

    tablename = form1.Text27.Value
    feldname = form1.Text51.Value
    output_name = form1.Text33.Value

    If Not IsNull(DLookup("Type", "MSYSObjects", "Name='" & output_name & "'")) Then
        MsgBox "Query " & output_name & ": an object already exists with this name, using this one instead."
    Else
        CurrentDb.CreateQueryDef output_name, "SELECT * FROM taglist_templ"
    End If

    DoCmd.SetWarnings False
    Set qdf = db.QueryDefs(output_name)

    sSQL = " SELECT First(" & tablename & ".[" & feldname & "]) AS [TAGNAME], Count(" & tablename & ".[" & feldname & "]) AS TAGCOUNT"
    sSQL = sSQL & " FROM " & tablename
    sSQL = sSQL & " GROUP BY " & tablename & ".[" & feldname & "]"
    sSQL = sSQL & " HAVING (((Count(" & tablename & ".[" & feldname & "]))>1))"
    sSQL = sSQL & " ORDER BY Count(" & tablename & ".[" & feldname & "]) DESC;"
    qdf.SQL = sSQL

    DoCmd.OpenQuery output_name
    DoCmd.SetWarnings True
          
    Set qdf = Nothing
    Set db = Nothing

    form1.Text54.Value = output_name
End Sub

Sub Month_statistics()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim sSQL As String
    Set form1 = Forms("Database Tools")
    Set db = CurrentDb

    tablename = form1.Text78.Value
    feldname = "DateTaken"
    feldname_Countdistinct = form1.Text90.Value
    output_name = tablename & "_perMonth"

    If Not IsNull(DLookup("Type", "MSYSObjects", "Name='" & output_name & "'")) Then
        MsgBox "Query " & output_name & ": an object already exists with this name, using this one instead."
    Else
        CurrentDb.CreateQueryDef output_name, "SELECT * FROM taglist_templ"
    End If

    DoCmd.SetWarnings False
    Set qdf = db.QueryDefs(output_name)
    sSQL = " SELECT MONTH, COUNT([" & feldname_Countdistinct & "]) AS " & feldname_Countdistinct & "_COUNT"
    sSQL = sSQL & " FROM ( SELECT DISTINCT Format(" & tablename & ".[" & feldname & "],'mm') AS MONTH, [" & feldname_Countdistinct & "] FROM " & tablename & ") AS TBL_tmp"
    sSQL = sSQL & " GROUP BY MONTH"
    qdf.SQL = sSQL

    DoCmd.OpenQuery output_name
    DoCmd.SetWarnings True
          
    Set qdf = Nothing
    Set db = Nothing
    DoCmd.OpenQuery output_name, acViewNormal, acReadOnly
End Sub

Sub Day_statistics()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim sSQL As String
    Set form1 = Forms("Database Tools")
    Set db = CurrentDb

    tablename = form1.Text78.Value
    feldname = "DateTaken"
    feldname_Countdistinct = form1.Text90.Value
    output_name = tablename & "_perDay"

    If Not IsNull(DLookup("Type", "MSYSObjects", "Name='" & output_name & "'")) Then
        MsgBox "Query " & output_name & ": an object already exists with this name, using this one instead."
    Else
        CurrentDb.CreateQueryDef output_name, "SELECT * FROM taglist_templ"
    End If

    DoCmd.SetWarnings False
    Set qdf = db.QueryDefs(output_name)
    sSQL = " SELECT DAY, COUNT([" & feldname_Countdistinct & "]) AS " & feldname_Countdistinct & "_COUNT"
    sSQL = sSQL & " FROM ( SELECT DISTINCT Format(" & tablename & ".[" & feldname & "],'dd') AS DAY, [" & feldname_Countdistinct & "] FROM " & tablename & ") AS TBL_tmp"
    sSQL = sSQL & " GROUP BY DAY"
    qdf.SQL = sSQL

    DoCmd.OpenQuery output_name
    DoCmd.SetWarnings True
          
    Set qdf = Nothing
    Set db = Nothing
    DoCmd.OpenQuery output_name, acViewNormal, acReadOnly
End Sub

Sub Weekday_statistics()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim sSQL As String
    Set form1 = Forms("Database Tools")
    Set db = CurrentDb

    tablename = form1.Text78.Value
    feldname = "DateTaken"
    feldname_Countdistinct = form1.Text90.Value
    output_name = tablename & "_perWeekday"

    If Not IsNull(DLookup("Type", "MSYSObjects", "Name='" & output_name & "'")) Then
        MsgBox "Query " & output_name & ": an object already exists with this name, using this one instead."
    Else
        CurrentDb.CreateQueryDef output_name, "SELECT * FROM taglist_templ"
    End If

    DoCmd.SetWarnings False
    Set qdf = db.QueryDefs(output_name)
    sSQL = " SELECT weekday, COUNT([" & feldname_Countdistinct & "]) AS " & feldname_Countdistinct & "_COUNT"
    sSQL = sSQL & " FROM ( SELECT DISTINCT WEEKDAY(" & tablename & ".[" & feldname & "]) AS weekday, [" & feldname_Countdistinct & "] FROM " & tablename & ") AS TBL_tmp"
    sSQL = sSQL & " GROUP BY weekday"
    qdf.SQL = sSQL

    DoCmd.OpenQuery output_name
    DoCmd.SetWarnings True
          
    Set qdf = Nothing
    Set db = Nothing
    DoCmd.OpenQuery output_name, acViewNormal, acReadOnly
End Sub

Sub Hour_statistics()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim sSQL As String
    Set form1 = Forms("Database Tools")
    Set db = CurrentDb

    tablename = form1.Text78.Value
    feldname = "DateTaken"
    feldname_Countdistinct = form1.Text90.Value
    output_name = tablename & "_perHour"

    If Not IsNull(DLookup("Type", "MSYSObjects", "Name='" & output_name & "'")) Then
        MsgBox "Query " & output_name & ": an object already exists with this name, using this one instead."
    Else
        CurrentDb.CreateQueryDef output_name, "SELECT * FROM taglist_templ"
    End If

    DoCmd.SetWarnings False
    Set qdf = db.QueryDefs(output_name)
    sSQL = " SELECT HOUR, COUNT([" & feldname_Countdistinct & "]) AS " & feldname_Countdistinct & "_COUNT"
    sSQL = sSQL & " FROM ( SELECT DISTINCT Format(" & tablename & ".[" & feldname & "],'hh') AS HOUR, [" & feldname_Countdistinct & "] FROM " & tablename & ") AS TBL_tmp"
    sSQL = sSQL & " GROUP BY HOUR"
    qdf.SQL = sSQL

    DoCmd.OpenQuery output_name
    DoCmd.SetWarnings True
          
    Set qdf = Nothing
    Set db = Nothing
    DoCmd.OpenQuery output_name, acViewNormal, acReadOnly
End Sub

Sub Year_statistics()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim sSQL As String
    Set form1 = Forms("Database Tools")
    Set db = CurrentDb

    tablename = form1.Text78.Value
    feldname = "DateTaken"
    feldname_Countdistinct = form1.Text90.Value
    output_name = tablename & "_perYear"

    If Not IsNull(DLookup("Type", "MSYSObjects", "Name='" & output_name & "'")) Then
        MsgBox "Query " & output_name & ": an object already exists with this name, using this one instead."
    Else
        CurrentDb.CreateQueryDef output_name, "SELECT * FROM taglist_templ"
    End If

    DoCmd.SetWarnings False
    Set qdf = db.QueryDefs(output_name)

    sSQL = " SELECT YEAR, COUNT([" & feldname_Countdistinct & "]) AS " & feldname_Countdistinct & "_COUNT"
    sSQL = sSQL & " FROM ( SELECT DISTINCT Format(" & tablename & ".[" & feldname & "],'yyyy') AS YEAR, [" & feldname_Countdistinct & "] FROM " & tablename & ") AS TBL_tmp"
    sSQL = sSQL & " GROUP BY YEAR"
    qdf.SQL = sSQL

    DoCmd.OpenQuery output_name
    DoCmd.SetWarnings True
          
    Set qdf = Nothing
    Set db = Nothing
    DoCmd.OpenQuery output_name, acViewNormal, acReadOnly
End Sub

Sub duplicate_table()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim sSQL As String
    Dim tablename As String
    Set form1 = Forms("Database Tools")
    Set db = CurrentDb

    tablename = form1.Text33.Value
    outputname = tablename & "_Sel"
    strPath = CurrentProject.FullName

    If Not IsNull(DLookup("Type", "MSYSObjects", "Name='" & outputname & "'")) Then
        MsgBox "Query " & outputname & ": an object already exists with this name, deleting old one."
        DoCmd.DeleteObject acTable, outputname
    End If
    DoCmd.TransferDatabase acImport, "Microsoft Access", strPath, acTable, tablename, outputname, False

    form1.Text54.Value = outputname

    tablename = outputname
    If tablename = "" Or tablename = "Enter Tablename" Then
        MsgBox("Please specify the tablename you would like to view.")
    Else
        If IsOpen(tablename, acTable) = False Then
            DoCmd.OpenTable tablename, acViewNormal, acEdit
            Else
            DoCmd.Close acTable, tablename, acSavePrompt
        End If
    End If
End Sub

Sub DayOfYear_statistics()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim sSQL As String
    Set form1 = Forms("Database Tools")
    Set db = CurrentDb

    tablename = form1.Text78.Value
    feldname = "DateTaken"
    feldname_Countdistinct = form1.Text90.Value
    output_name = tablename & "_perMONTHDAY"

    If Not IsNull(DLookup("Type", "MSYSObjects", "Name='" & output_name & "'")) Then
        MsgBox "Query " & output_name & ": an object already exists with this name, using this one instead."
    Else
        CurrentDb.CreateQueryDef output_name, "SELECT * FROM taglist_templ"
    End If

    DoCmd.SetWarnings False
    Set qdf = db.QueryDefs(output_name)
    sSQL = " SELECT MONTHDAY, COUNT([" & feldname_Countdistinct & "]) AS " & feldname_Countdistinct & "_COUNT"
    sSQL = sSQL & " FROM ( SELECT DISTINCT Format(" & tablename & ".[" & feldname & "],'mm/dd') AS MONTHDAY, [" & feldname_Countdistinct & "] FROM " & tablename & ") AS TBL_tmp"
    sSQL = sSQL & " GROUP BY MONTHDAY"
    qdf.SQL = sSQL

    DoCmd.OpenQuery output_name
    DoCmd.SetWarnings True
          
    Set qdf = Nothing
    Set db = Nothing
    DoCmd.OpenQuery output_name, acViewNormal, acReadOnly
End Sub

Sub UniqueDay_statistics()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim sSQL As String
    Set form1 = Forms("Database Tools")
    Set db = CurrentDb

    tablename = form1.Text78.Value
    feldname = "DateTaken"
    feldname_Countdistinct = form1.Text90.Value
    output_name = tablename & "_perUNIQUEDAY"

    If Not IsNull(DLookup("Type", "MSYSObjects", "Name='" & output_name & "'")) Then
        MsgBox "Query " & output_name & ": an object already exists with this name, using this one instead."
    Else
        CurrentDb.CreateQueryDef output_name, "SELECT * FROM taglist_templ"
    End If

    DoCmd.SetWarnings False
    Set qdf = db.QueryDefs(output_name)
    sSQL = " SELECT UNIQUEDAY, COUNT([" & feldname_Countdistinct & "]) AS " & feldname_Countdistinct & "_COUNT"
    sSQL = sSQL & " FROM ( SELECT DISTINCT Format(" & tablename & ".[" & feldname & "],'mm/dd/yyyy') AS UNIQUEDAY, [" & feldname_Countdistinct & "] FROM " & tablename & ") AS TBL_tmp"
    sSQL = sSQL & " GROUP BY UNIQUEDAY"
    qdf.SQL = sSQL

    DoCmd.OpenQuery output_name
    DoCmd.SetWarnings True
          
    Set qdf = Nothing
    Set db = Nothing
    DoCmd.OpenQuery output_name, acViewNormal, acReadOnly
End Sub


Sub export_taglist()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim sSQL As String
    Set form1 = Forms("Database Tools")
    Set db = CurrentDb

    path = Application.CurrentProject.path & form1.Text35.Value

    tablename = form1.Text54.Value
    path = form1.Text35.Value
    strPath = CurrentProject.path
    outputpath = strPath & path '& tablename '& ".dbf"
    DoCmd.OutputTo acOutputTable, tablename, acFormatXLS, strPath & path & tablename & ".xls", False
End Sub

Sub export_taglist_xlsx()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim sSQL As String
    Set form1 = Forms("Database Tools")
    Set db = CurrentDb

    path = Application.CurrentProject.path & form1.Text35.Value

    tablename = form1.Text54.Value
    path = form1.Text35.Value
    strPath = CurrentProject.path
    outputpath = strPath & path '& tablename '& ".dbf"
    DoCmd.OutputTo acOutputTable, tablename, acFormatXLSX, strPath & path & tablename & ".xlsx", False
End Sub

Sub export_taglist_dbf()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim sSQL As String
    Set form1 = Forms("Database Tools")
    Set db = CurrentDb

    path = Application.CurrentProject.path & form1.Text35.Value

    tablename = form1.Text54.Value
    path = form1.Text35.Value
    strPath = CurrentProject.path
    outputpath = strPath & path '& tablename '& ".dbf"
    DoCmd.TransferDatabase acExport, "dBase IV", outputpath, acTable, tablename, "Taglist", False
End Sub


Function CountCSWords(ByVal s) As Integer
    ' Counts the words in a string that are separated by commas.

    Dim WC As Integer, Pos As Integer
    If VarType(s) <> 8 Or Len(s) = 0 Then
        CountCSWords = 0
        Exit Function
    End If
    WC = 1
    Pos = InStr(s, ";")
    Do While Pos > 0
        WC = WC + 1
        Pos = InStr(Pos + 1, s, ";")
    Loop
    CountCSWords = WC
End Function

Function GetCSWord(ByVal s, Indx As Integer)
    ' Returns the nth word in a specific field.

    Dim WC As Integer, Count As Integer, SPos As Integer, EPos As Integer
    WC = CountCSWords(s)
    If Indx < 1 Or Indx > WC Then
        GetCSWord = Null
        Exit Function
    End If
    Count = 1
    SPos = 1
    For Count = 2 To Indx
        SPos = InStr(SPos, s, ";") + 1
    Next Count
    EPos = InStr(SPos, s, ";") - 1
    If EPos <= 0 Then EPos = Len(s)
    GetCSWord = Trim(Mid(s, SPos, EPos - SPos + 1))
End Function

myfile = Dir(path + "*.txt", vbHidden)


Sub import_txt()
    Dim myfile, tablename, specification, path As String
    Set form1 = Forms("Database Tools")

'MsgBox form1.Text6.Value
 
    path = Application.CurrentProject.path & form1.Text2.Value
    tablename = form1.Text6.Value
    specification = form1.Text40.Value

    myfile = Dir(path & "*.txt")
    Do While myfile <> "" 'will cause to loop through all txt files in path
        DoCmd.TransferText acImportDelim, specification, tablename, path + myfile, 0  'imports file
        myfile = Dir
    Loop


End Sub

Sub addprimary()
    Dim sSQL As String
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim ind As DAO.Index
    Dim strPath As String
    Set form1 = Forms("Database Tools")

    tablename = form1.Text6.Value
    tablename_new = tablename & "_UID"

    'Copy Structure of Table to new table
    DoCmd.SetWarnings False
    strPath = CurrentProject.FullName
    DoCmd.TransferDatabase acImport, "Microsoft Access", strPath, acTable, tablename, tablename_new, True

    'Add Primary Key with AutoNumber to new and empty table
    sSQL = "ALTER TABLE " & tablename_new & " ADD COLUMN prID COUNTER PRIMARY KEY"
    DoCmd.RunSQL sSQL

    'Copy Data from old table to new table, let access autonumber keyfield
    sSQL = "INSERT INTO " & tablename_new & " SELECT * FROM " & tablename & " ;"
    DoCmd.RunSQL sSQL
    DoCmd.SetWarnings True

    form1.Text11.Value = tablename & "_UID"
    form1.Text43.Value = "PhotoID"
End Sub

Sub distinct()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim sSQL As String
    Set form1 = Forms("Database Tools")
    Set db = CurrentDb

    tablename_new = form1.Text11.Value
    distfield = form1.Text43.Value

    If Not IsNull(DLookup("Type", "MSYSObjects", "Name='distQuery'")) Then
        MsgBox "Query distQuery: an object already exists with this name, using this one instead."
    Else
        CurrentDb.CreateQueryDef "distQuery", "SELECT * FROM " & tablename_new
    End If

    If Not IsNull(DLookup("Type", "MSYSObjects", "Name='distTotals'")) Then
        MsgBox "Query distTotals: an object already exists with this name, using this one instead."
    Else
        CurrentDb.CreateQueryDef "distTotals", "SELECT * FROM " & tablename_new
    End If
    DoCmd.SetWarnings False
     Set qdf = db.QueryDefs("distQuery")
     Set qdf2 = db.QueryDefs("distTotals")
      
     sSQL = " SELECT " & tablename_new & "." & distfield & ", First(" & tablename_new & ".prID) AS ErsterWertvonprID"
    sSQL = sSQL & " FROM " & tablename_new
    sSQL = sSQL & " GROUP BY " & tablename_new & "." & distfield & ";"
    qdf.SQL = sSQL
    'DoCmd.OpenQuery "distQuery"

    sSQL = "SELECT " & tablename_new & ".* " & " INTO " & tablename_new & "_Dist"
    sSQL = sSQL & " FROM " & tablename_new & " INNER JOIN distQuery AS [Distinct] ON " & tablename_new & ".prID = Distinct.ErsterWertvonprID"

    sSQL = sSQL & " ORDER BY DateTaken DESC;"
    ' qdf2.SQL = sSQL
    'DoCmd.OpenQuery "distTotals"
    DoCmd.RunSQL sSQL
     DoCmd.SetWarnings True
     
     Set qdf = Nothing
     Set db = Nothing
     DoCmd.DeleteObject acQuery, "distQuery"
     DoCmd.DeleteObject acQuery, "distTotals"

     'DoCmd.DeleteObject acQuery, "distTotals"
    form1.Text14.Value = tablename_new & "_Dist"
    form1.Text23.Value = "Taglist_" & tablename_new & "_Dist"
    form1.Text64.Value = tablename_new & "_Dist"

End Sub

Public Sub ExportAsMDB()
    Dim strTargetDB As String
    Dim qdf As QueryDef
    Dim tbl As TableDef
    Dim strNewDB As String
    Set form1 = Forms("Database Tools")
 

    strPath = form1.Text62.Value
    strTable = form1.Text64.Value
    strNewDB = "Export_" & strTable

    strTargetDB = CurrentProject.path & strPath & strNewDB & ".mdb"
    'DBEngine.CreateDatabase strTargetDB, dbLangGeneral

    DBEngine.CreateDatabase strTargetDB, dbLangGeneral, dbVersion40

 DoCmd.TransferDatabase acExport, "Microsoft Access", strTargetDB, acTable, strTable, strTable

 End Sub

Public Sub viewtable()
    Dim tablename As String
    Set form1 = Forms("Database Tools")

    tablename = form1.Text78.Value
    If tablename = "" Or tablename = "Enter Tablename" Then
        MsgBox("Please specify the tablename you would like to view.")
    Else
        If IsOpen(tablename, acTable) = False Then
            DoCmd.OpenTable tablename, acViewNormal, acEdit
            Else
            DoCmd.Close acTable, tablename, acSavePrompt
        End If
    End If
End Sub

Function IsOpen(strname As String, strtype As String) As Boolean

    If SysCmd(acSysCmdGetObjectState, strtype, strname) <> 0 Then

        IsOpen = True

    End If
End Function

Public Sub deletetable()
    Dim tablename As String
    Set form1 = Forms("Database Tools")

    tablename = form1.Text78.Value
    If tablename = "" Or tablename = "Enter Tablename" Then
        MsgBox("Please specify the tablename you would like to delete.")
    Else
        DoCmd.DeleteObject acTable, tablename
    End If
End Sub

Public Sub StopListGlobal()
    Dim tablename As String
    Set form1 = Forms("Database Tools")

    tablename = "SortOutAlways"
    If IsOpen(tablename, acTable) = False Then
        DoCmd.OpenTable tablename, acViewNormal, acEdit
            Else
        DoCmd.Close acTable, tablename, acSavePrompt
        End If
End Sub

Public Sub StopListInStr()
    Dim tablename As String
    Set form1 = Forms("Database Tools")

    tablename = "SortOutAlways_InStr"
    If IsOpen(tablename, acTable) = False Then
        DoCmd.OpenTable tablename, acViewNormal, acEdit
            Else
        DoCmd.Close acTable, tablename, acSavePrompt
        End If
End Sub

Function RegExprReplace(LookIn As String, PatternStr As String, Optional ReplaceWith As String = "", Optional ReplaceAll As Boolean = True, Optional MatchCase As Boolean = True, Optional MultiLine As Boolean = False)

    ' Function written by Patrick G. Matthews.  You may use and distribute this code freely,
    ' as long as you properly credit and attribute authorship and the URL of where you
    ' found the code

    ' For more info, please see:
    ' http://www.experts-exchange.com/articles/Programming/Languages/Visual_Basic/Using-Regular-Expressions-in-Visual-Basic-for-Applications-and-Visual-Basic-6.html

    ' This function relies on the VBScript version of Regular Expressions, and thus some of
    ' the functionality available in Perl and/or .Net may not be available.  The full extent
    ' of what functionality will be available on any given computer is based on which version
    ' of the VBScript runtime is installed on that computer

    ' This function uses Regular Expressions to parse a string, and replace parts of the string
    ' matching the specified pattern with another string.  The optional argument ReplaceAll
    ' controls whether all instances of the matched string are replaced (True) or just the first
    ' instance (False)

    ' If you need to replace the Nth match, or a range of matches, then use RegExpReplaceRange
    ' instead

    ' By default, RegExp is case-sensitive in pattern-matching.  To keep this, omit MatchCase or
    ' set it to True

    ' If you use this function from Excel, you may substitute range references for all the arguments

    ' Normally as an object variable I would set the RegX variable to Nothing; however, in cases
    ' where a large number of calls to this function are made, making RegX a static variable that
    ' preserves its state in between calls significantly improves performance

    Static RegX As Object

    If RegX Is Nothing Then Set RegX = CreateObject("VBScript.RegExp")
    With RegX
        .Pattern = PatternStr
        .Global = ReplaceAll
        .IgnoreCase = Not MatchCase
        .MultiLine = MultiLine
    End With

    RegExpReplace = RegX.Replace(LookIn, ReplaceWith)

End Function



