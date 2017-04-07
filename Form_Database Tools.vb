Option Compare Database

Private Sub Befehl1_Click()
    runsplit
End Sub

Private Sub Befehl13_Click()
    distinct
End Sub

Private Sub Befehl26_Click()
    count_duplicates
End Sub

Private Sub Befehl45_Click()
    Dim db As Database
    Dim i As Integer
    Dim deleted As Integer
    deleted = 0
    Set db = DBEngine(0)(0)
    For i = 0 To db.TableDefs.Count - 1
        If Right(db.TableDefs(i).Name, 12) = "Importfehler" Or Right(db.TableDefs(i).Name, 11) = "ImportError" Then
            db.TableDefs.Delete db.TableDefs(i).Name
    deleted = deleted + 1
            If i >= db.TableDefs.Count - 1 - deleted Then Exit For
        End If
    Next i
End Sub

Private Sub Befehl46_Click()
    addprimary
End Sub

Private Sub Befehl53_Click()
    export_taglist
End Sub

Private Sub Befehl84_Click()
    export_taglist_xlsx
End Sub

Private Sub Befehl82_Click()
    export_taglist_dbf
End Sub

Private Sub Befehl58_Click()
    duplicate_table
End Sub

Private Sub Befehl61_Click()
    ExportAsMDB
End Sub

Private Sub Befehl68_Click()
    saubern
End Sub

Private Sub Befehl88_Click()
    StopListGlobal
End Sub

Private Sub Befehl89_Click()
    StopListInStr
End Sub

Private Sub Befehl9_Click()
    import_txt
End Sub

Private Sub Befehl90_Click()
    Day_statistics
End Sub


Private Sub Befehl80_Click()
    viewtable
End Sub

Private Sub Befehl81_Click()
    deletetable
End Sub

Private Sub Befehl86_Click()
    Month_statistics
End Sub

Private Sub Befehl87_Click()
    Hour_statistics
End Sub

Private Sub Befehl91_Click()
    Weekday_statistics
End Sub

Private Sub Befehl93_Click()
    Year_statistics
End Sub

Private Sub Befehl94_Click()
    DayOfYear_statistics
End Sub

Private Sub Befehl95_Click()
    UniqueDay_statistics
End Sub

