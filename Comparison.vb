Imports System
Imports System.Collections.Specialized
Public Class Comparison
    Private curr, prev As String
    Private same As Boolean
    Private diffs As DataTable
    Sub New(ByVal today As String, ByVal yesterday As String)
        curr = today
        prev = yesterday
        If prev.CompareTo(curr) = 0 Then
            same = True
        Else
            same = False
        End If
        SetDiffsTable()
    End Sub
    Function Today() As String
        Today = curr
    End Function
    Function NoChange() As Boolean
        NoChange = same
    End Function
    Function ShortChangesReport() As String
        Dim rpt As String = ""
        Dim changes As DataRow() = diffs.Select("Change <> ''", "Table")
        For Each change As DataRow In changes
            Dim chng As String = change("Change")
            Select Case chng
                Case "New count"
                    rpt += change("Table") & ": " & chng & " (" & change("NewRows") & ")."
                    rpt += ControlChars.NewLine
                Case "New"
                    rpt += change("Table") & ": " & chng & " (" & change("Rows") & ")."
                    rpt += ControlChars.NewLine
                Case "Removed"
                    rpt += change("Table") & ": " & chng & "." & ControlChars.NewLine
            End Select
        Next
        ShortChangesReport = rpt
    End Function
    Function LongChangesReport() As String
        Dim c1w As Integer = 23
        Dim c2w As Integer = 11
        Dim c3w As Integer = 10
        Dim rpt As String = ""
        rpt += "NOW".PadRight(c1w + c2w + 1) & " " & "BEFORE".PadRight(c1w + c2w + 1) & ControlChars.NewLine
        rpt += "".PadRight(c1w + c2w + 1, "-") & " " & "".PadRight(c1w + c2w + 1, "-") & ControlChars.NewLine
        rpt += "Table".PadRight(c1w) & " " & "Rows".PadRight(c2w) & " "
        rpt += "Table".PadRight(c1w) & " " & "Rows".PadRight(c2w) & " "
        rpt += "Difference".PadRight(c3w) & ControlChars.NewLine
        rpt += "".PadRight(c1w, "-") & " " & "".PadRight(c2w, "-") & " "
        rpt += "".PadRight(c1w, "-") & " " & "".PadRight(c2w, "-") & " "
        rpt += "".PadRight(c3w, "-") & ControlChars.NewLine
        Dim rs As DataRow() = diffs.Select("", "Table")
        For Each row As DataRow In rs
            Dim table As String = row("Table")
            Dim rows As String = row("Rows")
            Dim newrows As String = row("NewRows")
            Dim change As String = row("Change")
            Select Case change
                Case ""
                    rpt += table.PadRight(c1w) & " " & rows.PadRight(c2w) & " "
                    rpt += table.PadRight(c1w) & " " & rows.PadRight(c2w) & ControlChars.NewLine
                Case "New count"
                    rpt += table.PadRight(c1w) & " " & newrows.PadRight(c2w) & " "
                    rpt += table.PadRight(c1w) & " " & rows.PadRight(c2w) & " "
                    rpt += change.PadRight(c3w) & ControlChars.NewLine
                Case "Removed"
                    rpt += "".PadRight(c1w) & " " & "".PadRight(c2w) & " "
                    rpt += table.PadRight(c1w) & " " & rows.PadRight(c2w) & " "
                    rpt += change.PadRight(c3w) & ControlChars.NewLine
                Case "New"
                    rpt += table.PadRight(c1w) & " " & rows.PadRight(c2w) & " "
                    rpt += "".PadRight(c1w) & " " & "".PadRight(c2w) & " "
                    rpt += change.PadRight(c3w) & ControlChars.NewLine
            End Select
        Next
        LongChangesReport = rpt
    End Function

    Function ChangedTables() As StringCollection
        Dim t As New StringCollection
        For Each change As DataRow In diffs.Select("Change = 'New' or Change = 'New count'")
            t.Add(change("Table"))
        Next
        Return t
    End Function

    Private Sub SetDiffsTable()
        Dim prevTables As DataTable = SetTable(prev)
        Dim currTables As DataTable = SetTable(curr)
        diffs = New DataTable
        diffs.Columns.Add("Table", GetType(System.String))
        diffs.Columns.Add("Rows", GetType(System.Int32))
        diffs.Columns.Add("Change", GetType(System.String))
        Dim nrs As New DataColumn("NewRows", GetType(System.Int32))
        nrs.DefaultValue = 0
        diffs.Columns.Add(nrs)
        For i As Integer = 0 To prevTables.Rows.Count - 1
            Dim atable As String = prevTables.Rows(i).Item("Table").ToString
            Dim rows As Integer = prevTables.Rows(i).Item("Rows")
            Dim exp As String = "Table = " & "'" & atable & "'"
            Dim foundRows As DataRow() = currTables.Select(exp)
            If foundRows.Length <= 0 Then
                Dim aRow As DataRow = diffs.NewRow
                aRow("Table") = atable.Replace(vbLf, "")
                aRow("Rows") = rows
                aRow("Change") = "Removed"
                diffs.Rows.Add(aRow)
            ElseIf foundRows(0).Item("Rows") <> rows Then
                Dim aRow As DataRow = diffs.NewRow
                aRow("Table") = atable.Replace(vbLf, "")
                aRow("Rows") = rows
                aRow("NewRows") = foundRows(0).Item("Rows")
                aRow("Change") = "New count"
                diffs.Rows.Add(aRow)
            Else
                Dim aRow As DataRow = diffs.NewRow
                aRow("Table") = atable.Replace(vbLf, "")
                aRow("Rows") = rows
                aRow("Change") = ""
                diffs.Rows.Add(aRow)
            End If
        Next
        For j As Integer = 0 To currTables.Rows.Count - 1
            Dim atable = currTables.Rows(j).Item("Table").ToString
            Dim exp As String = "Table = " & "'" & atable & "'"
            Dim foundRows As DataRow() = prevTables.Select(exp)
            If foundRows.Length <= 0 Then
                Dim trows As Integer = currTables.Rows(j).Item("Rows")
                Dim aRow As DataRow = diffs.NewRow
                aRow("Table") = atable.Replace(vbLf, "")
                aRow("Rows") = trows
                aRow("Change") = "New"
                diffs.Rows.Add(aRow)
            End If
        Next
    End Sub
    Private Function SetTable(ByVal TableContents As String) As DataTable
        SetTable = New DataTable
        SetTable.Columns.Add("Table", GetType(System.String))
        SetTable.Columns.Add("Rows", GetType(System.Int32))
        For Each s As String In TableContents.Split(vbCrLf)
            Dim tmpFields As String() = s.Split(",")
            If tmpFields.Length > 1 Then
                Dim aRow As DataRow = SetTable.NewRow
                aRow("Table") = tmpFields(0).Replace(vbLf, "")
                aRow("Rows") = tmpFields(1)
                SetTable.Rows.Add(aRow)
            End If
        Next s
    End Function
End Class