Attribute VB_Name = "Module1"
Sub GetGrades()
    Dim ie As Object
    Dim html As Object
    Dim loginURL As String
    Dim username As String
    Dim password As String
    Dim formElem As Object
    Dim centerElem As Object
    Dim h4Elem As Object
    Dim bElem As Object
    Dim tableElem As Object
    Dim theadElem As Object
    Dim tbodyElem As Object
    Dim headerRow As Object
    Dim headerCell As Object
    Dim dataRow As Object
    Dim dataCell As Object
    Dim ws As Worksheet
    Dim i As Integer
    Dim j As Integer

    ' Set login URL and credentials
    loginURL = "https://www.iitm.ac.in/viewgrades/"
    username = Sheets("Dashboard").Range("D2").Value
    password = Sheets("Dashboard").Range("D3").Value

    ' Create Internet Explorer Object
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = False

    ' Navigate to the login page
    ie.Navigate loginURL

    ' Wait for the page to load
    Do While ie.Busy Or ie.ReadyState <> 4
        DoEvents
    Loop

    ' Enter login credentials
    Set html = ie.Document
    html.getElementById("username").Value = username
    html.getElementById("password").Value = password

    ' Click the submit button
    html.getElementById("submit").Click

    ' Wait for the next page to load
    Do While ie.Busy Or ie.ReadyState <> 4
        DoEvents
    Loop

    ' Retrieve the form with ID "slotwise"
    Set formElem = html.getElementById("slotwise")

    ' Retrieve student name and roll number
    Set centerElem = formElem.getElementsByTagName("center")(0)
    If Not centerElem Is Nothing Then
        Set h4Elem = centerElem.getElementsByTagName("h4")(0)
        If Not h4Elem Is Nothing Then
            Set bElem = h4Elem.getElementsByTagName("b")(0)
            If Not bElem Is Nothing Then
                studentInfo = bElem.innerText
            End If
        End If
    End If

    ' Split student information into name and roll number
    studentName = ""
    rollNumber = ""
    If Len(studentInfo) > 0 Then
        infoArray = Split(studentInfo, " ")
        If UBound(infoArray) >= 1 Then rollNumber = Trim(infoArray(3))
        If UBound(infoArray) >= 0 Then
            i = 5
            While infoArray(i) <> "B.Tech"
                studentName = studentName + " " + Trim(infoArray(i))
                i = i + 1
            Wend
        End If
    End If

    ' Clear previous grades data in the sheet
    Set ws = Sheets("Data")
    ws.Cells.Clear

    ' Write student name and roll number to the sheet
    ws.Cells(1, 1).Value = "Student Name"
    ws.Cells(1, 2).Value = WorksheetFunction.Proper(studentName)
    ws.Cells(2, 1).Value = "Roll Number"
    ws.Cells(2, 2).Value = WorksheetFunction.Proper(rollNumber)

    ' Initialize row for Excel sheet
    i = 4

    ' Loop through each table within the form
    For Each tableElem In formElem.getElementsByTagName("table")
        ' Process table header
        Set theadElem = tableElem.getElementsByTagName("thead")(0)
        If Not theadElem Is Nothing Then
            Set headerRow = theadElem.getElementsByTagName("tr")(0)
            j = 1
            For Each headerCell In headerRow.getElementsByTagName("th")
                ws.Cells(i, j).Value = headerCell.getElementsByTagName("font")(0).innerText
                j = j + 1
            Next headerCell
            i = i + 1
        End If

        ' Process table body
        Dim semester As Integer
        
        Set tbodyElem = tableElem.getElementsByTagName("tbody")(0)
        If Not tbodyElem Is Nothing Then
            For Each dataRow In tbodyElem.getElementsByTagName("tr")
                j = 1
                For Each dataCell In dataRow.getElementsByTagName("td")
                    If dataCell.getElementsByTagName("font").Length > 0 Then
                        If dataCell.getElementsByTagName("font")(0).getElementsByTagName("b").Length > 0 Then
                            ' Process semester details
                            ws.Cells(i, j).Value = dataCell.getElementsByTagName("font")(0).getElementsByTagName("b")(0).innerText
                        
                            ' Track semester
                            Select Case Split(ws.Cells(i, j).Value, " ")(0)
                                Case "First": semester = 1
                                Case "Second": semester = 2
                                Case "Third": semester = 3
                                Case "Fourth": semester = 4
                                Case "Fifth": semester = 5
                                Case "Sixth": semester = 6
                                Case "Seventh": semester = 7
                                Case "Eighth": semester = 8
                                Case Else: semester = 0
                            End Select
                        Else
                            ' Process grade details
                            ws.Cells(i, j).Value = dataCell.innerText
                        End If
                    Else
                        ' Process grade details without font tag
                        ws.Cells(i, j).Value = dataCell.innerText
                        
                        ' Process earned credit, GPA and CGPA
                        If InStr(ws.Cells(i, j).Value, "Earned Credit") > 0 Then
                            Dim dataArray(1 To 3) As Variant
                            x = 1
                            infoArray = Split(ws.Cells(i, j).Value, " ")
                            For k = LBound(infoArray) To UBound(infoArray)
                                If InStr(infoArray(k), ":") > 0 Then
                                    partArray = Split(infoArray(k), ":")
                                    If UBound(partArray) >= 1 Then
                                        dataArray(x) = Trim(partArray(1))
                                        x = x + 1
                                    End If
                                End If
                            Next k
                            ws.Cells(i, 1) = "Earned Credit"
                            ws.Cells(i, 2) = dataArray(1)
                            ws.Cells(i, 3) = "GPA"
                            ws.Cells(i, 4) = dataArray(2)
                            ws.Cells(i, 5) = "CGPA"
                            ws.Cells(i, 6) = dataArray(3)
                            j = 6
                        End If
                        
                    End If
                    j = j + 1
                Next dataCell
                ws.Cells(i, 8).Value = semester
                i = i + 1
            Next dataRow
        End If

        ' Add a blank row between tables
        i = i + 1
    Next tableElem

    ' Clean up
    ie.Quit
    Set ie = Nothing
End Sub

