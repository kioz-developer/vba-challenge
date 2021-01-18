Attribute VB_Name = "analyst_module"
Sub init()
    
    Call print_labels(ActiveSheet)
    Call process_sheet(ActiveSheet)
    Call apply_column_format(ActiveSheet)
    
    MsgBox ("Sheet " + ActiveSheet.Name + " have been finished.")
    
End Sub

Sub process_sheet(sht As Worksheet)
    Dim last_row As Long
    Dim counter As Integer
    
    Dim current_ticker As String
    Dim opening_prince As String
    Dim closing_prince As String
    Dim yearly_change As String
    Dim percent_change As String
    Dim total_volume As String
    
    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_volume As Double
    Dim increase_ticker As String
    Dim decrease_ticker As String
    Dim volume_ticker As String
    
    last_row = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row + 1
    greatest_increase = 0
    greatest_decrease = 0
    greatest_volume = 0
    
    opening_prince = sht.Range("C2")
    current_ticker = sht.Range("A2")
    total_volume = 0
    counter = 1
    For i = 2 To last_row
        If (current_ticker <> sht.Cells(i, 1).Value) Then
            counter = counter + 1
            
            ' A. Calculating column values
            current_ticker = sht.Cells(i - 1, 1).Value
            closing_prince = sht.Cells(i - 1, 6).Value
            yearly_change = closing_prince - opening_prince
            percent_change = 0
            If opening_prince > 0 Then
                percent_change = (yearly_change * 100) / opening_prince
            End If
            
            ' B. Assigning column values
            sht.Range("I" & counter).Value = current_ticker
            sht.Range("J" & counter).Value = yearly_change
            sht.Range("K" & counter).Value = percent_change / 100
            sht.Range("L" & counter).Value = total_volume
            
            ' C. Applying cell format
            Call apply_cell_format(sht, counter)
            
            ' D. Calculating bonus step
            If greatest_increase < percent_change Then
                greatest_increase = percent_change
                increase_ticker = current_ticker
            End If
            
            If greatest_decrease > percent_change Then
                greatest_decrease = percent_change
                decrease_ticker = current_ticker
            End If
            
            If greatest_volume < total_volume Then
                greatest_volume = total_volume
                volume_ticker = current_ticker
            End If
            
            opening_prince = sht.Cells(i, 3).Value
            total_volume = sht.Cells(i, 7).Value
            current_ticker = sht.Cells(i, 1).Value
        Else
            total_volume = total_volume + sht.Cells(i, 7).Value
        End If
    Next i
    
    ' E. Assigning bonus values
    sht.Range("P2").Value = increase_ticker
    sht.Range("P3").Value = decrease_ticker
    sht.Range("P4").Value = volume_ticker
    
    sht.Range("Q2").Value = greatest_increase / 100
    sht.Range("Q3").Value = greatest_decrease / 100
    sht.Range("Q4").Value = greatest_volume
End Sub

Sub print_labels(sht As Worksheet)
    sht.Range("I1").Value = "Ticker"
    sht.Range("J1").Value = "Yearly change"
    sht.Range("K1").Value = "Percent change"
    sht.Range("L1").Value = "Total stock"
    
    sht.Range("P1").Value = "Ticker"
    sht.Range("Q1").Value = "Value"
    
    sht.Range("O2").Value = "Greatest % increase"
    sht.Range("O3").Value = "Greatest % decrease"
    sht.Range("O4").Value = "Greatest total volume"
End Sub

Sub apply_cell_format(sht As Worksheet, counter As Integer)
    Dim green As String
    Dim red As String
    
    green = 4
    red = 3
    
    If sht.Cells(counter, 10).Value >= 0 Then
        sht.Range("J" & counter).Interior.ColorIndex = green
    Else
        sht.Range("J" & counter).Interior.ColorIndex = red
    End If
End Sub

Sub apply_column_format(sht As Worksheet)
    Dim last_row As Long
    
    last_row = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row + 1
    
    sht.Range("K2:K" & last_row).NumberFormat = "0.00%"
    sht.Range("Q2").NumberFormat = "0.00%"
    sht.Range("Q3").NumberFormat = "0.00%"
End Sub
