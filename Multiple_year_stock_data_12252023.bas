Attribute VB_Name = "Module1"
Sub ticker()

    Dim ws As Worksheet
    
    Dim i As Long
    Dim j As Long
    Dim ticker_Name As String
    Dim Opening_P As Double
    Dim Closing_P As Double
    Dim Y_Change As Double
    Dim P_Change As Double
    Dim Total_Volume As Double
    Dim G_Increase As Double
    Dim G_Decrease As Double
    Dim G_Volume As Double
    Dim LastRow1 As Long
    Dim LastRow2 As Long

    For Each ws In ThisWorkbook.Sheets

    ws.Activate
    MsgBox (ws.Name)
        LastRow1 = Cells(Rows.Count, 1).End(xlUp).Row
        LastRow2 = Cells(Rows.Count, 10).End(xlUp).Row
    
        MsgBox (LastRow1)
    
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Opening Price"
        Cells(1, 11).Value = "Closing Price"
        Cells(1, 12).Value = "Yearly Change"
        Cells(1, 13).Value = "Percent Change"
        Cells(1, 14).Value = "Total Stock Volume"
    
        Cells(2, 16).Value = "Greatest % Increase"
        Cells(3, 16).Value = "Greatest % Decrease"
        Cells(4, 16).Value = "Greatest Total Volume"
    
        Cells(1, 17).Value = "Ticker"
        Cells(1, 18).Value = "Value"
    
        ' Initialize variables
        Opening_P = Cells(2, 3).Value
        Total_Volume = 0
        G_Increase = 0
        G_Decrease = 0
        G_Volume = 0
        ticker_Name = Cells(2, 1).Value
        j = 2
    
        For i = 2 To LastRow1
            If Cells(i, 1).Value <> ticker_Name Then
                ' New ticker starts
                Closing_P = Cells(i - 1, 6).Value
                Y_Change = Closing_P - Opening_P
                If Opening_P <> 0 Then
                    P_Change = (Y_Change / Opening_P)
                Else
                    P_Change = 0
                End If
                
                Cells(j, 9).Value = ticker_Name
                Cells(j, 10).Value = Opening_P
                Cells(j, 11).Value = Closing_P
                Cells(j, 12).Value = Y_Change
                Cells(j, 13).Value = P_Change
                Cells(j, 13).NumberFormat = "0.00%"
                Cells(j, 14).Value = Total_Volume
                
                'Set Yearly Change to red if negative and green if positive
                If Cells(j, 12).Value < 0 Then
                Cells(j, 12).Interior.Color = RGB(255, 0, 0)
                Else
                Cells(j, 12).Interior.Color = RGB(0, 255, 0)
                End If
                
    
                ' Update greatest percent increase, decrease, and volume logic here
                If P_Change > G_Increase Then
                    G_Increase = P_Change
                    Cells(2, 17).Value = ticker_Name
                    Cells(2, 18).Value = G_Increase
                    Cells(2, 18).NumberFormat = "0.00%"
                    
                End If
                
                If P_Change < G_Decrease Then
                    G_Decrease = P_Change
                    Cells(3, 17).Value = ticker_Name
                    Cells(3, 18).Value = G_Decrease
                    Cells(3, 18).NumberFormat = "0.00%"
                End If
                
                If Total_Volume > G_Volume Then
                    G_Volume = Total_Volume
                    Cells(4, 17).Value = ticker_Name
                    Cells(4, 18).Value = G_Volume
                End If
    
                ' Reset values for the next ticker
                Opening_P = Cells(i, 3).Value
                Total_Volume = 0
                j = j + 1
                ticker_Name = Cells(i, 1).Value
            Else
                ' Same ticker continues, accumulate volume
                Total_Volume = Total_Volume + Cells(i, 7).Value
            End If
        Next i
    Next ws
    
End Sub
