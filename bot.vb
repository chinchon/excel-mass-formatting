Sub bot()

    Dim all, filename, i, xlApp, xlBook
    
    Application.DisplayAlerts = False
    
    With Application.FileDialog(msoFileDialogOpen)
        .InitialFileName = "\\DATA1\FORShare_DATA1\2014.11.19_Panasonic_DA\05_PIDMY\04_Test_Resultset\export\"
        .AllowMultiSelect = True
        .Show
        all = .SelectedItems.Count
    End With
        
    If all <> 0 Then
        For i = 1 To all
        Set xlApp = CreateObject("Excel.Application")
        xlApp.Visible = True
        filename = Application.FileDialog(msoFileDialogOpen).SelectedItems(i)
    
        Set xlBook = xlApp.Workbooks.Add()
        With xlBook.Sheets(1).QueryTables.Add(Connection:="TEXT;" + filename, Destination:=xlBook.Sheets(1).Range("A1"))
            .TextFileColumnDataTypes = Array(2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2)
            .Refresh
        End With
        xlApp.Run "\\DATA1\FORShare_DATA1\2014.11.19_Panasonic_DA\00_All\07_Analysis_Scripts\TestResults\format.xlsm!format"
        xlBook.SaveAs Replace(filename, ".txt", ".xlsx"), 51
    
        xlApp.Quit
        Next i
    End If
    
    Application.DisplayAlerts = True
    
End Sub
