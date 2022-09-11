Attribute VB_Name = "delete"
Sub delete()
    Application.DisplayAlerts = False
        Sheets("A").delete
        Sheets("B").delete
        Sheets("C").delete
        Sheets("D").delete
        Sheets("E").delete
        Sheets("F").delete
        Sheets("G").delete
        Sheets("IB").delete
     Application.DisplayAlerts = True
End Sub
