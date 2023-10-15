Attribute VB_Name = "TestGraphClass"
Option Explicit


Sub test_graph1()
    Dim Graph As Graph
    Set Graph = New Graph
    
    Dim input_folder As String: input_folder = "C:\Users\ri003\Documents\Programming\ExcelVBA\test_data"
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file As Object
    
    For Each file In fso.GetFolder(input_folder).Files
        If LCase(file.Name) Like "*.csv" Then
            Dim file_path As String: file_path = file
               
            Graph.save_as_csv (file_path)
            Graph.open_file
            Graph.add_chart xlLineMarkers
            Graph.relocate_graph ("E2")
            Graph.resize_graph 300, 400
            Graph.set_tick xlValue, 0, 120, 20
            Graph.set_axis_title xlValue, "ps", 20
            Graph.set_axis_title xlCategory, "", 20
            Graph.save_png
            Graph.delete_chart
            ActiveWorkbook.Save
            ActiveWorkbook.Close
        End If
    Next
End Sub


