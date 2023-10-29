Attribute VB_Name = "TestGraphClass"
Option Explicit


Sub test_graph1()
    Dim graph As graph
    Set graph = New graph
    
    Dim input_folder As String: input_folder = "C:\Users\ri003\Documents\Programming\ExcelVBA\test_data"
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file As Object
    
    For Each file In fso.GetFolder(input_folder).Files
        If LCase(file.Name) Like "*.csv" Then
            Dim file_path As String: file_path = file
               
            graph.save_as_xlsx (file_path)
            graph.open_file
            graph.add_chart xlLineMarkers
            graph.relocate_graph ("E2")
            graph.resize_graph 300, 400
            graph.set_tick xlValue, 0, 120, 20
            graph.set_axis_title xlValue, "ps", 20
            graph.set_axis_title xlCategory, "", 20
            graph.save_png
            graph.delete_chart
            ActiveWorkbook.Save
            ActiveWorkbook.Close
        End If
    Next
End Sub

