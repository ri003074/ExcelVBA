Attribute VB_Name = "Module1"
Option Explicit

Sub make_graph()
    Dim graph As graph
    Set graph = New graph
    
    Dim input_folder As String
    Dim fso As Object
    Dim file As Object
        
    input_folder = "C:\Users\ri003\Documents\Programming\ExcelVBA\data"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    For Each file In fso.GetFolder(input_folder).Files
        If LCase(file.Name) Like "*.csv" Then
            
            Dim file_path As String
            file_path = file
               
            graph.save_as_csv (file_path)
            graph.open_file
            graph.add_chart
            graph.relocate_graph ("E2")
            graph.resize_graph 300, 400
            graph.set_tick xlValue, 0, 120, 20
            graph.set_axis_title xlValue, "ps", 20
            graph.set_axis_title xlCategory, "", 20
            graph.save_png
            ActiveWorkbook.Close
        End If
    Next
End Sub

