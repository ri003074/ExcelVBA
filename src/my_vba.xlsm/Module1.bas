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


Sub pic_to_pptx()
    Dim input_folder As String
    Dim fso As Object
    Dim file As Object
    Dim pptx As PowerPo
    Dim layout_number As Integer
    
    Set pptx = New PowerPo
    layout_number = 16
    
    pptx.activate_powerpoint
'    pptx.setup_new_powerpoint
    pptx.delete_all_slides
   
    input_folder = "C:\Users\ri003\Documents\Programming\ExcelVBA\data"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    For Each file In fso.GetFolder(input_folder).Files
        If LCase(file.Name) Like "*.png" Then
            Dim file_path As String
            file_path = file
            pptx.add_slide layout_number
            pptx.add_picture file_path
        End If
    Next
     
    'pptx.add_slide 16
'    pptx.delete_all_slides
'    pptx.add_all_slides
    
End Sub

Sub ab()
    Call make_graph
    Call pic_to_pptx
End Sub
