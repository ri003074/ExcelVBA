Attribute VB_Name = "Module1"
Option Explicit

Sub make_graph()
    Dim graph As crow
    Set graph = New crow
    
    Dim input_folder As String: input_folder = "C:\Users\ri003\Documents\Programming\ExcelVBA\data"
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
            ActiveWorkbook.Close
        End If
    Next
End Sub


Sub pic_to_pptx()
    Dim input_folder As String: input_folder = "C:\Users\ri003\Documents\Programming\ExcelVBA\data"
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim pptx As PowerPo: Set pptx = New PowerPo
    Dim layout_number As Long: layout_number = 16
'    Dim layout_number As Integer: layout_number = 29
'    Dim layout_number As Integer: layout_number = 11
        
'    pptx.activate_powerpoint
    pptx.setup_new_powerpoint
    pptx.delete_all_slides
   
    Dim file As Object
    For Each file In fso.GetFolder(input_folder).Files
        If LCase(file.Name) Like "*.png" Then
            Dim file_path As String: file_path = file
            
            pptx.add_slide layout_number
            pptx.add_picture file_path
        End If
    Next
     
    'pptx.add_slide 16
    'pptx.delete_all_slides
    'pptx.add_all_slides
    
End Sub

Sub add_text_box()
    Dim input_folder As String: input_folder = "C:\Users\ri003\Documents\Programming\ExcelVBA\data"
    get_file_list input_folder, "png"
        
    Dim pptx As PowerPo: Set pptx = New PowerPo
    'pptx.activate_powerpoint
    pptx.setup_new_powerpoint
    pptx.delete_all_slides
    pptx.add_pictures 11, 200, 150, 350, 2, 2, input_folder, input_folder, input_folder, input_folder
    
End Sub

Sub ab()
    Call make_graph
    Call pic_to_pptx
End Sub
