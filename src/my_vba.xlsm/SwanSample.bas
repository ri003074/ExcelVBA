Attribute VB_Name = "SwanSample"
Option Explicit

Sub picture_to_pptx1()

    Dim input_folder As String: input_folder = "C:\Users\ri003\Documents\Programming\ExcelVBA\data"
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim swan As swan: Set swan = New swan
    Dim layout_number As Long: layout_number = 16 '29,11
    
    'swan.activate_powerpoint
    swan.setup_new_powerpoint
    swan.delete_all_slides
   
    Dim file As Object
    For Each file In fso.GetFolder(input_folder).Files
        If LCase(file.Name) Like "*.png" Then
            Dim file_path As String: file_path = file
            
            swan.add_slide layout_number
            swan.add_picture file_path
        End If
    Next
       
End Sub


Sub picture_to_pptx2()

    Dim Util As Util
    Set Util = New Util

    Dim input_folder As String: input_folder = "C:\Users\ri003\Documents\Programming\ExcelVBA\data"
    Util.get_file_list input_folder, "png"
        
    Dim swan As swan: Set swan = New swan
    'swan.activate_powerpoint
    swan.setup_new_powerpoint
    swan.delete_all_slides
    
    Dim rep_title As Object
    Set rep_title = CreateObject("Scripting.Dictionary")
    rep_title.Add "_.*", ""
    
    Dim rep_each As Object
    Set rep_each = CreateObject("Scripting.Dictionary")
    rep_each.Add "Book.*_", ""
    
    swan.add_pictures 11, 200, 150, 350, 2, 2, rep_title, rep_each, input_folder, input_folder, input_folder, input_folder
    
End Sub
