VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Swan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Dim power_point As PowerPoint.Application
    Dim active_presentation As PowerPoint.presentation
    Dim slide_width As Long

Sub activate_powerpoint()
    
    Set power_point = GetObject("", "PowerPoint.Application")
    Set active_presentation = power_point.ActivePresentation
    slide_width = active_presentation.PageSetup.SlideWidth
    
    power_point.visible = True
                            
End Sub


Sub setup_new_powerpoint()
    
    Set power_point = CreateObject("PowerPoint.Application")
    Set active_presentation = power_point.Presentations.Add
    
    power_point.visible = True

End Sub


Sub add_slide(ByVal layout As Long, Optional ByVal title As String = "")
    
    With active_presentation
        Dim slide_count As Long: slide_count = active_presentation.Slides.Count
        
        .Slides.Add Index:=slide_count + 1, layout:=layout
        .Slides.Item(slide_count + 1).Select
        
        If title <> "" Then
            .Slides.Item(slide_count + 1).Shapes.Item(1).TextFrame.TextRange.text = title
        End If
    End With
    
End Sub

Sub add_picture(ByVal file_path As String)
    
    Dim object_file_sys As Object: Set object_file_sys = CreateObject("Scripting.FileSystemObject")
    
    With active_presentation
        Dim file_name As String: file_name = object_file_sys.GetBaseName(file_path)
        
       .Slides.Item(.Slides.Count).Shapes.AddPicture file_path, -1, -1, 0, 0
       .Slides.Item(.Slides.Count).Shapes.Item(1).TextFrame.TextRange.text = file_name
    End With
    
End Sub


Sub add_pictures(ByVal slide_layout As Long, ByVal pic_width As Long, ByVal pic_top1 As Long, ByVal pic_top2 As Long, ByVal vertical As Long, ByVal horizontal As Long, ByVal rep_title As Object, ByVal rep_each As Object, ParamArray file_paths() As Variant)

    Dim Util As Util
    Set Util = New Util

    Dim tmp_file_paths As Variant: tmp_file_paths = file_paths
    Dim pic_per_slide_count As Long: pic_per_slide_count = Util.CalcArrayLength(tmp_file_paths)
    
    Dim tmp_file_path As String: tmp_file_path = file_paths(0)
    Dim tmp_arr As Variant: tmp_arr = Util.get_file_list(tmp_file_path, "png")
    Dim slide_count As Long: slide_count = Util.CalcArrayLength(tmp_arr)
    
    Dim data() As String
    ReDim data(pic_per_slide_count, slide_count) As String
    
    Dim i As Long
    Dim j As Long

    For i = 0 To pic_per_slide_count - 1
        Dim fp As String: fp = file_paths(i)
        Dim file_list As Variant: file_list = Util.get_file_list(fp, "png")

        For j = 0 To slide_count - 1
            data(i, j) = file_list(j)
        Next j
    Next i
    
    Dim pic_left1 As Long: pic_left1 = slide_width * 0.125 - pic_width * 0.5
    Dim pic_left2 As Long: pic_left2 = slide_width * 0.166 - pic_width * 0.5
    Dim pic_left3 As Long: pic_left3 = slide_width * 0.25 - pic_width * 0.5
    Dim pic_left4 As Long: pic_left4 = slide_width * 0.375 - pic_width * 0.5
    Dim pic_left5 As Long: pic_left5 = slide_width * 0.5 - pic_width * 0.5
    Dim pic_left6 As Long: pic_left6 = slide_width * 0.625 - pic_width * 0.5
    Dim pic_left7 As Long: pic_left7 = slide_width * 0.75 - pic_width * 0.5
    Dim pic_left8 As Long: pic_left8 = slide_width * 0.834 - pic_width * 0.5
    Dim pic_left9 As Long: pic_left9 = slide_width * 0.875 - pic_width * 0.5

    Dim pic_lefts() As Long
    Dim pic_tops() As Long
    ReDim pic_lefts(pic_per_slide_count) As Long
    ReDim pic_tops(pic_per_slide_count) As Long
    
    If vertical = 1 And horizontal = 1 Then
        pic_lefts(0) = pic_left5
        pic_tops(0) = pic_top1
    End If
        
    If vertical = 1 And horizontal = 2 Then
        pic_lefts(0) = pic_left3
        pic_lefts(1) = pic_left7
        pic_tops(0) = pic_top1
        pic_tops(1) = pic_top1
    End If
    
    If vertical = 2 And horizontal = 2 Then
        pic_lefts(0) = pic_left3
        pic_lefts(1) = pic_left7
        pic_lefts(2) = pic_left3
        pic_lefts(3) = pic_left7
        pic_tops(0) = pic_top1
        pic_tops(1) = pic_top1
        pic_tops(2) = pic_top2
        pic_tops(3) = pic_top2
    End If
    
    If vertical = 1 And horizontal = 3 Then
        pic_lefts(0) = pic_left2
        pic_lefts(1) = pic_left5
        pic_lefts(2) = pic_left8
        pic_tops(0) = pic_top1
        pic_tops(1) = pic_top1
        pic_tops(2) = pic_top1
    End If
    
    If vertical = 1 And horizontal = 4 Then
        pic_lefts(0) = pic_left1
        pic_lefts(1) = pic_left4
        pic_lefts(2) = pic_left6
        pic_lefts(3) = pic_left9
        pic_tops(0) = pic_top1
        pic_tops(1) = pic_top1
        pic_tops(2) = pic_top1
        pic_tops(3) = pic_top1
    End If
    

    For i = 0 To slide_count - 1
        With active_presentation
            add_slide (slide_layout)
            For j = 0 To pic_per_slide_count - 1
            
                Dim pic As Object: Set pic = .Slides.Item(.Slides.Count).Shapes.AddPicture(data(j, i), -1, -1, pic_lefts(j), pic_tops(j))
                pic.width = pic_width
                
                Dim object_file_sys As Object: Set object_file_sys = CreateObject("Scripting.FileSystemObject")
                Dim file_name As String: file_name = object_file_sys.GetBaseName(data(j, i))
                Dim file_name_each As String: file_name_each = file_name
                
                Dim key As Variant
                Dim re As Object
                Set re = CreateObject("VBScript.RegExp")
                
                For Each key In rep_each
                    re.Pattern = key
                    file_name_each = re.Replace(file_name_each, rep_each(key))
                Next
                
                Dim textbox_height As Long: textbox_height = 30
                add_textbox file_name_each, pic_lefts(j), pic_tops(j) - textbox_height, pic_width, textbox_height
                
                If j = 0 Then
                                   
                    Dim file_name_title As String
                    
                    For Each key In rep_title
                        re.Pattern = key
                        file_name_title = re.Replace(file_name, rep_title(key))
                    Next
                    
                    .Slides.Item(.Slides.Count).Shapes.Item(1).TextFrame.TextRange.text = file_name_title
                End If
    
            Next j
        End With
    Next i

End Sub


Sub add_all_slides()
        
    With active_presentation
        Dim i As Long
        For i = 1 To 36
            .Slides.Add Index:=.Slides.Count + 1, layout:=i
        Next i
    End With

End Sub


Sub delete_all_slides()
        
    With active_presentation
        Dim i As Long
        For i = .Slides.Count To 1 Step -1
            .Slides.Item(i).Delete
        Next i
    End With

End Sub


Sub add_textbox(ByVal text As String, ByVal left As Long, ByVal top As Long, ByVal width As Long, ByVal height As Long)
    
    With active_presentation
        Dim shp As Object: Set shp = .Slides.Item(.Slides.Count).Shapes.AddTextbox(msoTextOrientationHorizontal, left, top, width, height)
        
        shp.TextFrame.TextRange.text = text
        shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
        shp.TextFrame.VerticalAnchor = 3
    End With
    
End Sub


Sub add_table(ByVal data, ByVal slide_layout, Optional ByVal column_width As Long = 100, Optional ByVal title As String = "")

    Dim row_max_index As Integer: row_max_index = UBound(data, 1)
    Dim column_max_index As Integer: column_max_index = UBound(data, 2)
    
    add_slide layout:=slide_layout, title:=title
    
    Dim row As Integer
    Dim column As Integer
    With active_presentation
        Dim shp As Object: Set shp = .Slides.Item(.Slides.Count).Shapes.AddTable(row_max_index, column_max_index)
        Dim table As Object: Set table = shp.table
       
        For row = 1 To row_max_index
            For column = 1 To column_max_index
                Dim text_range As Object: Set text_range = table.Cell(row, column).Shape.TextFrame.TextRange
                text_range.text = data(row, column)
                text_range.ParagraphFormat.Alignment = 2
                table.Columns(column).width = column_width
                shp.left = slide_width / 2 - shp.width / 2
            Next
        Next
    End With
    
End Sub


