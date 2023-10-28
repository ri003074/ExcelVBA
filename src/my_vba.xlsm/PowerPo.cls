VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PowerPo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    Dim power_point As PowerPoint.Application
    Dim active_presentation As PowerPoint.presentation

Sub activate_powerpoint()
    
    Set power_point = GetObject("", "PowerPoint.Application")
    Set active_presentation = power_point.ActivePresentation
    
    power_point.visible = True
                            
End Sub


Sub setup_new_powerpoint()
    
    Set power_point = CreateObject("PowerPoint.Application")
    Set active_presentation = power_point.Presentations.Add
    
    power_point.visible = True

End Sub


Sub add_slide(ByVal layout As Integer)
    
    With active_presentation
        Dim slide_count As Integer: slide_count = active_presentation.Slides.Count
        
        .Slides.Add Index:=slide_count + 1, layout:=layout
        .Slides(slide_count + 1).Select
    End With
    
End Sub

Sub add_picture(ByVal file_path As String)
    
    Dim object_file_sys As Object: Set object_file_sys = CreateObject("Scripting.FileSystemObject")
    
    With active_presentation
        Dim file_name As String: file_name = object_file_sys.GetBaseName(file_path)
        
       .Slides(.Slides.Count).Shapes.AddPicture file_path, -1, -1, 0, 0
       .Slides(.Slides.Count).Shapes(1).TextFrame.TextRange.text = file_name
    End With
    
End Sub


Sub add_pictures(ByVal slide_layout As Integer, ByVal pic_width As Integer, ByVal pic_top1 As Integer, ByVal pic_top2 As Integer, ByVal vertical As Integer, ByVal horizontal As Integer, ParamArray file_paths() As Variant)

    Dim tmp_file_paths As Variant: tmp_file_paths = file_paths
    Dim pic_per_slide_count As Integer: pic_per_slide_count = CalcArrayLength(tmp_file_paths)
    
    Dim file_path1 As String: file_path1 = file_paths(0)
    Dim arr2 As Variant: arr2 = get_file_list(file_path1, "png")
    Dim slide_count As Integer: slide_count = CalcArrayLength(arr2)
    
    Dim data() As String
    ReDim data(pic_per_slide_count, slide_count) As String
    
    Dim i As Integer
    Dim j As Integer

    For i = 0 To pic_per_slide_count - 1
        Dim fp As String: fp = file_paths(i)
        Dim arr3 As Variant: arr3 = get_file_list(fp, "png")

        For j = 0 To slide_count - 1
            data(i, j) = arr3(j)
        Next j
    Next i
    
    Dim pic_left1 As Integer: pic_left1 = active_presentation.PageSetup.SlideWidth * 0.125 - pic_width * 0.5
    Dim pic_left2 As Integer: pic_left2 = active_presentation.PageSetup.SlideWidth * 0.166 - pic_width * 0.5
    Dim pic_left3 As Integer: pic_left3 = active_presentation.PageSetup.SlideWidth * 0.25 - pic_width * 0.5
    Dim pic_left4 As Integer: pic_left4 = active_presentation.PageSetup.SlideWidth * 0.375 - pic_width * 0.5
    Dim pic_left5 As Integer: pic_left5 = active_presentation.PageSetup.SlideWidth * 0.5 - pic_width * 0.5
    Dim pic_left6 As Integer: pic_left6 = active_presentation.PageSetup.SlideWidth * 0.625 - pic_width * 0.5
    Dim pic_left7 As Integer: pic_left7 = active_presentation.PageSetup.SlideWidth * 0.75 - pic_width * 0.5
    Dim pic_left8 As Integer: pic_left8 = active_presentation.PageSetup.SlideWidth * 0.834 - pic_width * 0.5
    Dim pic_left9 As Integer: pic_left9 = active_presentation.PageSetup.SlideWidth * 0.875 - pic_width * 0.5

    Dim pic_lefts() As Integer
    Dim pic_tops() As Integer
    ReDim pic_lefts(pic_per_slide_count) As Integer
    ReDim pic_tops(pic_per_slide_count) As Integer
    
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
            
                Dim pic As Object: Set pic = .Slides(.Slides.Count).Shapes.AddPicture(data(j, i), -1, -1, pic_lefts(j), pic_tops(j))
                pic.width = pic_width
                
                Dim object_file_sys As Object: Set object_file_sys = CreateObject("Scripting.FileSystemObject")
                Dim file_name As String: file_name = object_file_sys.GetBaseName(data(j, i))
                
                Dim textbox_height As Integer: textbox_height = 30
                add_textbox file_name, pic_lefts(j), pic_tops(j) - textbox_height, pic_width, textbox_height
            Next j
        End With
    Next i

End Sub


Sub add_all_slides()
        
    With active_presentation
        Dim i As Integer
        For i = 1 To 36
            .Slides.Add Index:=.Slides.Count + 1, layout:=i
        Next i
    End With

End Sub


Sub delete_all_slides()
        
    With active_presentation
        Dim i As Integer
        For i = .Slides.Count To 1 Step -1
            .Slides(i).Delete
        Next i
    End With

End Sub


Sub add_textbox(ByVal text As String, ByVal left As Integer, ByVal top As Integer, ByVal width As Integer, ByVal height As Integer)
    
    With active_presentation
        Dim shp As Object: Set shp = .Slides(.Slides.Count).Shapes.AddTextbox(msoTextOrientationHorizontal, left, top, width, height)
        
        shp.TextFrame.TextRange.text = text
        shp.TextFrame.TextRange.ParagraphFormat.Alignment = 2
        shp.TextFrame.VerticalAnchor = 3
    End With
    
End Sub
