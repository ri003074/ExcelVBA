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
    
    power_point.Visible = True
                            
End Sub


Sub setup_new_powerpoint()
    
    Set power_point = CreateObject("PowerPoint.Application")
    Set active_presentation = power_point.Presentations.Add
    
    power_point.Visible = True

End Sub


Sub add_slide(layout As Integer)
    
    With active_presentation
        Dim slide_count As Integer: slide_count = active_presentation.Slides.Count
        
        .Slides.Add Index:=slide_count + 1, layout:=layout
        .Slides(slide_count + 1).Shapes(1).TextFrame.TextRange.Text = "a"
    End With
    
End Sub

Sub add_picture(file_path As String)
    
    Dim object_file_sys As Object: Set object_file_sys = CreateObject("Scripting.FileSystemObject")
    
    With active_presentation
        Dim file_name As String: file_name = object_file_sys.GetBaseName(file_path)
        
       .Slides(.Slides.Count).Shapes.AddPicture file_path, -1, -1, 0, 0
       .Slides(.Slides.Count).Shapes(1).TextFrame.TextRange.Text = file_name
    End With
    
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
