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
    
    Dim slide_count As Integer
    
    With active_presentation
        slide_count = active_presentation.Slides.Count
        .Slides.Add Index:=slide_count + 1, layout:=layout
        .Slides(slide_count + 1).Shapes(1).TextFrame.TextRange.Text = "a"
    End With
    
End Sub

Sub add_picture(file_path As String)
    Dim slide_count As Integer
    With active_presentation
        slide_count = active_presentation.Slides.Count
        .Slides(slide_count).Shapes.AddPicture file_path, -1, -1, 0, 0
        .Slides(slide_count).Shapes(1).TextFrame.TextRange.Text = Dir(file_path)
    End With
    

End Sub


Sub add_all_slides()
        
    Dim i As Integer
    Dim slide_count As Integer

    With active_presentation
        For i = 1 To 40
            slide_count = active_presentation.Slides.Count
            .Slides.Add Index:=slide_count + 1, layout:=i
'            .Slides(slide_count + 1).Shapes(1).TextFrame.TextRange.Text = i
        Next i
    End With

End Sub


Sub delete_all_slides()
    
    Dim slide_count As Integer
    Dim i As Integer
        
    With active_presentation
        slide_count = active_presentation.Slides.Count
        For i = slide_count To 1 Step -1
            .Slides(i).Delete
        Next i
    End With

End Sub
