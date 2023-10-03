Attribute VB_Name = "Module1"
Option Explicit


Sub resize_graph(height As Integer, width As Integer)

    Dim chart_object As ChartObject
    
    For Each chart_object In ActiveSheet.ChartObjects
        With chart_object
            .width = width
            .height = height
        End With
    Next chart_object
    
End Sub

Sub relocate_graph(base_cell As String)
    
    Dim chart_object As ChartObject
    Dim top As Integer
    Dim left As Integer
    
    top = Range(base_cell).Row
    left = Range(base_cell).Column
    
    For Each chart_object In ActiveSheet.ChartObjects
        With chart_object
            .left = ActiveSheet.Cells(top, left).left
            .top = ActiveSheet.Cells(top, left).top
        End With
        
        top = chart_object.BottomRightCell.Row + 1
    Next chart_object
    
End Sub


Sub set_axis_title(title As String, font_size As Integer)

    Dim chart_object As ChartObject
    
    For Each chart_object In ActiveSheet.ChartObjects
        Dim axis As axis
        With chart_object
            Set axis = chart_object.Chart.Axes(xlValue)
            axis.HasTitle = True
            axis.AxisTitle.Characters.Text = title
            axis.AxisTitle.Format.TextFrame2.TextRange.Font.Size = font_size
        End With
    
    Next chart_object
    
End Sub




Sub All()

    Call resize_graph(300, 300)
    Call relocate_graph("C4")
    Call set_axis_title("mV", 20)
    
End Sub
