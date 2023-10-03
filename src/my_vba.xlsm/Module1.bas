Attribute VB_Name = "Module1"
Option Explicit


Sub add_chart()
    
    Call delete_chart
    
    Dim chart_object As ChartObject
    Dim chart As chart
    Dim range As range
    
    Set range = Cells(4, 4).Resize(12, 7)
    
    Set chart_object = ActiveSheet.ChartObjects.Add(range.left, range.top, range.width, range.height)
    
    Set chart = chart_object.chart
    
    With chart
        .SetSourceData Source:=Cells(1, 1).Resize(3, 3)
        .ChartType = xlLineMarkers
        .HasTitle = True
        .ChartTitle.Text = "title"
        .Legend.Position = xlLegendPositionBottom
        
    End With
    
End Sub

Sub delete_chart()
    Dim chart_object As ChartObject
    
    For Each chart_object In ActiveSheet.ChartObjects
        chart_object.Delete
    Next chart_object
    
End Sub

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
    
    top = range(base_cell).Row
    left = range(base_cell).Column
    
    For Each chart_object In ActiveSheet.ChartObjects
        With chart_object
            .left = ActiveSheet.Cells(top, left).left
            .top = ActiveSheet.Cells(top, left).top
        End With
        
        top = chart_object.BottomRightCell.Row + 1
    Next chart_object
    
End Sub


Sub set_axis_title(axis_type As Integer, title As String, font_size As Integer, Optional tick_mark_position As Integer = xlTickMarkNone)

    Dim chart_object As ChartObject
    
    For Each chart_object In ActiveSheet.ChartObjects
        Dim axis As axis
        With chart_object
            Set axis = chart_object.chart.Axes(axis_type)
            
            If title <> "" Then
                axis.HasTitle = True
                axis.AxisTitle.Characters.Text = title
                axis.AxisTitle.Format.TextFrame2.TextRange.Font.Size = font_size
            End If
            
            axis.MajorTickMark = tick_mark_position
        End With
    
    Next chart_object
    
End Sub


Sub set_tick(axis_type As Integer, minimum As Double, maximum As Double, resolution As Double)

    Dim chart_object As ChartObject
    
    For Each chart_object In ActiveSheet.ChartObjects
        Dim axis As axis
        With chart_object
            Set axis = chart_object.chart.Axes(axis_type)
            axis.MinimumScale = minimum
            axis.MaximumScale = maximum
            axis.MajorUnit = resolution
        End With
    
    Next chart_object
        
End Sub


Sub All()
    Call add_chart
    Call resize_graph(300, 400)
    Call relocate_graph("D4")
    Call set_tick(xlValue, 0, 120, 20)
    Call set_axis_title(xlValue, "ps", 20)
    Call set_axis_title(xlCategory, "", 20)
    
End Sub
