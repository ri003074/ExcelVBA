Attribute VB_Name = "GraphSample"
Option Explicit

Sub add_chart()

        Dim graph As graph
        Set graph = New graph
    
        graph.add_chart xlLineMarkers
        
End Sub

Sub save_graph_as_png()
        
        Dim graph As graph
        Set graph = New graph
        
        graph.save_png
        
End Sub

Sub set_chart_title()
        
        Dim graph As graph
        Set graph = New graph

        graph.set_chart_title "abc"

End Sub

Sub set_graph_tick()
        
        Dim graph As graph
        Set graph = New graph

        graph.set_tick xlValue, 0, 100, 20

End Sub

Sub set_graph_title()
        
        Dim graph As graph
        Set graph = New graph

        graph.set_axis_title xlValue, "ps", 20
End Sub

Sub resize_graph()
        
        Dim graph As graph
        Set graph = New graph

        graph.resize_graph 300, 400

End Sub

Sub relocate_graph()

        Dim graph As graph
        Set graph = New graph

        graph.relocate_graph "E5"
        
End Sub

Sub set_line_visible()
    
    Dim graph As graph
    Set graph = New graph
    
    graph.set_line_visible msoFalse
    
End Sub

Sub set_legend_font_size()
    
    Dim graph As graph
    Set graph = New graph
    
    graph.set_legend_font_size 10, xlLegendPositionRight

End Sub

Sub set_axis_font_size()
    
    Dim graph As graph
    Set graph = New graph
    
    graph.set_axis_font_size xlValue, 20
    
End Sub

Sub set_tick_font_size()

    Dim graph As graph
    Set graph = New graph
        
    graph.set_axis_tick_font_size xlValue, 20
    graph.set_axis_tick_font_size xlCategory, 10
    
End Sub
