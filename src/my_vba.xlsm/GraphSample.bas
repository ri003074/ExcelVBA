Attribute VB_Name = "GraphSample"
Option Explicit

Sub all_trtf()
        Dim graph As graph
        Set graph = New graph
    
        graph.add_chart xlLineMarkers
        graph.set_chart_title "abcdefg"
        
        graph.set_axis_title xlValue, "y title", 20
        graph.set_axis_title xlCategory, "x title", 20
        
        graph.set_tick xlValue, 0, 100, 20
        graph.set_line_visible False
        
        graph.set_tick_label xlValue, 10
        graph.set_tick_label xlCategory, 20
        
        graph.set_legend 20, xlLegendPositionRight
        
        graph.resize_graph 300, 400
        graph.relocate_graph "D4"
        graph.save_png
        
End Sub

Sub all_sin()
    Dim graph As graph
    Set graph = New graph
    
    graph.add_chart xlLineMarkers
    graph.set_chart_title "sin"
    
    graph.set_axis_title xlValue, "V", 10
    graph.set_axis_title xlCategory, "Time", 20
    
    graph.set_tick xlValue, -2, 2, 0.5
    
    graph.set_tick_label xlValue, 8
    graph.set_tick_label xlCategory, 8, 50
    
    graph.set_legend 10, xlLegendPositionRight
    
    graph.resize_graph 300, 500
    graph.relocate_graph "E5"
    
End Sub

