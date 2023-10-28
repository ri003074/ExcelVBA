Attribute VB_Name = "GraphSample"
Option Explicit

Sub add_chart()

        Dim Graph As Graph
        Set Graph = New Graph
    
        Graph.add_chart xlLineMarkers
        
End Sub

Sub save_graph_as_png()
        
        Dim Graph As Graph
        Set Graph = New Graph
        
        Graph.save_png
        
End Sub

Sub set_chart_title()
        
        Dim Graph As Graph
        Set Graph = New Graph

        Graph.set_chart_title "abc"

End Sub

Sub set_graph_tick()
        
        Dim Graph As Graph
        Set Graph = New Graph

        Graph.set_tick xlValue, 0, 100, 20

End Sub

Sub set_graph_title()
        
        Dim Graph As Graph
        Set Graph = New Graph

        Graph.set_axis_title xlValue, "ps", 10
End Sub


Sub resize_graph()
        
        Dim Graph As Graph
        Set Graph = New Graph

        Graph.resize_graph 300, 400

End Sub

Sub relocate_graph()

        Dim Graph As Graph
        Set Graph = New Graph

        Graph.relocate_graph "E5"
        
End Sub
