Attribute VB_Name = "CrowSample"
Option Explicit

Sub all_trtf()
        Dim crow As crow
        Set crow = New crow
    
        crow.add_chart xlLineMarkers
        crow.set_chart_title title:="abcdefg"
        
        crow.set_axis_title xlValue, "y title", 20
        crow.set_axis_title xlCategory, "x title", 20
        
        crow.set_tick xlValue, 0, 100, 20
        crow.set_line_visible False
        
        crow.set_tick_label xlValue, 10
        crow.set_tick_label xlCategory, 20
        
        crow.set_legend 20, xlLegendPositionRight
        
        crow.resize_graph 300, 400
        crow.relocate_graph "D4"
        crow.save_png
        
End Sub

Sub all_sin()
    Dim crow As crow
    Set crow = New crow
    
    Dim input_file As String: input_file = "C:\Users\ri003\Documents\Programming\ExcelVBA\data\Book2.csv"
    crow.save_as_xlsx (input_file)
    crow.open_file
    crow.add_chart xlLineMarkers
    crow.set_chart_title "sin"
    
    crow.set_axis_title xlValue, "V", 10
    crow.set_axis_title xlCategory, "Time", 20
    
    crow.set_tick xlValue, -2, 2, 0.5
    
    crow.set_tick_label xlValue, 8
    crow.set_tick_label xlCategory, 8, 50
    
    crow.set_legend 10, xlLegendPositionRight
    
    crow.resize_graph 300, 500
    crow.relocate_graph "E5"
    crow.save_png
    crow.close_workbook
End Sub

