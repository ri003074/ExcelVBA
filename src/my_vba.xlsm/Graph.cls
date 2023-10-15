VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private xlsx_file_path As String


Sub save_as_csv(ByVal csv_file_path As String)
    Application.DisplayAlerts = False
    
    Workbooks.Open (csv_file_path)
    
    Dim object_file_sys As Object: Set object_file_sys = CreateObject("Scripting.FileSystemObject")
    Dim file_name As String: file_name = object_file_sys.GetBaseName(csv_file_path)
    
    xlsx_file_path = Replace(csv_file_path, "csv", "xlsx")
    
    ActiveWorkbook.SaveAs Filename:=xlsx_file_path, FileFormat:=xlOpenXMLWorkbook
    ActiveWorkbook.Close

End Sub
 
 
Sub open_file()
    
    Workbooks.Open (xlsx_file_path)

End Sub


Sub add_chart(ByVal chart_type As Variant)
    
    Call delete_chart
    
    Dim range As range: Set range = ActiveSheet.range("C5:I15")
    Dim chart_object As ChartObject: Set chart_object = ActiveSheet.ChartObjects.Add(range.left, range.top, range.width, range.height)
    Dim chart As chart: Set chart = chart_object.chart
    
    Dim last_row As Long: last_row = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    Dim last_column As Long: last_column = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    
    Dim object_file_sys As Object: Set object_file_sys = CreateObject("Scripting.FileSystemObject")
    Dim file_name As String: file_name = object_file_sys.GetBaseName(ActiveWorkbook.FullName)
    
    With chart
        .SetSourceData Source:=Cells(1, 1).Resize(last_row, last_column)
        .ChartType = chart_type
        .HasTitle = True
        .charttitle.text = file_name
        .Legend.Position = xlLegendPositionBottom
        .ChartStyle = 332
    End With
    
End Sub

Sub delete_chart()
    Dim chart_object As ChartObject
    
    For Each chart_object In ActiveSheet.ChartObjects
        chart_object.Delete
    Next chart_object
    
End Sub

Sub resize_graph(ByVal height As Long, ByVal width As Long)

    Dim chart_object As ChartObject
    
    For Each chart_object In ActiveSheet.ChartObjects
        With chart_object
            .width = width
            .height = height
        End With
    Next chart_object
    
End Sub


Sub relocate_graph(ByVal base_cell As String)
    
    Dim chart_object As ChartObject
    Dim top As Long: top = range(base_cell).Row
    Dim left As Long: left = range(base_cell).Column
    
    For Each chart_object In ActiveSheet.ChartObjects
        With chart_object
            .left = ActiveSheet.Cells(top, left).left
            .top = ActiveSheet.Cells(top, left).top
        End With
        
        top = chart_object.BottomRightCell.Row + 1
    Next chart_object
    
End Sub


Sub set_axis_title(ByVal axis_type As Long, ByVal title As String, ByVal font_size As Long, Optional ByVal tick_mark_position As Long = xlTickMarkNone)

    Dim chart_object As ChartObject
    
    For Each chart_object In ActiveSheet.ChartObjects
        With chart_object
            Dim axis As axis: Set axis = chart_object.chart.Axes(axis_type)
            
            If title <> "" Then
                axis.HasTitle = True
                axis.AxisTitle.Characters.text = title
                axis.AxisTitle.Format.TextFrame2.TextRange.Font.Size = font_size
            End If
            
            axis.MajorTickMark = tick_mark_position
        End With
    
    Next chart_object
    
End Sub


Sub set_tick(ByVal axis_type As Long, ByVal minimum As Double, ByVal maximum As Double, ByVal resolution As Double)

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


Sub save_png()

    Dim chart_object As ChartObject
    For Each chart_object In ActiveSheet.ChartObjects
        With chart_object
            Dim address As String
            address = chart_object.TopLeftCell.address(False, False)
            range(address).Select
            Dim c As chart
            Set c = chart_object.chart
            
            Dim chart_title As charttitle
            Set chart_title = chart_object.chart.charttitle
            
            Call c.Export(ActiveWorkbook.Path + "\" + chart_title.text + ".png")
        End With
    Next chart_object

End Sub
