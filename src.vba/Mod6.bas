Sub trocaDadosChart()
'
' trocaDadosChart Macro
'

'
    Range("R14").Select
    ActiveSheet.ChartObjects("Gráfico 14").Activate
    ActiveChart.PlotArea.Select
    ActiveChart.SetSourceData Source:=Range( _
        "'Análise Basiléia_graf'!$S$7:$AU$7,'Análise Basiléia_graf'!$S$12:$AU$13")
    ActiveChart.ChartArea.Select
    'Range("T18").Select
End Sub