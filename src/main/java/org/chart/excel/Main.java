package org.chart.excel;

import com.aspose.cells.*;

import java.util.Map;

public class Main {
    public static void main(String[] args) throws Exception {

        String filePath = FileSelector.getFilePath();
        Workbook workbook = new Workbook(filePath);

        WorksheetCollection worksheetCollection = workbook.getWorksheets();
        Worksheet tabelaBoletim = worksheetCollection.get(0);
        Worksheet tabelaGraficos = worksheetCollection.add("Gráficos");

        Map<String, Integer> colunaSistema = ChartData.getDataSeries(tabelaBoletim, "G");
        Map<String, Integer> colunaStatus = ChartData.getDataSeries(tabelaBoletim, "N");

        WorksheetEditor EditorTabelaGraficos = new WorksheetEditor(workbook, tabelaGraficos);

        EditorTabelaGraficos.addValues(colunaSistema);
        EditorTabelaGraficos.addValues(colunaStatus);

        EditorTabelaGraficos.sort(colunaSistema);
        EditorTabelaGraficos.sort(colunaStatus);

        Chart graficoSistema = EditorTabelaGraficos.createChart(1, 4, 20, 20);
        Chart graficoStatus = EditorTabelaGraficos.createChart(22, 4, 38, 18);

        graficoSistema.setChartDataRange("A1:B" + colunaSize(colunaSistema), true);
        graficoStatus.setChartDataRange("C1:D" + colunaSize(colunaStatus), true);

        Formatter.chartFormatter(graficoSistema, "Sistema");
        Formatter.chartFormatter(graficoStatus, "Status");

        try {
            workbook.save("Boletim_out.xlsx", SaveFormat.XLSX);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    public static int colunaSize(Map<String, Integer> coluna){
        if (coluna.size() <= 1){
            return 1;
        }
        return coluna.size();
    }


}