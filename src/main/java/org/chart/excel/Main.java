package org.chart.excel;

import com.aspose.cells.*;

import java.io.File;
import java.util.Map;

public class Main {
    public static void main(String[] args) throws Exception {

        String filePath = FileSelector.getFilePath();
        Workbook workbook = new Workbook(filePath);

        WorksheetCollection worksheetCollection = workbook.getWorksheets();
        Worksheet tabelaBoletim = worksheetCollection.get(0);
        Worksheet tabelaGraficos = worksheetCollection.add("Gráficos");

        ChartData colunaSistema = new ChartData(workbook, tabelaBoletim, "G");
        ChartData colunaStatus = new ChartData(workbook, tabelaBoletim, "N");

        WorksheetEditor EditorTabelaGraficos = new WorksheetEditor(workbook, tabelaGraficos);
        EditorTabelaGraficos.addValues(colunaSistema.getDataSeries());
        EditorTabelaGraficos.addValues(colunaStatus.getDataSeries());

        Chart graficoSistema = EditorTabelaGraficos.createChart(1, 4, 20, 20);
        Chart graficoStatus = EditorTabelaGraficos.createChart(22, 4, 38, 18);

        graficoSistema.setChartDataRange("A1:B" + (colunaSistema.getColunaSize()), true);
        graficoStatus.setChartDataRange("C1:D" + (colunaStatus.getColunaSize()), true);

        ChartFormatter.chartFormatter(graficoSistema, "Sistema");
        ChartFormatter.chartFormatter(graficoStatus, "Status");

        try {
            workbook.save("Boletim_out.xlsx", SaveFormat.XLSX);
            java.awt.Desktop.getDesktop().open(new File("Boletim_out.xlsx"));
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}