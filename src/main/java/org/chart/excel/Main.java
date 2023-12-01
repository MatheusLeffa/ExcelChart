package org.chart.excel;

import com.aspose.cells.*;

import java.util.Map;

public class Main {
    public static void main(String[] args) throws Exception {

        String filePath = FileSelector.getFilePath();
        Workbook workbook = new Workbook(filePath);

        WorksheetCollection worksheetCollection = workbook.getWorksheets();
        Worksheet worksheet0 = worksheetCollection.get(0);
        Worksheet worksheet1 = worksheetCollection.getCount() > 1 ? worksheetCollection.get(1) : worksheetCollection.add("Gráficos");

        Map<String, Integer> ocorrencias = ChartData.getOcorrenciasSerie(worksheet0);

        WorksheetEditor worksheetEditor = new WorksheetEditor(worksheet1);
        worksheetEditor.addValues(ocorrencias);
        Chart chart = worksheetEditor.createChart();

        // Configurar dados do gráfico
        chart.setChartDataRange("A1:B" + ocorrencias.size(), true);

        // Formatação gráfico
        ChartEditor.chartFormatter(chart);

        // Salvar o arquivo
        try {
            workbook.save("Test_out.xlsx", SaveFormat.XLSX);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}