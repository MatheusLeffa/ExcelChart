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

        Map<String, Integer> dataSeries1 = ChartData.getDataSeries(worksheet0, "G");
        Map<String, Integer> dataSeries2 = ChartData.getDataSeries(worksheet0, "N");


        WorksheetEditor worksheetEditor = new WorksheetEditor(worksheet1);
        worksheetEditor.addValues(dataSeries1);
        worksheetEditor.addValues(dataSeries2);
        Chart chart1 = worksheetEditor.createChart(1, 6, 20, 20);
        Chart chart2 = worksheetEditor.createChart(26, 6, 38, 20);

        // Configurar dados do gráfico
        chart1.setChartDataRange("A1:B" + dataSeries1.size(), true);
        chart2.setChartDataRange("C1:D" + dataSeries2.size(), true);

        // Formatação gráfico
        ChartEditor.chartFormatter(chart1, "Produto/Sistema");
        ChartEditor.chartFormatter(chart2, "Status");

        // Salvar o arquivo
        try {
            workbook.save(System.getProperty("user.home") + "\\OneDrive - Sicredi\\Desktop\\Boletim_out.xlsx", SaveFormat.XLSX);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}