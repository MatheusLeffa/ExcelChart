package org.chart.excel;

import com.aspose.cells.*;

import java.util.Map;

public class Main {
    public static void main(String[] args) throws Exception {

        String filePath = FileSelector.getFilePath();
        Workbook workbook = new Workbook(filePath);

        WorksheetCollection worksheetCollection = workbook.getWorksheets();
        Worksheet worksheet0 = worksheetCollection.get(0);
        Worksheet worksheet1 = worksheetCollection.add("Gráficos");

        ChartData chartData1 = new ChartData();
        ChartData chartData2 = new ChartData();

        Map<String, Integer> dataSeries1 = chartData1.getDataSeries(worksheet0, "G");
        Map<String, Integer> dataSeries2 = chartData2.getDataSeries(worksheet0, "N");

        WorksheetEditor worksheetEditor = new WorksheetEditor(workbook, worksheet1);

        worksheetEditor.addValues(dataSeries1);
        worksheetEditor.addValues(dataSeries2);

        worksheetEditor.sort(dataSeries1);
        worksheetEditor.sort(dataSeries2);

        Chart chart1 = worksheetEditor.createChart(1, 4, 20, 20);
        Chart chart2 = worksheetEditor.createChart(22, 4, 38, 18);
        chart1.setChartDataRange("A1:B" + dataSeries1.size(), true);
        chart2.setChartDataRange("C1:D" + dataSeries2.size(), true);

        Formatter.chartFormatter(chart1, "Produto/Sistema");
        Formatter.chartFormatter(chart2, "Status");

        try {
            workbook.save("Boletim_out.xlsx", SaveFormat.XLSX);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}