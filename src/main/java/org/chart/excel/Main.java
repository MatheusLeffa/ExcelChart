package org.chart.excel;

import com.aspose.cells.*;

public class Main {
    public static void main(String[] args) throws Exception {

        String filePath = FileSelector.getFilePath();
        Workbook workbook = new Workbook(filePath);
        Worksheet worksheet = workbook.getWorksheets().get(1);

        ChartCollection charts = worksheet.getCharts();
        int chartIndex = charts.add(ChartType.COLUMN, 5, 0, 20, 10);

        Chart chart = worksheet.getCharts().get(chartIndex);
        chart.setChartDataRange("", true);

        try {
            workbook.save("Test_out.xlsx", SaveFormat.XLSX);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}