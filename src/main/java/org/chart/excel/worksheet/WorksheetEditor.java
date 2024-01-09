package org.chart.excel.worksheet;

import com.aspose.cells.*;

import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.stream.Collectors;

public class WorksheetEditor {

    private final Worksheet worksheet;
    private final Cells cells;
    private final ChartCollection charts;


    public WorksheetEditor(Worksheet worksheet) {
        this.worksheet = worksheet;
        this.cells = worksheet.getCells();
        this.charts = worksheet.getCharts();
    }


    public void addValues(Map<String, Integer> dataSeries) {

        if (cells.isBlankColumn(0)) {
            putColumnValues(dataSeries, "A", "B", "Produto/Sistema");
        } else {
            putColumnValues(dataSeries, "C", "D", "Status");
        }
    }

    private void putColumnValues(Map<String, Integer> dataSeries, String firstColumn, String secondColumn, String columnTitle) {
        int i = 2;

        cells.get(firstColumn + "1").putValue(columnTitle);
        cells.get(secondColumn + "1").putValue("Cont");

        for (String data : dataSeries.keySet()) {
            cells.get(firstColumn + i).putValue(data);
            cells.get(secondColumn + i).putValue(dataSeries.get(data));
            i++;
        }
    }

    public Chart createChart(int upperLeftRow, int upperLeftColumn, int lowerRightRow, int lowerRightColumn) {
        int chartIndex = charts.add(ChartType.COLUMN, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);
        return worksheet.getCharts().get(chartIndex);
    }
}
