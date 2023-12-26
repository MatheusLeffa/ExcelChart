package org.chart.excel;

import com.aspose.cells.*;

import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.stream.Collectors;

public class WorksheetEditor {

    Workbook workbook;
    Worksheet worksheet;


    public WorksheetEditor(Workbook workbook, Worksheet worksheet) {
        this.workbook = workbook;
        this.worksheet = worksheet;
    }

    public void addValues(Map<String, Integer> dataSeries) {
        Cells cells = worksheet.getCells();
        int i = 1;

        if (cells.isBlankColumn(0)) {
            for (String data : dataSeries.keySet()) {
                cells.get("A" + i).putValue(data);
                cells.get("B" + i).putValue(dataSeries.get(data));
                i++;
            }
        } else {
            for (String data : dataSeries.keySet()) {
                cells.get("C" + i).putValue(data);
                cells.get("D" + i).putValue(dataSeries.get(data));
                i++;
            }
        }
    }

    public Chart createChart(int upperLeftRow, int upperLeftColumn, int lowerRightRow, int lowerRightColumn) {
        ChartCollection charts = worksheet.getCharts();
        int chartIndex = charts.add(ChartType.COLUMN, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);
        return worksheet.getCharts().get(chartIndex);
    }

//    public void sortViaExcel(Map<String, Integer> dataSeries) {
//        DataSorter sorter = workbook.getDataSorter();
//        sorter.setKey1(1);
//        sorter.setKey2(3);
//        sorter.setOrder1(SortOrder.DESCENDING);
//        sorter.setOrder2(SortOrder.DESCENDING);
//
//        CellArea cellArea1 = CellArea.createCellArea("A1", "B" + dataSeries.size());
//        sorter.sort(worksheet.getCells(), cellArea1);
//
//        CellArea cellArea2 = CellArea.createCellArea("C1", "D" + dataSeries.size());
//        sorter.sort(worksheet.getCells(), cellArea2);
//    }

}
