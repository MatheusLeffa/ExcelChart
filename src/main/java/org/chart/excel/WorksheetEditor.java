package org.chart.excel;

import com.aspose.cells.*;

import java.util.Map;

public class WorksheetEditor {

    Worksheet worksheet;

    public WorksheetEditor(Worksheet worksheet) {
        this.worksheet = worksheet;
    }

    public void addValues(Map<String, Integer> ocorrencias) {
        Cells cells = worksheet.getCells();
        int i = 1;
        for (String valorUnico : ocorrencias.keySet()) {
            if(cells.isBlankColumn(0)){
                cells.get("A" + i).putValue(valorUnico);
                cells.get("B" + i).putValue(ocorrencias.get(valorUnico));
            } else {
                cells.get("C" + i).putValue(valorUnico);
                cells.get("D" + i).putValue(ocorrencias.get(valorUnico));
            }
            i++;
        }
    }

    public Chart createChart(int upperLeftRow, int upperLeftColumn, int lowerRightRow, int lowerRightColumn) {
        ChartCollection charts = worksheet.getCharts();
        int chartIndex = charts.add(ChartType.COLUMN, upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn);
        return worksheet.getCharts().get(chartIndex);
    }

}
