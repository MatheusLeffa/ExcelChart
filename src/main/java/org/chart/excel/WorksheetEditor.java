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
            cells.get("A" + i).putValue(valorUnico);
            cells.get("B" + i).putValue(ocorrencias.get(valorUnico));
            i++;
        }
    }

    public Chart createChart() {
        ChartCollection charts = worksheet.getCharts();
        int chartIndex = charts.add(ChartType.COLUMN, 1, 6, 32, 25);
        return worksheet.getCharts().get(chartIndex);
    }

}
