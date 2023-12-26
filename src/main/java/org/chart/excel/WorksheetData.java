package org.chart.excel;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public abstract class WorksheetData {

    Workbook workbook;
    Worksheet worksheet;
    String coluna;


    public WorksheetData(Workbook workbook, Worksheet worksheet, String coluna) {
        this.workbook = workbook;
        this.worksheet = worksheet;
        this.coluna = coluna;
    }
}
