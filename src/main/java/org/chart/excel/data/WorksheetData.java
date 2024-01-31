package org.chart.excel.data;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class WorksheetData{

    protected Workbook workbook;
    protected Worksheet worksheet;
    public WorksheetData(Workbook workbook, Worksheet worksheet) {
        this.workbook = workbook;
        this.worksheet = worksheet;
    }

}
