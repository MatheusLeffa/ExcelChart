package org.chart.excel.worksheet.data;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class WorksheetData extends ExcelData {

    protected Worksheet worksheet;

    public WorksheetData(Workbook workbook, Worksheet worksheet) {
        super(workbook);
        this.worksheet = worksheet;
    }

}
