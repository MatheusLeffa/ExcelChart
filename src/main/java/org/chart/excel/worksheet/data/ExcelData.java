package org.chart.excel.worksheet.data;

import com.aspose.cells.Workbook;

public abstract class ExcelData {

    protected Workbook workbook;

    public ExcelData(Workbook workbook) {
        this.workbook = workbook;
    }

}
