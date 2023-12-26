package org.chart.excel;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public abstract class Data {

    Workbook workbook;
    Worksheet worksheet;
    String coluna;


    public Data(Workbook workbook, Worksheet worksheet, String coluna) {
        this.workbook = workbook;
        this.worksheet = worksheet;
        this.coluna = coluna;
    }
}
