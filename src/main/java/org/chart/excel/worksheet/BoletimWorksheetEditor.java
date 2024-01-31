package org.chart.excel.worksheet;

import com.aspose.cells.Cells;
import com.aspose.cells.Worksheet;

public class BoletimWorksheetEditor {

    private Worksheet worksheet;
    private Cells cells;

    public BoletimWorksheetEditor(Worksheet worksheet) {
        this.worksheet = worksheet;
        this.cells = worksheet.getCells();
    }
}
