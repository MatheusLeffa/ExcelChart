package org.chart.excel.worksheet;

import com.aspose.cells.CellsHelper;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public abstract class WorksheetData {

    protected Workbook workbook;
    protected Worksheet worksheet;
    protected String coluna;


    public WorksheetData(Workbook workbook, Worksheet worksheet, String coluna) {
        this.workbook = workbook;
        this.worksheet = worksheet;
        this.coluna = setColuna(coluna);
    }

    private String setColuna(String coluna) {

        for (int i = 0; i < worksheet.getCells().getMaxColumn(); i++) {
            String valorCelula = this.worksheet.getCells().get(0, i).getStringValue();
            if (valorCelula.equals(coluna)) {
                return CellsHelper.columnIndexToName(i);
            }
        }
        throw new RuntimeException("A coluna '" + coluna + "' não foi localizada!");
    }
}
