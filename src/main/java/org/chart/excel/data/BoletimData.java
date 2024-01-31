package org.chart.excel.data;

import com.aspose.cells.CellsHelper;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import javax.swing.*;

public class BoletimData extends WorksheetData {
    String coluna;
    int colunaIndex;


    public BoletimData(Workbook workbook, Worksheet worksheet) {
        super(workbook, worksheet);
    }

    public void setColuna(String coluna) {

        this.coluna = null;

        for (int i = 0; i < worksheet.getCells().getMaxColumn(); i++) {
            String valorCelula = this.worksheet.getCells().get(0, i).getStringValue();
            if (valorCelula.equalsIgnoreCase(coluna)) {
                this.colunaIndex = i;
                this.coluna = CellsHelper.columnIndexToName(i);
            }
        }

        if (this.coluna == null) {
            JOptionPane.showMessageDialog(null, "A coluna '" + coluna + "' não foi localizada!", "Erro", JOptionPane.ERROR_MESSAGE);
            throw new RuntimeException("A coluna '" + coluna + "' não foi localizada!");
        }
    }

    public void deleteRowByStatus(String coluna, String[] statusValues) {
        setColuna(coluna);
        for (String status : statusValues) {
            for (int rowIndex = worksheet.getCells().getMaxDataRow(); rowIndex >= 1; rowIndex--) {
                String valorCelula = worksheet.getCells().get(rowIndex, colunaIndex).getStringValue();
                if (valorCelula.equalsIgnoreCase(status)) {
                    worksheet.getCells().deleteRow(rowIndex);
                }
            }
        }
    }

    public void deleteRowBySeguradora(String coluna, String seguradora) {
        setColuna(coluna);
        for (int rowIndex = worksheet.getCells().getMaxDataRow(); rowIndex >= 1; rowIndex--) {
            String valorCelula = worksheet.getCells().get(rowIndex, colunaIndex).getStringValue();
            if (!valorCelula.equalsIgnoreCase(seguradora)) {
                worksheet.getCells().deleteRow(rowIndex);
            }
        }
    }

    public void deleteColumns(String[] colunas) {
        for (String coluna : colunas) {
            setColuna(coluna);
            worksheet.getCells().deleteColumn(colunaIndex);
        }
    }
}
