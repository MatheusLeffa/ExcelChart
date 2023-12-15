package org.chart.excel;

import com.aspose.cells.Cells;
import com.aspose.cells.Worksheet;

import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;

public class ChartData {

    private final Map<String, Integer> dataSeries = new LinkedHashMap<>();

    public Map<String, Integer> getDataSeries(Worksheet worksheet, String coluna) {

        int indexColuna = worksheet.getCells().get(coluna + 1).getColumn();
        int maxLinhas = getMaxLinhas(worksheet, indexColuna);

        Set<String> valoresUnicos = getValoresUnicosSerie(worksheet, coluna, maxLinhas);

        for (String valorUnico : valoresUnicos) {
            int contagem = 0;

            for (int i = 2; i <= maxLinhas + 1; i++) {
                String valorCelula = worksheet.getCells().get(coluna + i).getStringValue();
                if (valorUnico.equals(valorCelula)) {
                    contagem++;
                }
            }
            dataSeries.put(valorUnico, contagem);
        }
        return this.dataSeries;
    }

    private static Set<String> getValoresUnicosSerie(Worksheet worksheet, String coluna, int maxLinhas) {
        Set<String> valoresUnicos = new HashSet<>();

        for (int i = 2; i <= maxLinhas + 1; i++) {
            String valorCelula = worksheet.getCells().get(coluna + i).getStringValue();
            valoresUnicos.add(valorCelula);
        }
        return valoresUnicos;
    }

    private static int getMaxLinhas(Worksheet worksheet, int columnIndex) {
        int maxRow = 0;
        Cells cells = worksheet.getCells();

        for (int row = 0; row < cells.getMaxDataRow() + 1; row++) {
            if (!cells.get(row, columnIndex).getStringValue().isEmpty()) {
                maxRow = row;
            }
        }
        return maxRow;
    }
}
