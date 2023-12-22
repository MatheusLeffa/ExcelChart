package org.chart.excel;

import com.aspose.cells.Cells;
import com.aspose.cells.Worksheet;

import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;

public class ChartData {

    public static Map<String, Integer> getDataSeries(Worksheet worksheet, String coluna) {
        Map<String, Integer> dataSeries = new LinkedHashMap<>();
        int maxLinhas = worksheet.getCells().getMaxDataRow();

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
        return dataSeries;
    }

    private static Set<String> getValoresUnicosSerie(Worksheet worksheet, String coluna, int maxLinhas) {
        Set<String> valoresUnicos = new HashSet<>();

        for (int i = 2; i <= maxLinhas + 1; i++) {
            String valorCelula = worksheet.getCells().get(coluna + i).getStringValue();
            valoresUnicos.add(valorCelula);
        }
        return valoresUnicos;
    }
}
