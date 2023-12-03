package org.chart.excel;

import com.aspose.cells.Worksheet;

import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;

public class ChartData {

    private static final Map<String, Integer> ocorrencias = new LinkedHashMap<>();

    public static Map<String, Integer> getDataSeries(Worksheet worksheet, String coluna) {

        Set<String> valoresUnicos = getValoresUnicosSerie(worksheet, coluna);
        int maxLinhas = getMaxLinhas(worksheet);

        for (String valorUnico : valoresUnicos) {
            int contagem = 0;

            for (int i = 2; i <= maxLinhas; i++) {
                String valorCelula = worksheet.getCells().get(coluna + i).getStringValue();
                if (valorUnico.equals(valorCelula)) {
                    contagem++;
                }
            }
            ocorrencias.put(valorUnico, contagem);
        }
        return ocorrencias;
    }

    private static Set<String> getValoresUnicosSerie(Worksheet worksheet, String coluna) {
        Set<String> valoresUnicos = new HashSet<>();

        int maxLinhas = getMaxLinhas(worksheet);

        for (int i = 2; i <= maxLinhas; i++) {
            String valorCelula = worksheet.getCells().get(coluna + i).getStringValue();
            valoresUnicos.add(valorCelula);
        }
        return valoresUnicos;
    }

    private static int getMaxLinhas(Worksheet worksheet) {
        return worksheet.getCells().getMaxDataRow() + 1;
    }
}
