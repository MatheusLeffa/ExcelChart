package org.chart.excel;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

public class ChartData extends WorksheetData {
    private Map<String, Integer> dataSeries;
    private final Integer maxLinhas = super.worksheet.getCells().getMaxDataRow();


    public ChartData(Workbook workbook, Worksheet worksheet, String coluna) {
        super(workbook, worksheet, coluna);
        this.dataSeries = new LinkedHashMap<>();
        setDataSeries();
        sortDataSeries();
    }

    private void setDataSeries() {
        Set<String> valoresUnicos = getValoresUnicos();

        for (String valorUnico : valoresUnicos) {
            int contagem = 0;

            for (int i = 2; i <= this.maxLinhas + 1; i++) {
                String valorCelula = super.worksheet.getCells().get(super.coluna + i).getStringValue();
                if (valorUnico.equals(valorCelula)) {
                    contagem++;
                }
            }
            dataSeries.put(valorUnico, contagem);
        }
    }

    private Set<String> getValoresUnicos() {
        Set<String> valoresUnicos = new HashSet<>();

        for (int i = 2; i <= maxLinhas + 1; i++) {
            String valorCelula = worksheet.getCells().get(coluna + i).getStringValue();
            valoresUnicos.add(valorCelula);
        }
        return valoresUnicos;
    }


    private void sortDataSeries() {
        this.dataSeries = dataSeries.entrySet()
                .stream()
                .sorted(Map.Entry.<String, Integer>comparingByValue().reversed())
                .collect(Collectors.toMap(
                        Map.Entry::getKey,
                        Map.Entry::getValue,
                        (oldValue, newValue) -> oldValue, LinkedHashMap::new));
    }

    public Map<String, Integer> getDataSeries() {
        return dataSeries;
    }

}
