package org.chart.excel.worksheet.data;

import com.aspose.cells.CellsHelper;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import javax.swing.*;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

public class ChartData extends WorksheetData {
    private Map<String, Integer> dataSeries;
    private final String coluna;
    private final Integer maxLinhas;

    public ChartData(Workbook workbook, Worksheet worksheet, String coluna) {
        super(workbook, worksheet);
        this.dataSeries = new LinkedHashMap<>();
        this.coluna = setColuna(coluna);
        this.maxLinhas = worksheet.getCells().getMaxDataRow();
        setDataSeries();
        sortDataSeries();
    }

    private String setColuna(String coluna) {

        for (int i = 0; i < worksheet.getCells().getMaxColumn(); i++) {
            String valorCelula = this.worksheet.getCells().get(0, i).getStringValue();
            if (valorCelula.equalsIgnoreCase(coluna)) {
                return CellsHelper.columnIndexToName(i);
            }
        }
        JOptionPane.showMessageDialog(null, "A coluna '" + coluna + "' não foi localizada!", "Erro", JOptionPane.ERROR_MESSAGE);
        throw new RuntimeException("A coluna '" + coluna + "' não foi localizada!");
    }

    private void setDataSeries() {
        Set<String> valoresUnicos = getValoresUnicos();

        for (String valorUnico : valoresUnicos) {
            int contagem = 0;

            for (int i = 2; i <= this.maxLinhas + 1; i++) {
                String valorCelula = super.worksheet.getCells().get(coluna + i).getStringValue().toUpperCase();
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
            String valorCelula = worksheet.getCells().get(coluna + i).getStringValue().toUpperCase();
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

    public int getColumnSize() {
        return dataSeries.size() + 1;
    }

}
