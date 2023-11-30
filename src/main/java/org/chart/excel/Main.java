package org.chart.excel;

import com.aspose.cells.*;

import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Set;

public class Main {
    public static void main(String[] args) throws Exception {

        String filePath = FileSelector.getFilePath();
        Workbook workbook = new Workbook(filePath);
        WorksheetCollection worksheetCollection = workbook.getWorksheets();

        Worksheet worksheet0 = worksheetCollection.get(0);
        Worksheet worksheet1;

        if (workbook.getWorksheets().getCount() > 1) {
            worksheet1 = worksheetCollection.get(1);
        } else {
            worksheet1 = worksheetCollection.add("Gráficos");
        }

        // Criar um conjunto para armazenar valores únicos
        Set<String> valoresUnicos = new HashSet<>();

        int maxLinhas = worksheet0.getCells().getMaxDataRow() + 1;

        for (int i = 2; i <= maxLinhas; i++) {
            String valorCelula = worksheet0.getCells().get("A" + i).getStringValue();
            valoresUnicos.add(valorCelula);
        }

        // Contar as ocorrências do valor específico
        Map<String, Integer> ocorrencias = new LinkedHashMap<>();

        for (String valorUnico : valoresUnicos) {
            int contagem = 0;

            for (int i = 2; i <= maxLinhas; i++) {

                String valorCelula = worksheet0.getCells().get("A" + i).getStringValue();

                if (valorUnico.equals(valorCelula)) {
                    contagem++;
                }
            }
            ocorrencias.put(valorUnico, contagem);
        }

        // Adicionar dados do gráfico
        Cells cells = worksheet1.getCells();
        int i = 1;
        for (String valorUnico : ocorrencias.keySet()) {
            cells.get("A" + i).putValue(valorUnico);
            cells.get("B" + i).putValue(ocorrencias.get(valorUnico));
            i++;
        }

        // Criar gráfico de colunas
        ChartCollection charts = worksheet1.getCharts();
        int chartIndex = charts.add(ChartType.COLUMN, 1, 6, 32, 25);
        Chart chart = worksheet1.getCharts().get(chartIndex);

        // Configurar dados do gráfico
        chart.setChartDataRange("A1:B" + ocorrencias.size(), true);

        // Formatação gráfico
        chart.getTitle().setText("Incidentes");
        chart.getChartArea().getArea().setForegroundColor(Color.getWhite());
        chart.getPlotArea().getArea().setForegroundColor(Color.getWhite());
        chart.getNSeries().get(0).getArea().setForegroundColor(Color.fromArgb(46, 117, 182));
        chart.setShowLegend(false);
        chart.getValueAxis().setMajorUnitScale(1);

        // Salvar o arquivo
        try {
            workbook.save("Test_out.xlsx", SaveFormat.XLSX);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}