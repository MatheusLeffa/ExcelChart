package org.chart.excel;

import com.aspose.cells.*;
import com.aspose.cells.Color;

import java.awt.*;
import java.io.IOException;
import java.util.*;

public class Main {
    public static void main(String[] args) throws Exception {

        String filePath = FileSelector.getFilePath();
        Workbook workbook = new Workbook(filePath);

        Worksheet worksheet0 = workbook.getWorksheets().get(0);
        Worksheet worksheet1 = workbook.getWorksheets().get(1);
        worksheet1.setName("Gráficos");

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
        int chartIndex = charts.add(ChartType.COLUMN, 1, 5, 15, 15);
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