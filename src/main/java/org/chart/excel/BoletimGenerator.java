package org.chart.excel;

import com.aspose.cells.*;
import org.chart.excel.utils.ChartFormatter;
import org.chart.excel.data.BoletimData;
import org.chart.excel.data.ChartData;

import javax.swing.*;
import java.io.File;
import java.io.IOException;

public class BoletimGenerator {
    private String filePath;
    private String seguradora;
    private final LoadOptions loadOptions = setLoadOptions();

    public void setFilePath(String filePath) {
        this.filePath = filePath;
    }

    public void setSeguradora(String seguradora) {
        this.seguradora = seguradora;
    }

    public void generateBoletim() {

        Workbook workbook = setWorkbook(filePath, loadOptions);
        WorksheetCollection worksheetCollection = workbook.getWorksheets();
        Worksheet tabelaBoletim = worksheetCollection.get(0);
        Worksheet tabelaGraficos = worksheetCollection.add("Gráficos");

        filterBoletim(workbook, tabelaBoletim, seguradora);
        chartCreator(workbook, tabelaBoletim, tabelaGraficos);
        saveWorkbookFile(workbook);
        openWorkbookFile();
    }

    private LoadOptions setLoadOptions() {
        LoadOptions loadOptions = new LoadOptions(LoadFormat.XLSX);
        loadOptions.setLoadFilter(new LoadFilter(LoadDataFilterOptions.CELL_DATA));
        return loadOptions;
    }
    private Workbook setWorkbook(String filePath, LoadOptions loadOptions) {
        try {
            return new Workbook(filePath, loadOptions);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
    private void chartCreator(Workbook workbook, Worksheet tabelaBoletimOrigem, Worksheet tabelaGraficos) {
        ChartData colunaSistema = new ChartData(workbook, tabelaBoletimOrigem, "Produto/Sistema");
        ChartData colunaStatus = new ChartData(workbook, tabelaBoletimOrigem, "Status");

        if(colunaSistema.getDataSeries().size() == 0) {
            JOptionPane.showMessageDialog(null, "Não possuem incidentes abertos para esta seguradora.", "Erro", JOptionPane.ERROR_MESSAGE);
            throw new RuntimeException("Não possuem incidentes abertos para esta seguradora.");
        }

        ChartEditor EditorTabelaGraficos = new ChartEditor(tabelaGraficos);
        EditorTabelaGraficos.addValues(colunaSistema.getDataSeries());
        EditorTabelaGraficos.addValues(colunaStatus.getDataSeries());

        Chart graficoSistema = EditorTabelaGraficos.createChart(1, 4, 20, 20);
        Chart graficoStatus = EditorTabelaGraficos.createChart(22, 4, 38, 18);

        graficoSistema.setChartDataRange("A1:B" + (colunaSistema.getColumnSize()), true);
        graficoStatus.setChartDataRange("C1:D" + (colunaStatus.getColumnSize()), true);

        ChartFormatter.chartFormatter(graficoSistema, "Sistema");
        ChartFormatter.chartFormatter(graficoStatus, "Status");
    }
    private void saveWorkbookFile(Workbook workbook) {
        try {
            workbook.save("Boletim_out.xlsx", SaveFormat.XLSX);
        } catch (Exception e) {
            JOptionPane.showMessageDialog(null, "O arquivo 'Boletim_out.xlsx' está aberto em outro programa!", "Erro", JOptionPane.ERROR_MESSAGE);
            throw new RuntimeException(e);
        }
    }
    private void openWorkbookFile(){
        try {
            java.awt.Desktop.getDesktop().open(new File("Boletim_out.xlsx"));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
    private void filterBoletim(Workbook workbook, Worksheet tabelaBoletimOrigem, String seguradora) {
        String[] statusToRemove = {"Finalizado", "Solucionado", "Fechado"};
        String[] colunasToRemove = {"Data conclusão", "Hora conclusão", "Criticidade", "Crítico / Impactante"};

        BoletimData boletimData = new BoletimData(workbook, tabelaBoletimOrigem);
        boletimData.deleteRowByStatus("Status", statusToRemove);
        boletimData.deleteRowBySeguradora("Responsável pelo Ajuste", seguradora);
        boletimData.deleteColumns(colunasToRemove);
    }

}