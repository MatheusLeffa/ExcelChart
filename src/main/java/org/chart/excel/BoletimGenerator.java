package org.chart.excel;

import com.aspose.cells.*;
import org.chart.excel.worksheet.WorksheetEditor;
import org.chart.excel.worksheet.chart.ChartData;
import org.chart.excel.worksheet.chart.ChartFormatter;

import javax.swing.*;
import java.io.File;

public class BoletimGenerator {
    
    String filePath;
    
    
    public BoletimGenerator(){
    }
    
    public void setFilePath(String filePath){
        this.filePath = filePath;
    }
    
    public void generate() throws Exception {

        Workbook workbook = new Workbook(filePath);

        WorksheetCollection worksheetCollection = workbook.getWorksheets();
        Worksheet tabelaBoletim = worksheetCollection.get(0);
        Worksheet tabelaGraficos = worksheetCollection.add("Gráficos");

        ChartData colunaSistema = new ChartData(workbook, tabelaBoletim, "Produto/Sistema");
        ChartData colunaStatus = new ChartData(workbook, tabelaBoletim, "Status");

        WorksheetEditor EditorTabelaGraficos = new WorksheetEditor(tabelaGraficos);
        EditorTabelaGraficos.addValues(colunaSistema.getDataSeries());
        EditorTabelaGraficos.addValues(colunaStatus.getDataSeries());

        Chart graficoSistema = EditorTabelaGraficos.createChart(1, 4, 20, 20);
        Chart graficoStatus = EditorTabelaGraficos.createChart(22, 4, 38, 18);

        graficoSistema.setChartDataRange("A1:B" + (colunaSistema.getColumnSize()), true);
        graficoStatus.setChartDataRange("C1:D" + (colunaStatus.getColumnSize()), true);

        ChartFormatter.chartFormatter(graficoSistema, "Sistema");
        ChartFormatter.chartFormatter(graficoStatus, "Status");

        try {
            workbook.save("Boletim_out.xlsx", SaveFormat.XLSX);
        } catch (Exception e) {
            JOptionPane.showMessageDialog(null,"O arquivo 'Boletim_out.xlsx' está aberto em outro programa!","Erro", JOptionPane.ERROR_MESSAGE);
            throw new RuntimeException(e);
        }
        java.awt.Desktop.getDesktop().open(new File("Boletim_out.xlsx"));
    }
}