package org.chart.excel.utils;

import com.aspose.cells.Axis;
import com.aspose.cells.Chart;
import com.aspose.cells.Color;
import com.aspose.cells.Line;

public class ChartFormatter {

    public static void chartFormatter(Chart chart, String title) {
        chart.getTitle().setText(title);
        chart.getTitle().getFont().setSize(14);
        chart.getTitle().getFont().setBold(true);
        chart.getChartArea().getArea().setForegroundColor(Color.getWhite());
        chart.getPlotArea().getArea().setForegroundColor(Color.getWhite());
        chart.getNSeries().get(0).getArea().setForegroundColor(Color.fromArgb(46, 117, 182));
        chart.setShowLegend(false);
        chart.getValueAxis().setVisible(false);
        chart.getPlotArea().getBorder().setVisible(false);
        chart.getNSeries().get(0).getDataLabels().setShowValue(true);
        disableGridLines(chart);
    }

    private static void disableGridLines(Chart chart) {
        Axis valueAxis = chart.getValueAxis();
        Line majorGridLines = valueAxis.getMajorGridLines();
        majorGridLines.setVisible(false);
    }

}
