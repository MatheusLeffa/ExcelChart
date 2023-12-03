package org.chart.excel;

import com.aspose.cells.Chart;
import com.aspose.cells.Color;

public class ChartEditor {

    public static void chartFormatter(Chart chart, String title) {
        chart.getTitle().setText(title);
        chart.getChartArea().getArea().setForegroundColor(Color.getWhite());
        chart.getPlotArea().getArea().setForegroundColor(Color.getWhite());
        chart.getNSeries().get(0).getArea().setForegroundColor(Color.fromArgb(46, 117, 182));
        chart.setShowLegend(false);
        chart.getValueAxis().setMajorUnitScale(1);
        chart.setWallsAndGridlines2D(false);
    }
}
