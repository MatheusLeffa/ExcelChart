package org.chart.excel;

import org.chart.excel.gui.FormPrincipal;

public class Main {
    public static void main(String... args){

        //Form configuration
        FormPrincipal.setUpLookAndFeel();
        /* Create and display the form */
        java.awt.EventQueue.invokeLater(() -> new FormPrincipal().setVisible(true));
    }
}
