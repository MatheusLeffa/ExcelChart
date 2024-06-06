package org.chart.excel;

import com.formdev.flatlaf.FlatDarculaLaf;
import org.chart.excel.gui.MainMenu;

public class Main {

    //Inícializa o Menu Principal da aplicação com o tema FlatDarculaLaf.
    public static void main(String[] args) {
        FlatDarculaLaf.setup();
        MainMenu mainMenu = new MainMenu();
        mainMenu.setVisible(true);
    }
}
