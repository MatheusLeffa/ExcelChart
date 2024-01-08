package org.chart.excel;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;

public class FileSelector {

    public static String getFilePath() {
        JFileChooser file = selectFile();
        return file.getSelectedFile().getPath();
    }

    private static JFileChooser selectFile() {
        JFileChooser chooser = new JFileChooser();
        try {
            FileNameExtensionFilter filter = new FileNameExtensionFilter("xlsx", "xlsx");
            chooser.setFileFilter(filter);
            int returnVal = chooser.showOpenDialog(null);

            if (returnVal == JFileChooser.APPROVE_OPTION) {
                return chooser;
            } else {
                throw new RuntimeException("The file selection was canceled.");
            }
        } catch (Exception e) {
            throw new RuntimeException();
        }
    }

}
