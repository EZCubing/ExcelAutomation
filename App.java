package com.zeak.ExcelAUTO2;

import javax.swing.*;
import java.awt.*; 
import java.awt.event.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {
    public static void main(String[] args) { 
        // Create the window
        JFrame frame = new JFrame("Simple GUI Example");
        frame.setSize(400, 400); 
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setLayout(new BoxLayout(frame.getContentPane(), BoxLayout.Y_AXIS));

        // Create a text area for feedback
        JTextArea textArea = new JTextArea("");
        textArea.setAlignmentX(Component.CENTER_ALIGNMENT);
        textArea.setEditable(false);
        textArea.setLineWrap(true); 
        textArea.setWrapStyleWord(true);
        frame.add(new JScrollPane(textArea));

        frame.add(Box.createVerticalStrut(20));

        // Create and customize a button
        JButton button = new JButton("Choose Files");
        button.setPreferredSize(new Dimension(150, 50)); 
        button.setAlignmentX(Component.CENTER_ALIGNMENT);
        frame.add(button);

        // Add action listener to the button
        button.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                fileChooser.setMultiSelectionEnabled(true);
                int result = fileChooser.showOpenDialog(frame);
                if (result == JFileChooser.APPROVE_OPTION) {
                    File[] selectedFiles = fileChooser.getSelectedFiles();
                    try {
                        // Open the existing Excel file
                    	// Maybe have them choose the file with the GUI and make it where they do not have to open java to type in the file
                        File excelFile = new File("/Users/ezcubing/L3Harris_Java_Tool/ExcelPractice.xlsx"); // Your existing file 
                        if (!excelFile.exists()) {
                            textArea.setText("Error: 'input.xlsx' not found in project directory!");
                            return; 
                        }
                        FileInputStream fis = new FileInputStream(excelFile);
                        XSSFWorkbook workbook = new XSSFWorkbook(fis);

                        // Check if there are enough sheets
                        if (selectedFiles.length > workbook.getNumberOfSheets()) {
                            textArea.setText("Error: Not enough sheets (" + workbook.getNumberOfSheets() + 
                                            ") for " + selectedFiles.length + " files!");
                            fis.close();
                            workbook.close();
                            return;
                        }

                        // Process each selected file
                        for (int i = 0; i < selectedFiles.length; i++) {
                            File file = selectedFiles[i];
                            String content = new String(Files.readAllBytes(file.toPath()));

                            // Get the corresponding sheet (Sheet1 = 0, Sheet2 = 1, etc.)
                            Sheet sheet = workbook.getSheetAt(i + 1);  

                            // Define where to paste (e.g., row 2, column 1 = B3) 
                            int rowNum = 14; // Row 3 in Excel (0-based)
                            int colNum = 0; // Column B in Excel (0-based) 

                            // Get or create the row
                            Row row = sheet.getRow(rowNum);
                            if (row == null) {
                                row = sheet.createRow(rowNum);
                            }

                            // Get or create the cell and paste content
                            Cell cell = row.getCell(colNum);
                            if (cell == null) {
                                cell = row.createCell(colNum);
                            }
                            cell.setCellValue(content);
                        } 

                        // Save changes back to the same file
                        fis.close();
                        FileOutputStream fos = new FileOutputStream(excelFile);
                        workbook.write(fos);
                        fos.close();
                        workbook.close(); 

                        textArea.setText("Updated " + excelFile + " with content from " + selectedFiles.length + " files!");
                    } catch (IOException ex) {
                        textArea.setText("Error: " + ex.getMessage());
                    }
                } else {
                    textArea.setText("No files selected.");
                }
            }
        });

        // Make the window visible
        frame.setVisible(true);
    }
}
