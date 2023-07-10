package org.example;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;
import java.util.HashSet;
import java.util.Set;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelFileReaderGUI extends JFrame {

        private JLabel lblExcelFile;
        private JButton btnBrowse;
        private JButton btnImport;
        private JTextArea txtLog;
        private JFileChooser fileChooser;

        public ExcelFileReaderGUI() {
            // Set up the main JFrame
            setTitle("Excel Import");
            setSize(400, 300);
            setDefaultCloseOperation(EXIT_ON_CLOSE);
            setLayout(new BorderLayout());

            // Create and configure components
            lblExcelFile = new JLabel("Excel File:");
            btnBrowse = new JButton("Browse");
            btnImport = new JButton("Upload");
            txtLog = new JTextArea();
            fileChooser = new JFileChooser();

            // Add components to the JFrame
            JPanel topPanel = new JPanel();
            topPanel.add(lblExcelFile);
            topPanel.add(btnBrowse);
            add(topPanel, BorderLayout.NORTH);
            add(new JScrollPane(txtLog), BorderLayout.CENTER);
            add(btnImport, BorderLayout.SOUTH);

            // Browse button action listener
            btnBrowse.addActionListener(new ActionListener() {
                public void actionPerformed(ActionEvent e) {
                    int returnVal = fileChooser.showOpenDialog(ExcelFileReaderGUI.this);
                    if (returnVal == JFileChooser.APPROVE_OPTION) {
                        File file = fileChooser.getSelectedFile();
                        lblExcelFile.setText("Excel File: " + file.getName());
                    }
                }
            });

            // Import button action listener
            btnImport.addActionListener(new ActionListener() {
                public void actionPerformed(ActionEvent e) {

                    File selectedFile = fileChooser.getSelectedFile();
                    if (selectedFile != null) {
                        importData(selectedFile);
                    } else {
                        JOptionPane.showMessageDialog(ExcelFileReaderGUI.this, "Please select an Excel file.");
                    }
                }
            });
        }

    private void importData(File file) {
        try {
            // Read the Excel file
            FileInputStream fis = new FileInputStream(file);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);

            // Establish database connection
            Connection connection = DriverManager.getConnection("jdbc:mysql://localhost:3306/excel_test", "root", "lynn471997");

            // Prepare the SQL statement for checking duplicates
            String sqlCheckDuplicate = "SELECT COUNT(*) FROM employee WHERE employee_name = ?";

          /*  String sqlCheckDuplicate = "SELECT id FROM employee WHERE id = ?";*/

            // Prepare the SQL statement for inserting data
            String sqlInsertData = "INSERT INTO employee (id, employee_name, position, department, salary, joined_date) VALUES (?, ?, ?, ?, ?, ?)";

            try (PreparedStatement statementCheckDuplicate = connection.prepareStatement(sqlCheckDuplicate);
                 PreparedStatement statementInsertData = connection.prepareStatement(sqlInsertData)) {

                //Create a set to store unique employee names
                Set<String> uniqueNames = new HashSet<>();
                /* Set<Integer> uniqueIDs = new HashSet<>();*/

                // Iterate through rows and insert data into the database
                int totalDuplicates = 0;
                int totalSkipped = 0;

                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    int id = (int) row.getCell(0).getNumericCellValue();
                    String name = row.getCell(1).getStringCellValue();
                    String position = row.getCell(2).getStringCellValue();
                    String department = row.getCell(3).getStringCellValue();
                    double salary = row.getCell(4).getNumericCellValue();
                    java.util.Date joinedDate = row.getCell(5).getDateCellValue();


                    if (!uniqueNames.add(name)){
                        //Name is a duplicate, increment the duplicated count
                        totalDuplicates++;
                        totalSkipped++;
                        continue;
                    }

                    // Name is not a duplicate, insert the data into the database

                    statementInsertData.setInt(1, id);
                    statementInsertData.setString(2, name);
                    statementInsertData.setString(3, position);
                    statementInsertData.setString(4, department);
                    statementInsertData.setDouble(5, salary);
                    statementInsertData.setDate(6, new java.sql.Date(joinedDate.getTime()));

                    statementInsertData.executeUpdate();
                }

                // Close the resources
                statementCheckDuplicate.close();
                statementInsertData.close();
                connection.close();
                workbook.close();
                fis.close();

                txtLog.setText("Data imported successfully. Total duplicates found: " + totalDuplicates + " and Total skipped found: " + totalSkipped + ".");
            } catch (IOException | SQLException ex) {
                txtLog.setText("Error occurred: " + ex.getMessage());
            }
        }catch (IOException | SQLException ex) {
            txtLog.setText("Error occurred: " + ex.getMessage());
        }
    }

        public static void main(String[] args) {
            SwingUtilities.invokeLater(new Runnable() {
                public void run() {
                    new ExcelFileReaderGUI().setVisible(true);
                }
            });
        }
}