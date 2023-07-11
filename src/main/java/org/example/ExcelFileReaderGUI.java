package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashSet;
import java.util.Set;

public class ExcelFileReaderGUI extends JFrame {

    private final JLabel lblExcelFile;
    private final JTextArea txtLog;
    private final JFileChooser fileChooser;

    public ExcelFileReaderGUI() {
        // Set up the main JFrame
        setTitle("Excel Import");
        setSize(500, 400);
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setLayout(new BorderLayout());

        // Create and configure components
        lblExcelFile = new JLabel("Excel File:");
        JButton btnBrowse = new JButton("Browse");
        JButton btnImport = new JButton("Upload");
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
        btnBrowse.addActionListener(e -> {
            int returnVal = fileChooser.showOpenDialog(ExcelFileReaderGUI.this);
            if (returnVal == JFileChooser.APPROVE_OPTION) {
                File file = fileChooser.getSelectedFile();
                lblExcelFile.setText("Excel File: " + file.getName());
            }
        });

        // Import button action listener
        btnImport.addActionListener(e -> {

            File selectedFile = fileChooser.getSelectedFile();
            if (selectedFile != null) {
                importData(selectedFile);
            } else {
                JOptionPane.showMessageDialog(ExcelFileReaderGUI.this, "Please select an Excel file.");
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

            // Prepare the SQL statement for inserting data
            String sqlInsertData = "INSERT INTO employee (id, employee_name, position, department, salary, joined_date) VALUES (?, ?, ?, ?, ?, ?)";

            // Unique Name and Id
            Set<Integer> existingIds = new HashSet<>();
            Set<String> existingNames = new HashSet<>();

            ResultSet rs = connection.createStatement().executeQuery("SELECT id,employee_name FROM employee");
            while (rs.next()) {
                existingIds.add(rs.getInt("id"));
                existingNames.add(rs.getString("employee_name"));
            }

            try (PreparedStatement statementInsertData = connection.prepareStatement(sqlInsertData)) {

                // Iterate through rows and insert data into the database
                int totalDuplicates = 0;
                int totalSkipped = 0;

                for (Row row : sheet) {
                    Cell idCell = row.getCell(0);
                    Cell nameCell = row.getCell(1);
                    Cell positionCell = row.getCell(2);
                    Cell departmentCell = row.getCell(3);
                    Cell salaryCell = row.getCell(4);
                    Cell joinedDateCell = row.getCell(5);

                    // Handle cell types correctly
                    int id;
                    String name;
                    String position;
                    String department;
                    double salary;
                    java.sql.Date joinedDate;

                    if (idCell.getCellType() == CellType.NUMERIC) {
                        id = (int) idCell.getNumericCellValue();
                    } else {
                        continue; // Skip this row if ID is not numeric
                    }

                    if (existingIds.contains(id)) {
                        totalDuplicates++;
                        totalSkipped++;
                        continue; // Skip this row if ID already exists in the database
                    }

                    if (nameCell.getCellType() == CellType.STRING) {
                        name = nameCell.getStringCellValue();
                    } else {
                        continue; // Skip this row if name is not a string
                    }

                    if (existingNames.contains(name)) {
                        totalDuplicates++;
                        totalSkipped++;
                        continue; // Skip this row if name already exists in the database
                    }


                    if (positionCell.getCellType() == CellType.STRING) {
                        position = positionCell.getStringCellValue();
                    } else {
                        continue; // Skip this row if position is not a string
                    }

                    if (departmentCell.getCellType() == CellType.STRING) {
                        department = departmentCell.getStringCellValue();
                    } else {
                        continue; // Skip this row if department is not a string
                    }

                    if (salaryCell.getCellType() == CellType.NUMERIC) {
                        salary = salaryCell.getNumericCellValue();
                    } else {
                        continue; // Skip this row if salary is not numeric
                    }

                    if (joinedDateCell.getCellType() == CellType.NUMERIC) {
                        // Cell contains a numeric value
                        Date dateValue = joinedDateCell.getDateCellValue();
                        joinedDate = new java.sql.Date(dateValue.getTime());
                    } else if (joinedDateCell.getCellType() == CellType.STRING) {
                        // Cell contains a string value
                        String dateString = joinedDateCell.getStringCellValue();
                        // Parse the string into a date using the desired format
                        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                        try {
                            java.util.Date parsedDate = dateFormat.parse(dateString);
                            joinedDate = new java.sql.Date(parsedDate.getTime());
                        } catch (ParseException ex) {
                            // Invalid date format, handle the error as needed
                            continue; // Skip this row if the date format is invalid
                        }
                    } else {
                        // Cell is of an unsupported type, handle the error as needed
                        continue; // Skip this row if the cell type is unsupported
                    }

                    // Insert the data into the database
                    statementInsertData.setInt(1, id);
                    statementInsertData.setString(2, name);
                   statementInsertData.setString(3,position);
                   statementInsertData.setString(4,department);
                   statementInsertData.setDouble(5,salary);
                   statementInsertData.setDate(6, (java.sql.Date) joinedDate);

                    statementInsertData.executeUpdate();
                    existingIds.add(id);
                    existingNames.add(name);
                }

                // Close the resources
                statementInsertData.close();
                connection.close();
                workbook.close();
                fis.close();

                txtLog.setText("Data imported successfully. Total duplicates found: " + totalDuplicates + " and Total skipped found: " + totalSkipped + ".");
            } catch (IOException | SQLException ex) {
                txtLog.setText("Error occurred: " + ex.getMessage());
            }
        } catch (IOException | SQLException ex) {
            txtLog.setText("Error occurred: " + ex.getMessage());
        }
    }
            public static void main (String[]args){
                SwingUtilities.invokeLater(() -> new ExcelFileReaderGUI().setVisible(true));
            }
}