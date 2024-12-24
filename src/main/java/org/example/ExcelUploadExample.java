package org.example;

import com.itextpdf.io.font.PdfEncodings;
import com.itextpdf.io.font.constants.StandardFonts;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.AreaBreak;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.properties.AreaBreakType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.itextpdf.kernel.font.*;
import com.itextpdf.kernel.pdf.*;

import javax.swing.*;
import java.awt.*;
import java.awt.Desktop;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.URL;
import java.util.Iterator;

import static com.itextpdf.kernel.pdf.PdfName.BaseFont;

public class ExcelUploadExample {

    public static void main(String[] args) {
        // สร้างกรอบ (Frame)
        JFrame frame = new JFrame("ตารางสินค้า");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(800, 600);
        frame.setLocationRelativeTo(null);

        // กำหนดฟอนต์ให้รองรับภาษาไทย
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            Font thaiFont = new Font("Tahoma", Font.PLAIN, 14); // หรือฟอนต์อื่นๆ ที่รองรับภาษาไทย เช่น TH Sarabun
            UIManager.put("Label.font", thaiFont);
            UIManager.put("Table.font", thaiFont);
        } catch (Exception e) {
            e.printStackTrace();
        }

        // สร้างปุ่มให้เลือกไฟล์ Excel
        JButton uploadButton = new JButton("อัปโหลดไฟล์ Excel");
        uploadButton.addActionListener(e -> {
            // เลือกไฟล์ Excel
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setDialogTitle("เลือกไฟล์ Excel");
            int result = fileChooser.showOpenDialog(frame);
            if (result == JFileChooser.APPROVE_OPTION) {
                File selectedFile = fileChooser.getSelectedFile();
                try {
                    // อ่านข้อมูลจากไฟล์ Excel
                    Object[][] data = readExcelFile(selectedFile);
                    // สร้างตารางแสดงข้อมูล
                    String[] columnNames = {
                            "เลือก", "ยาว (ฟุต)", "กว้าง (นิ้ว)", "หนา (นิ้ว)", "จำนวน/แผ่น",
                            "รหัสสินค้า", "วันที่", "ชื่อสินค้า", "เกรด", "Location", "Barcode"
                    };

//                    String[] columnNames = {
//                            "01", "02", "03", "04", "05", "06", "07",
//
//                    };

                    // สร้าง DefaultTableModel และ JCheckBox ในคอลัมน์ "เลือก"
                    Object[][] dataWithCheckbox = new Object[data.length][columnNames.length];
                    for (int i = 0; i < data.length; i++) {
                        dataWithCheckbox[i][0] = false; // คอลัมน์ "เลือก" สำหรับ JCheckBox
                        for (int j = 1; j < columnNames.length; j++) {
                            dataWithCheckbox[i][j] = data[i][j - 1]; // ข้อมูลจริงจาก Excel
                        }
                    }

                    JTable table = new JTable(dataWithCheckbox, columnNames) {
                        @Override
                        public Class<?> getColumnClass(int columnIndex) {
                            if (columnIndex == 0) {
                                return Boolean.class; // ให้คอลัมน์แรกเป็น JCheckBox
                            }
                            return super.getColumnClass(columnIndex);
                        }
                    };
                    table.setFont(new Font("Tahoma", Font.PLAIN, 12));  // ใช้ฟอนต์ที่รองรับภาษาไทย
                    table.setFillsViewportHeight(true);
                    table.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);  // การเลือกหลายแถว
                    JScrollPane scrollPane = new JScrollPane(table);

                    // สร้าง JTextArea สำหรับแสดงรายละเอียดของแถวที่เลือก
                    JTextArea detailArea = new JTextArea(10, 40);
                    detailArea.setFont(new Font("Tahoma", Font.PLAIN, 14));  // ฟอนต์ที่รองรับภาษาไทย
                    detailArea.setEditable(false);

                    // ฟังการเลือกแถวใน JTable
                    table.getSelectionModel().addListSelectionListener(e1 -> {
                        if (!e1.getValueIsAdjusting()) {
                            int[] selectedRows = table.getSelectedRows();
                            StringBuilder details = new StringBuilder();
                            for (int row : selectedRows) {
                                details.append("แถวที่ ").append(row + 1).append(":\n");
                                for (int i = 0; i < columnNames.length; i++) {
                                    details.append(columnNames[i]).append(": ")
                                            .append(table.getValueAt(row, i)).append("\n");
                                }
                                details.append("\n");
                            }
                            detailArea.setText(details.toString()); // แสดงข้อมูลของแถวที่เลือกทั้งหมด
                        }
                    });

                    // ปุ่มสำหรับเลือกทั้งหมด
                    JButton selectAllButton = new JButton("เลือกทั้งหมด");
                    selectAllButton.addActionListener((ActionEvent e2) -> {
                        for (int i = 0; i < table.getRowCount(); i++) {
                            table.setValueAt(true, i, 0); // ตั้งค่าเครื่องหมายถูกในทุกแถว
                        }
                    });

                    // ปุ่มสำหรับยกเลิกการเลือกทั้งหมด
                    JButton deselectAllButton = new JButton("ยกเลิกการเลือกทั้งหมด");
                    deselectAllButton.addActionListener((ActionEvent e2) -> {
                        for (int i = 0; i < table.getRowCount(); i++) {
                            table.setValueAt(false, i, 0); // ลบเครื่องหมายถูกจากทุกแถว
                        }
                    });

                    // ปุ่มสำหรับสร้างไฟล์ PDF
                    JButton exportPdfButton = new JButton("สร้าง PDF");
                    exportPdfButton.addActionListener(e2 -> {
                        int[] selectedRows = table.getSelectedRows();
                        if (selectedRows.length == 0) {
                            try {
                                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
                                Font thaiFont = new Font("Tahoma", Font.PLAIN, 14);
                                UIManager.put("OptionPane.messageFont", thaiFont);
                            } catch (Exception ex) {
                                ex.printStackTrace();
                            }
                            JOptionPane.showMessageDialog(frame, "กรุณาเลือกแถวก่อนที่จะสร้าง PDF", "ข้อความแจ้งเตือน", JOptionPane.WARNING_MESSAGE);
                            return;
                        }
                        try {
                            createPdf(selectedRows, data, columnNames);
                            try {
                                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
                                Font thaiFont = new Font("Tahoma", Font.PLAIN, 14);
                                UIManager.put("OptionPane.messageFont", thaiFont);
                            } catch (Exception ex) {
                                ex.printStackTrace();
                            }
                            JOptionPane.showMessageDialog(frame, "สร้าง PDF เสร็จสิ้นแล้ว", "สำเร็จ", JOptionPane.INFORMATION_MESSAGE);
                        } catch (IOException | UnsupportedLookAndFeelException | ClassNotFoundException |
                                 InstantiationException | IllegalAccessException ex) {
                            ex.printStackTrace();
                            JOptionPane.showMessageDialog(frame, "เกิดข้อผิดพลาดในการสร้าง PDF: " + ex.getMessage(), "ข้อผิดพลาด", JOptionPane.ERROR_MESSAGE);
                        }
                    });


                    // เพิ่ม JTable และ JTextArea เข้าไปในกรอบ
                    JPanel buttonPanel = new JPanel();
                    buttonPanel.setLayout(new FlowLayout());
                    buttonPanel.add(selectAllButton);
                    buttonPanel.add(deselectAllButton);
                    buttonPanel.add(exportPdfButton); // เพิ่มปุ่มสร้าง PDF

                    frame.getContentPane().removeAll();  // ลบข้อมูลเดิม
                    frame.add(uploadButton, BorderLayout.NORTH); // เพิ่มปุ่ม
                    frame.add(scrollPane, BorderLayout.CENTER); // เพิ่มตาราง
                    frame.add(buttonPanel, BorderLayout.SOUTH); // เพิ่มปุ่มที่ด้านล่าง
                    frame.add(new JScrollPane(detailArea), BorderLayout.EAST); // เพิ่มพื้นที่แสดงรายละเอียด
                    frame.revalidate(); // รีเฟรชกรอบ
                    frame.repaint();
                } catch (IOException ex) {
                    JOptionPane.showMessageDialog(frame, "ไม่สามารถอ่านไฟล์ได้: " + ex.getMessage());
                }
            }
        });

        // เพิ่มปุ่มเข้าไปในกรอบ
        frame.add(uploadButton, BorderLayout.NORTH);

        // แสดงกรอบ
        frame.setVisible(true);
    }

    private static Object[][] readExcelFile(File file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        // จำนวนแถวใน Excel
        int rowCount = sheet.getPhysicalNumberOfRows();
        // จำนวนคอลัมน์ใน Excel
        int colCount = sheet.getRow(0).getPhysicalNumberOfCells();

        // สร้าง Object[][] เพื่อเก็บข้อมูลในตาราง
        Object[][] data = new Object[rowCount - 1][colCount];

        Iterator<Row> rowIterator = sheet.iterator();
        rowIterator.next(); // ข้ามแถวแรกที่เป็นหัวข้อ

        int rowIndex = 0;
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            for (int colIndex = 0; colIndex < colCount; colIndex++) {
                Cell cell = row.getCell(colIndex);
                if (cell != null) {
                    switch (cell.getCellType()) {
                        case STRING:
                            data[rowIndex][colIndex] = cell.getStringCellValue();
                            break;
                        case NUMERIC:
                            data[rowIndex][colIndex] = cell.getNumericCellValue();
                            break;
                        default:
                            data[rowIndex][colIndex] = "";
                    }
                }
            }
            rowIndex++;
        }

        workbook.close();
        fis.close();

        return data;
    }

//    public static void createPdf(int[] selectedRows, Object[][] data, String[] columnNames) throws IOException {
//        // สร้าง PDF
//        File outputPdf = new File("output.pdf");
//        PdfWriter writer = new PdfWriter(outputPdf);
//        PdfDocument pdfDoc = new PdfDocument(writer);
//        Document document = new Document(pdfDoc);
//
//        // เขียนข้อมูลจากแถวที่เลือก
//        for (int row : selectedRows) {
//            document.add(new Paragraph("ข้อมูลแถวที่ " + (row + 1)));
//            for (int i = 0; i < columnNames.length; i++) {
//                document.add(new Paragraph(columnNames[i] + ": " + data[row][i]));
//            }
//            document.add(new Paragraph("\n"));
//
//            // ขึ้นหน้าใหม่สำหรับแถวถัดไป
//            document.add(new AreaBreak(AreaBreakType.NEXT_PAGE));
//        }
//
//        document.close();
//    }


    // Function to create PDF from selected rows
    private static void createPdf(int[] selectedRows, Object[][] data, String[] columnNames) throws IOException, UnsupportedLookAndFeelException, ClassNotFoundException, InstantiationException, IllegalAccessException {
        String outputPath = "selected_rows_output.pdf";
        PdfWriter writer = new PdfWriter(outputPath);
        PdfDocument pdf = new PdfDocument(writer);
        Document document = new Document(pdf);

        // ตั้งค่า Look and Feel ให้เหมือนกับระบบปัจจุบัน
        UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());

        URL fontUrl = ExcelUploadExample.class.getResource("/fonts/NotoSerifThai-Regular.ttf");
        if (fontUrl == null) {
            throw new IOException("Font file not found!");
        }
        PdfFont tfont = PdfFontFactory.createFont(fontUrl.toString());

        // Add the table headers
        document.add(new Paragraph("Stock Card"));

        // Create a table with 4 columns (for length, width, thickness, amount, etc.)
        for (int rowIndex : selectedRows) {
            document.add(new Paragraph("Row: " + (rowIndex + 1)));
            document.add(new Paragraph("ยาว (ฟุต): " + data[rowIndex][1]).setFont(tfont));
            document.add(new Paragraph("กว้าง (นิ้ว): " + data[rowIndex][2]).setFont(tfont));
            document.add(new Paragraph("หนา (นิ้ว): " + data[rowIndex][3]).setFont(tfont));
            document.add(new Paragraph("จำนวน/แผ่น: " + data[rowIndex][4]).setFont(tfont));
            document.add(new AreaBreak(AreaBreakType.NEXT_PAGE)); // Add a new page for each row
        }

        document.close(); // Finalize the document

        // Open the generated PDF in the default system PDF viewer
        File pdfFile = new File(outputPath);
        if (pdfFile.exists()) {
            Desktop.getDesktop().open(pdfFile);  // Open the file in default viewer
        }
    }
}
