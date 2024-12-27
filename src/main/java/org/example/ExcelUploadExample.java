package org.example;

import com.itextpdf.io.font.PdfEncodings;
import com.itextpdf.io.font.constants.StandardFonts;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.AreaBreak;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.properties.AreaBreakType;

//import com.itextpdf.layout.element.Cell;
import com.itextpdf.layout.properties.TextAlignment;
import org.apache.poi.ss.usermodel.Cell;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.itextpdf.kernel.font.*;
import com.itextpdf.kernel.pdf.*;

import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.List;

import javax.swing.*;
import javax.swing.table.JTableHeader;
import java.awt.*;
import java.awt.Desktop;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.Iterator;

//การตั้งค่ารูปแบบใน pdf
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.geom.Rectangle;
import com.itextpdf.layout.Canvas;
import com.itextpdf.kernel.pdf.canvas.PdfCanvas;
import com.itextpdf.layout.element.Table;

import static com.itextpdf.kernel.pdf.PdfName.BaseFont;

public class ExcelUploadExample {

    public static void main(String[] args) {
        // สร้างกรอบ (Frame)
        JFrame frame = new JFrame("ตารางสินค้า");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(1200, 800);
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
        uploadButton.setPreferredSize(new Dimension(0, 50)); // ปรับขนาดของปุ่ม
        uploadButton.setBackground(new Color(255, 162, 0)); // เปลี่ยนสีพื้นหลังเป็นสีน้ำเงิน
//        uploadButton.setOpaque(true);  // ทำให้ปุ่มมีความทึบแสง
//        uploadButton.setBorderPainted(true); // ทำให้กรอบของปุ่มปรากฏ
        uploadButton.setFont(new Font("Tahoma", Font.CENTER_BASELINE, 17));
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
                            "ลำดับ","รหัสสินค้า","ชื่อสินค้า",  "เกรด", "ยาว (ฟุต)", "กว้าง (นิ้ว)", "หนา (นิ้ว)", "จำนวน/แผ่น",
                            "วันที่", "Location", "Barcode"
                    };

                    // สร้าง DefaultTableModel และ JCheckBox ในคอลัมน์ "เลือก"
                    Object[][] dataWithCheckbox = new Object[data.length][columnNames.length + 1];
                    for (int i = 0; i < data.length; i++) {
                        dataWithCheckbox[i][0] = false; // คอลัมน์ "เลือก" สำหรับ JCheckBox
                        for (int j = 0; j < columnNames.length; j++) {
                            dataWithCheckbox[i][j + 1] = data[i][j]; // ข้อมูลจริงจาก Excel
                        }
                    }

                    // รวมคอลัมน์ "เลือก" เข้าไปใน columnNames
                    String[] allColumnNames = new String[columnNames.length + 1];
                    allColumnNames[0] = "เลือก"; // ชื่อคอลัมน์ "เลือก"
                    System.arraycopy(columnNames, 0, allColumnNames, 1, columnNames.length); // คัดลอก columnNames ลงใน allColumnNames

                    JTable table = new JTable(dataWithCheckbox, allColumnNames) {
                        @Override
                        public Class<?> getColumnClass(int columnIndex) {
                            if (columnIndex == 0) {
                                return Boolean.class; // ให้คอลัมน์แรกเป็น JCheckBox
                            }
                            return super.getColumnClass(columnIndex);
                        }
                    };
                    table.setFont(new Font("Tahoma", Font.PLAIN, 14));  // ใช้ฟอนต์ที่รองรับภาษาไทย
                    table.setFillsViewportHeight(true);
                    table.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);  // การเลือกหลายแถว
                    table.getTableHeader().setFont(new Font("Tahoma", Font.BOLD, 15));
                    JScrollPane scrollPane = new JScrollPane(table);

                    // สร้าง JTextArea สำหรับแสดงรายละเอียดของแถวที่เลือก
                    JTextArea detailArea = new JTextArea(10, 40);
                    detailArea.setFont(new Font("Tahoma", Font.PLAIN, 14));  // ฟอนต์ที่รองรับภาษาไทย
                    detailArea.setEditable(false);

                    // ฟังการเลือกแถวใน JTable
                    table.getSelectionModel().addListSelectionListener(e1 -> {
                        if (!e1.getValueIsAdjusting()) {
                            StringBuilder details = new StringBuilder();

                            // แสดงข้อมูลของทุกแถวที่มี checkbox เป็น true
                            for (int i = 0; i < table.getRowCount(); i++) {
                                Boolean isChecked = (Boolean) table.getValueAt(i, 0);
                                if (isChecked != null && isChecked) {
                                    details.append("แถวที่ ").append(i + 1).append(":\n");
                                    for (int j = 0; j < allColumnNames.length; j++) {
                                        Object value = table.getValueAt(i, j);
                                        String displayValue = value != null ? value.toString() : "";
                                        details.append(allColumnNames[j]).append(": ")
                                                .append(displayValue).append("\n");
                                    }
                                    details.append("\n");
                                }
                            }

                            // ตั้งค่า checkbox เป็น true สำหรับแถวที่เพิ่งเลือก
                            int[] selectedRows = table.getSelectedRows();
                            for (int row : selectedRows) {
                                if (row >= 0 && row < table.getRowCount()) {
                                    table.setValueAt(true, row, 0);
                                }
                            }

                            // อัพเดทพื้นที่แสดงรายละเอียด
                            detailArea.setText(details.toString());
                            detailArea.setCaretPosition(0);
                        }
                    });

                    // ปุ่มสำหรับเลือกทั้งหมด
                    JButton selectAllButton = new JButton("เลือกทั้งหมด");
                    selectAllButton.setPreferredSize(new java.awt.Dimension(200, 50)); // ปรับขนาดของปุ่ม
                    selectAllButton.setBackground(new java.awt.Color(0, 123, 255)); // เปลี่ยนสีพื้นหลังเป็นสีน้ำเงิน
//                    selectAllButton.setForeground(java.awt.Color.WHITE); // เปลี่ยนสีข้อความเป็นสีขาว
                    selectAllButton.setFont(new Font("Tahoma", Font.PLAIN, 15));
                    selectAllButton.addActionListener((ActionEvent e2) -> {
                        for (int i = 0; i < table.getRowCount(); i++) {
                            table.setValueAt(true, i, 0); // ตั้งค่าเครื่องหมายถูกในทุกแถว
                        }
                    });

                    // ปุ่มสำหรับยกเลิกการเลือกทั้งหมด
                    JButton deselectAllButton = new JButton("ยกเลิกการเลือกทั้งหมด");
                    deselectAllButton.setPreferredSize(new java.awt.Dimension(200, 50)); // ปรับขนาดของปุ่ม
                    deselectAllButton.setBackground(new java.awt.Color(220, 53, 69)); // เปลี่ยนสีพื้นหลังเป็นสีแดง
                    deselectAllButton.setFont(new Font("Tahoma", Font.PLAIN, 15));
                    deselectAllButton.addActionListener((ActionEvent e2) -> {
                        for (int i = 0; i < table.getRowCount(); i++) {
                            table.setValueAt(false, i, 0); // ลบเครื่องหมายถูกจากทุกแถว
                        }
                    });

                    // ปุ่มสำหรับสร้างไฟล์ PDF
                    JButton exportPdfButton = new JButton("สร้าง PDF");
                    exportPdfButton.setPreferredSize(new java.awt.Dimension(200, 50)); // ปรับขนาดของปุ่ม
                    exportPdfButton.setBackground(new java.awt.Color(40, 167, 69)); // เปลี่ยนสีพื้นหลังเป็นสีเขียว
                    exportPdfButton.setFont(new Font("Tahoma", Font.PLAIN, 15));
                    exportPdfButton.addActionListener(e2 -> {
                        // เปลี่ยนจากการใช้ selectedRows เป็นการตรวจสอบ checkbox
                        List<Integer> checkedRows = new ArrayList<>();
                        for (int i = 0; i < table.getRowCount(); i++) {
                            Boolean isChecked = (Boolean) table.getValueAt(i, 0);
                            if (isChecked != null && isChecked) {
                                checkedRows.add(i);
                            }
                        }

                        if (checkedRows.isEmpty()) {
                            try {
                                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
                                Font thaiFont = new Font("Tahoma", Font.PLAIN, 14);
                                UIManager.put("OptionPane.messageFont", thaiFont);
                            } catch (Exception ex) {
                                ex.printStackTrace();
                            }
                            JOptionPane.showMessageDialog(frame,
                                    "กรุณาเลือกแถวก่อนที่จะสร้าง PDF",
                                    "ข้อความแจ้งเตือน",
                                    JOptionPane.WARNING_MESSAGE);
                            return;
                        }

                        try {
                            try {
                                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
                                Font thaiFont = new Font("Tahoma", Font.PLAIN, 14);
                                UIManager.put("OptionPane.messageFont", thaiFont);
                            } catch (Exception ex) {
                                ex.printStackTrace();
                            }

                            // แปลง List<Integer> เป็น int[]
                            int[] selectedRows = checkedRows.stream().mapToInt(Integer::intValue).toArray();
                            for (int rowIndex : selectedRows) {
                                if (data[rowIndex].length < allColumnNames.length) {
//                                    System.out.println("Skipping incomplete row: " + rowIndex);
//                                    continue; // ข้ามแถวนี้ไป
                                    throw new IllegalArgumentException("ข้อมูลในแถวไม่สมบูรณ์: แถว " + rowIndex);
                                }
                            }
                            createPdf(selectedRows, data, allColumnNames);
                            JOptionPane.showMessageDialog(frame,
                                    "สร้าง PDF เสร็จสิ้นแล้ว",
                                    "สำเร็จ",
                                    JOptionPane.INFORMATION_MESSAGE);
                        } catch (Exception ex) {
                            ex.printStackTrace();
                            JOptionPane.showMessageDialog(frame,
                                    "เกิดข้อผิดพลาดในการสร้าง PDF: " + ex.getMessage(),
                                    "ข้อผิดพลาด",
                                    JOptionPane.ERROR_MESSAGE);
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
//                    frame.add(new JScrollPane(detailArea), BorderLayout.EAST); // เพิ่มพื้นที่แสดงรายละเอียด
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
        Object[][] data = new Object[rowCount - 1][colCount + 20]; //มีผลต่อจำนวน column ที่ทำให้ print pdf file

        Iterator<Row> rowIterator = sheet.iterator();
        rowIterator.next(); // ข้ามแถวแรกที่เป็นหัวข้อ

        int rowIndex = 0;
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            for (int colIndex = 0; colIndex < colCount; colIndex++) {
                Cell cell = (Cell) row.getCell(colIndex);

                if (cell != null) {
                    switch (cell.getCellType()) {
                        case STRING:
                            String cellValue = cell.getStringCellValue();
                            // ตรวจสอบว่าในสตริงมีตัวอักษรพิมพ์ใหญ่แค่ 1 ตัวหรือไม่
                            String upperCaseLetter = cellValue.replaceAll("[^A-Z]", ""); // ลบอักขระที่ไม่ใช่ตัวพิมพ์ใหญ่
                            if (upperCaseLetter.length() == 1) {
                                // หากมีตัวพิมพ์ใหญ่แค่ 1 ตัว ก็ให้แสดงตัวอักษรนั้น
                                data[rowIndex][colIndex] = upperCaseLetter;
                            } else {
                                // หากไม่มีหรือตัวพิมพ์ใหญ่หลายตัว ให้แสดงข้อมูลทั้งหมด
                                data[rowIndex][colIndex] = cellValue;
                            }
                            break;
                        case NUMERIC:
                            // ตรวจสอบว่าเป็นค่าที่เป็นเลขทศนิยมหรือไม่
                            if (cell.getCellStyle().getDataFormatString().contains("%")) {
                                // หากเป็นเปอร์เซ็นต์
                                data[rowIndex][colIndex] = cell.getNumericCellValue() / 100;
                            } else {
                                double numericValue = cell.getNumericCellValue();
                                // แปลงค่าตัวเลขให้อยู่ในรูปแบบที่ต้องการ (ไม่ให้เป็น scientific notation)
                                if (numericValue == (long) numericValue) {
                                    // ถ้าเป็นจำนวนเต็ม เช่น 10.0 ให้แสดงเป็น 10
                                    data[rowIndex][colIndex] = (long) numericValue;
                                } else {
                                    // หากไม่ใช่จำนวนเต็ม เช่น 0.75 ให้แสดงเป็นทศนิยม
                                    data[rowIndex][colIndex] = numericValue;
                                }
                            }
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


    private static void createPdf(int[] selectedRows, Object[][] data, String[] columnNames) throws IOException, UnsupportedLookAndFeelException, ClassNotFoundException, InstantiationException, IllegalAccessException {
//        String outputPath = "selected_rows_output.pdf";
        String userHome = System.getProperty("user.home");
        String outputPath = userHome + File.separator + "Desktop" + File.separator + "selected_rows_output.pdf";
        PdfWriter writer = new PdfWriter(outputPath);
        PdfDocument pdf = new PdfDocument(writer);

        // ตั้งค่าขนาดกระดาษและหมุนเป็นแนวนอน
        pdf.setDefaultPageSize(PageSize.A4.rotate());
        Document document = new Document(pdf);

        // ตั้งค่า Look and Feel ให้เหมือนกับระบบปัจจุบัน
        UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        URL fontUrl = ExcelUploadExample.class.getResource("/fonts/NotoSerifThai-Regular.ttf");
        if (fontUrl == null) {
            throw new IOException("Font file not found!");
        }
        PdfFont tfont = PdfFontFactory.createFont(fontUrl.toString());

        // สร้างแถวตามข้อมูลที่เลือก
        for (int rowIndex : selectedRows) {
            // ตรวจสอบว่า rowIndex อยู่ในช่วงที่ถูกต้อง
            if (rowIndex >= 0 && rowIndex < data.length) {

                // เพิ่มข้อความหัวเรื่อง "Stock Card"
                Paragraph title = new Paragraph("Stock Card")
                        .setFont(tfont)
                        .setFontSize(24)  // ตั้งขนาดฟอนต์
                        .setTextAlignment(TextAlignment.CENTER);  // จัดให้อยู่กลาง
                document.add(title);

                // สร้างตารางเพื่อแสดงข้อมูล
                Table table1 = new Table(4).setMinHeight(300);  // จำนวนคอลัมน์ในตาราง
                Table table2 = new Table(4).setMinHeight(150);  // จำนวนคอลัมน์ในตาราง
                Table table3 = new Table(3).setMinHeight(160);  // จำนวนคอลัมน์ในตาราง

//                // เพิ่มหัวตาราง
//                table1.addCell(new com.itextpdf.layout.element.Cell().add(new Paragraph("Row")));
                table1.addCell(new com.itextpdf.layout.element.Cell().add(new Paragraph("ยาว (ฟุต)").setFont(tfont))).setTextAlignment(TextAlignment.CENTER).setFontSize(17);
                table1.addCell(new com.itextpdf.layout.element.Cell().add(new Paragraph("กว้าง (นิ้ว)").setFont(tfont))).setTextAlignment(TextAlignment.CENTER).setFontSize(17);
                table1.addCell(new com.itextpdf.layout.element.Cell().add(new Paragraph("หนา (นิ้ว)").setFont(tfont))).setTextAlignment(TextAlignment.CENTER).setFontSize(17);
                table1.addCell(new com.itextpdf.layout.element.Cell().add(new Paragraph("จำนวน/แผ่น").setFont(tfont))).setTextAlignment(TextAlignment.CENTER).setFontSize(17);

                // เพิ่มข้อมูลในแต่ละแถว
//                table1.addCell(new com.itextpdf.layout.element.Cell().add(new Paragraph("Row: " + (rowIndex + 1))));
                table1.addCell(new com.itextpdf.layout.element.Cell().add(new Paragraph(String.valueOf(data[rowIndex][4])).setFont(tfont))).setTextAlignment(TextAlignment.CENTER);
                table1.addCell(new com.itextpdf.layout.element.Cell().add(new Paragraph(String.valueOf(data[rowIndex][5])).setFont(tfont))).setTextAlignment(TextAlignment.CENTER);
                table1.addCell(new com.itextpdf.layout.element.Cell().add(new Paragraph(String.valueOf(data[rowIndex][6])).setFont(tfont))).setTextAlignment(TextAlignment.CENTER);
                table1.addCell(new com.itextpdf.layout.element.Cell().add(new Paragraph(String.valueOf(data[rowIndex][7])).setFont(tfont))).setTextAlignment(TextAlignment.CENTER);


                table3.addCell(new com.itextpdf.layout.element.Cell().add(new Paragraph(String.valueOf(data[rowIndex][3])))).setTextAlignment(TextAlignment.CENTER);
                table3.addCell(new com.itextpdf.layout.element.Cell().add(new Paragraph("2"))).setTextAlignment(TextAlignment.CENTER);

//                // ตั้งค่าความกว้างของตารางให้เต็มหน้ากระดาษ
                float pageWidth = pdf.getDefaultPageSize().getWidth();
                float pageHeight = pdf.getDefaultPageSize().getHeight();
                float padding = 36f; // กำหนดระยะห่างจากขอบ
                table1.setWidth(pageWidth - 2 * padding);  // กำหนดความกว้างให้เต็มหน้ากระดาษ
//                table1.setHeight(pageHeight - 2 * padding);  // กำหนดความสูงให้เต็มหน้ากระดาษ

//                table2.setWidth(pageWidth - 2 * padding);  // กำหนดความกว้างให้เต็มหน้ากระดาษ
                table3.setWidth(pageWidth - 2 * padding);  // กำหนดความกว้างให้เต็มหน้ากระดาษ

                // เพิ่มตารางลงในเอกสาร
                document.add(table1);
//                document.add(table2);
                document.add(table3);

                // เพิ่มกรอบสี่เหลี่ยมรอบตาราง
                PdfCanvas pdfCanvas = new PdfCanvas(pdf.getLastPage());
                Rectangle rectangle = new Rectangle(padding, padding, pageWidth - 2 * padding, pdf.getDefaultPageSize().getHeight() - 2 * padding);
                pdfCanvas.rectangle(rectangle);
                pdfCanvas.stroke();

                document.add(new AreaBreak(AreaBreakType.NEXT_PAGE));

            }
        }

        document.close();

        // Open the generated PDF in the default system PDF viewer
        File pdfFile = new File(outputPath);
        if (pdfFile.exists()) {
            Desktop.getDesktop().open(pdfFile);
        }
    }
}
