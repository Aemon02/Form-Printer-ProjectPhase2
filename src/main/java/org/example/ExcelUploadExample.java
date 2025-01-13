package org.example;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Desktop;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.io.*;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.ListSelectionModel;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;

import com.itextpdf.io.font.PdfEncodings;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.geom.Rectangle;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.kernel.pdf.canvas.PdfCanvas;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.AreaBreak;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.element.Text;
import com.itextpdf.layout.properties.AreaBreakType;
import com.itextpdf.layout.properties.HorizontalAlignment;
import com.itextpdf.layout.properties.TextAlignment;
import com.itextpdf.layout.properties.VerticalAlignment;

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
                            "No", "Code", "Description", "เกรด", "หนา (นิ้ว)", "กว้าง (นิ้ว)", "ยาว (ฟุต)", "จำนวน/แผ่น", "ปริมาตร",
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
                            if (DateUtil.isCellDateFormatted(cell)) {
                                // หากเซลล์เป็นวันที่ ให้แปลงเป็นรูปแบบวันที่ที่ต้องการ
                                Date date = cell.getDateCellValue();
                                SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy"); // กำหนดรูปแบบวันที่
                                data[rowIndex][colIndex] = dateFormat.format(date); // แปลงวันที่เป็นสตริง
                            } else if (cell.getCellStyle().getDataFormatString().contains("%")) {
                                // หากเป็นเปอร์เซ็นต์
                                data[rowIndex][colIndex] = cell.getNumericCellValue() / 100;
                            } else {
                                double numericValue = cell.getNumericCellValue();
                                // แปลงค่าตัวเลขให้อยู่ในรูปแบบที่ต้องการ (ไม่ให้เป็น scientific notation)
                                if (numericValue == (long) numericValue) {
                                    // ถ้าเป็นจำนวนเต็ม เช่น 10.0 ให้แสดงเป็น 10
                                    data[rowIndex][colIndex] = (long) numericValue;
                                } else {
                                    // หากไม่ใช่จำนวนเต็ม เช่น 0.75678 ให้แสดงเป็นทศนิยมไม่เกิน 3 ตำแหน่ง
                                    numericValue = Math.round(numericValue * 1000.0) / 1000.0; // ปัดเศษให้เหลือ 3 ตำแหน่ง
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
        //วิธีการบันทึก1 PDF ใน project folder
//        String outputPath = "selected_rows_output.pdf";

        //วิธีการบันทึก2 ให้ path ของ file pdf ไปอยู่ที่ desktop
//        String userHome = System.getProperty("user.home");
//        String outputPath = userHome + File.separator + "Desktop" + File.separator + "selected_rows_output.pdf";

        //วิธีการบันทึก3 โดยเปิดหน้าต่างให้ผู้ใช้งานเลือกตำแหน่งที่จะบันทึกไฟล์
//        JFileChooser fileChooser = new JFileChooser();
//        fileChooser.setDialogTitle("บันทึกไฟล์ PDF");
//        fileChooser.setSelectedFile(new File("selected_rows_output.pdf")); // ตั้งชื่อไฟล์เริ่มต้น
//        int userSelection = fileChooser.showSaveDialog(null);

        //วิธีการบันทึก4
        // ใช้ ByteArrayOutputStream แทนการบันทึกไฟล์
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        PdfWriter writer = new PdfWriter(byteArrayOutputStream);
        PdfDocument pdf = new PdfDocument(writer);

        //วิธีการบันทึก3
//        if (userSelection == JFileChooser.APPROVE_OPTION) {
//            File fileToSave = fileChooser.getSelectedFile();
//            String outputPath = fileToSave.getAbsolutePath();
//            if (!outputPath.endsWith(".pdf")) {
//                outputPath += ".pdf"; // เพิ่ม .pdf หากผู้ใช้งานไม่ได้ระบุ
//            }

//            PdfWriter writer = new PdfWriter(outputPath);
//            PdfDocument pdf = new PdfDocument(writer);

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

                    // โหลดรูปภาพ
                    Image logoImage = new Image(ImageDataFactory.create("src/main/resources/logo/logo01.png"));

                    // เพิ่มข้อความหัวเรื่อง "Stock Card"
                    logoImage.setHeight(100); // กำหนดขนาดความสูงของรูป
                    logoImage.setHorizontalAlignment(HorizontalAlignment.LEFT);
                    Paragraph title = new Paragraph()
                            .add(logoImage)
                            .add(new Text("  Stock Card  ").setFont(tfont).setFontSize(27))
                            .add(new Text("ไม้สักแปรรูป").setFont(tfont).setFontSize(20))
                            .setTextAlignment(TextAlignment.CENTER);

                    document.add(title);

                    // สร้างตารางเพื่อแสดงข้อมูล
                    Table table = new Table(5);  // จำนวนคอลัมน์ในตาราง

                    // row 1
                    table.addCell(new com.itextpdf.layout.element.Cell()
                            .add(new Paragraph("ยาว (ฟุต)").setFont(tfont)).setFontSize(25)
                            .setTextAlignment(TextAlignment.CENTER)
                            .setVerticalAlignment(VerticalAlignment.MIDDLE));
                    table.addCell(new com.itextpdf.layout.element.Cell()
                            .add(new Paragraph("กว้าง (นิ้ว)").setFont(tfont)).setFontSize(25)
                            .setTextAlignment(TextAlignment.CENTER)
                            .setVerticalAlignment(VerticalAlignment.MIDDLE));
                    table.addCell(new com.itextpdf.layout.element.Cell()
                            .add(new Paragraph("หนา (นิ้ว)").setFont(tfont)).setFontSize(25)
                            .setTextAlignment(TextAlignment.CENTER)
                            .setVerticalAlignment(VerticalAlignment.MIDDLE));
                    table.addCell(new com.itextpdf.layout.element.Cell()
                            .add(new Paragraph("จำนวน/แผ่น").setFont(tfont).setFontSize(25))
                            .setTextAlignment(TextAlignment.CENTER)
                            .setVerticalAlignment(VerticalAlignment.MIDDLE));
                    table.addCell(new com.itextpdf.layout.element.Cell()
                            .add(new Paragraph("ปริมาตร").setFont(tfont).setFontSize(25))
                            .setTextAlignment(TextAlignment.CENTER)
                            .setVerticalAlignment(VerticalAlignment.MIDDLE));

                    // row 2
                    table.addCell(new com.itextpdf.layout.element.Cell()
                            .add(new Paragraph(String.valueOf(data[rowIndex][5])).setFont(tfont).setBold())
                            .setFontSize(40)
                            .setTextAlignment(TextAlignment.CENTER)
                            .setVerticalAlignment(VerticalAlignment.MIDDLE));
                    table.addCell(new com.itextpdf.layout.element.Cell()
                            .add(new Paragraph(String.valueOf(data[rowIndex][4])).setFont(tfont))
                            .setFontSize(40)
                            .setTextAlignment(TextAlignment.CENTER)
                            .setVerticalAlignment(VerticalAlignment.MIDDLE));
                    table.addCell(new com.itextpdf.layout.element.Cell()
                            .add(new Paragraph(String.valueOf(data[rowIndex][3])).setFont(tfont))
                            .setFontSize(40)
                            .setTextAlignment(TextAlignment.CENTER)
                            .setVerticalAlignment(VerticalAlignment.MIDDLE));
                    table.addCell(new com.itextpdf.layout.element.Cell()
                            .add(new Paragraph(String.valueOf(data[rowIndex][6])).setFont(tfont))
                            .setFontSize(40)
                            .setTextAlignment(TextAlignment.CENTER)
                            .setVerticalAlignment(VerticalAlignment.MIDDLE));
                    table.addCell(new com.itextpdf.layout.element.Cell()
                            .add(new Paragraph(String.valueOf(data[rowIndex][7])).setFont(tfont))
                            .setFontSize(40)
                            .setTextAlignment(TextAlignment.CENTER)
                            .setVerticalAlignment(VerticalAlignment.MIDDLE));

                    // row 3
                    table.addCell(new com.itextpdf.layout.element.Cell(1, 1)
                            .add(new Paragraph()
                                    .add(new Text("เกรด").setFont(tfont).setFontSize(20))
                                    .add(new Text("\n ").setFont(tfont).setFontSize(20))
                                    .add(new Text(String.valueOf(data[rowIndex][3])).setFont(tfont).setFontSize(40).setBold())
                            )
                            .setTextAlignment(TextAlignment.CENTER));
                    table.addCell(new com.itextpdf.layout.element.Cell(1, 3)
                            .add(new Paragraph()
                                    .add(new Text("รหัสสินค้า :  ").setFont(tfont).setFontSize(20))  // ตัวหนังสือปกติ
                                    .add(new Text(String.valueOf(data[rowIndex][1])).setFont(tfont).setFontSize(20).setBold())  // ข้อความที่เป็นตัวหนา
                                    .add(new Text("\nชื่อสินค้า :  ").setFont(tfont).setFontSize(20))  // ตัวหนังสือปกติ
                                    .add(new Text(String.valueOf(data[rowIndex][2])).setFont(tfont).setFontSize(20).setBold())  // ข้อความที่เป็นตัวหนา
                            )
                            .setTextAlignment(TextAlignment.LEFT)
                            .setVerticalAlignment(VerticalAlignment.MIDDLE)
                            .setFont(tfont));
                    table.addCell(new com.itextpdf.layout.element.Cell(1, 2)
                            .add(new Paragraph()
                                            .add(new Text("วันที่").setFont(tfont).setFontSize(20))  // ตัวหนังสือปกต
                                            .add(new Text("\n ").setFont(tfont).setFontSize(20))  // ตัวหนังสือปกติ
                                            .add(new Text(String.valueOf(data[rowIndex][8])).setFont(tfont).setFontSize(20).setBold())  // ข้อความที่เป็นตัวหนา
                                            .setTextAlignment(TextAlignment.CENTER)
//                                .add(new Text(String.valueOf(data[rowIndex][9])).setFont(tfont).setFontSize(15).setBold())  // ข้อความที่เป็นตัวหนา
                            )
                            .setVerticalAlignment(VerticalAlignment.MIDDLE)
                            .setFont(tfont));
                    table.addCell(new com.itextpdf.layout.element.Cell(1, 3)
//                        .add(new Paragraph("Location"))
                            .add(new Paragraph("barcode").setFontSize(15))
                            .setTextAlignment(TextAlignment.CENTER)
                            .setVerticalAlignment(VerticalAlignment.MIDDLE));
                    table.addCell(new com.itextpdf.layout.element.Cell(1, 3)
                            .add(new Paragraph("Location").setFontSize(15))
                            .setTextAlignment(TextAlignment.CENTER)
                            .setVerticalAlignment(VerticalAlignment.MIDDLE));

//                // ตั้งค่าความกว้างของตารางให้เต็มหน้ากระดาษ
                    float pageWidth = pdf.getDefaultPageSize().getWidth();
                    float pageHeight = pdf.getDefaultPageSize().getHeight();
                    float padding = 36f;
                    table.setWidth(pageWidth - 2 * padding);  // กำหนดความกว้างให้เต็มหน้ากระดาษ
                    table.setHeight(pageHeight - 2 * padding);  // กำหนดความสูงให้เต็มหน้ากระดาษ


                    // เพิ่มตารางลงในเอกสาร
                    document.add(table);

                    // เพิ่มกรอบสี่เหลี่ยมรอบตาราง
                    PdfCanvas pdfCanvas = new PdfCanvas(pdf.getLastPage());
                    Rectangle rectangle = new Rectangle(padding, padding, pageWidth - 2 * padding, pdf.getDefaultPageSize().getHeight() - 2 * padding);
                    pdfCanvas.rectangle(rectangle);
                    pdfCanvas.stroke();

                    document.add(new AreaBreak(AreaBreakType.NEXT_PAGE));

                }
            }

            document.close();

            // วิธีการบันทึก3  แจ้งให้ผู้ใช้งานทราบว่าการสร้าง PDF สำเร็จ
//            JOptionPane.showMessageDialog(null, "PDF ถูกบันทึกเรียบร้อยแล้ว: " + outputPath, "สำเร็จ", JOptionPane.INFORMATION_MESSAGE);

            // แสดง PDF โดยตรงจากหน่วยความจำ
            byte[] pdfBytes = byteArrayOutputStream.toByteArray();
            InputStream pdfStream = new ByteArrayInputStream(pdfBytes);

            // ใช้ PDF viewer ภายนอก เช่น PDF.js หรือเครื่องมืออื่นๆ
            File tempPdf = File.createTempFile("print_wood", ".pdf");
            try (OutputStream tempOut = new FileOutputStream(tempPdf)) {
                tempOut.write(pdfBytes);
            }

            // เปิด PDF ด้วย viewer ของระบบ
            Desktop.getDesktop().open(tempPdf);

            // เปิดไฟล์ PDF ที่สร้างขึ้นในโปรแกรม PDF viewer เริ่มต้น วิธีการบันทึก1 2 3
//            File pdfFile = new File(outputPath);
//            if (pdfFile.exists()) {
//                Desktop.getDesktop().open(pdfFile);
//            }
//        } else {
//            JOptionPane.showMessageDialog(null, "การบันทึก PDF ถูกยกเลิก", "ยกเลิก", JOptionPane.WARNING_MESSAGE);
//        }
    }
}

//    // Open the generated PDF in the default system PDF viewer
//        File pdfFile = new File(outputPath);
//        if (pdfFile.exists()) {
//            Desktop.getDesktop().open(pdfFile);
//        }
