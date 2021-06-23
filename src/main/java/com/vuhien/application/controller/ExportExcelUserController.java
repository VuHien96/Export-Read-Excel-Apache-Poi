package com.vuhien.application.controller;

import com.vuhien.application.model.dto.UserDTO;
import com.vuhien.application.service.UserService;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

@RestController
public class ExportExcelUserController {

    public static final int COLUMN_INDEX_ID = 0;
    public static final int COLUMN_INDEX_EMAIL = 1;
    public static final int COLUMN_INDEX_FULL_NAME = 2;
    public static final int COLUMN_INDEX_AGE = 3;
    public static final int COLUMN_INDEX_ADDRESS = 4;
    public static final int COLUMN_INDEX_PHONE = 5;
    public static final int COLUMN_INDEX_CREATED_AT = 6;
    public static final int COLUMN_INDEX_MODIFIED_AT = 7;
    private static CellStyle cellStyle = null;

    public final String excelFilePath = "D:/customer.xlsx";

    @Autowired
    private UserService userService;

    @GetMapping("/users")
    public ResponseEntity<Object> getListUsers() {
        List<UserDTO> userDTOS = userService.getListUsers();
        writeExcel(userDTOS, excelFilePath);
        return ResponseEntity.ok(userDTOS);
    }

    // create workbook
    private static Workbook getWorkbook(String excelFilePath) {
        Workbook workbook = null;
        if (excelFilePath.endsWith("xlsx")) {
            workbook = new XSSFWorkbook();
        } else if (excelFilePath.endsWith("xls")) {
            workbook = new HSSFWorkbook();
        } else {
            throw new IllegalArgumentException("File không đúng định dạng!");
        }
        return workbook;
    }

    // Create CellStyle for header
    private static CellStyle createCellStyleHeader(Sheet sheet) {
        // Create font
        Font font = sheet.getWorkbook().createFont();
        font.setFontName("Times New Roman");
        font.setBold(true);
        font.setFontHeightInPoints((short) 14);
        font.setColor(IndexedColors.GREY_50_PERCENT.getIndex());

        // Create cellStyle
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setFont(font);
        cellStyle.setFillBackgroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
        cellStyle.setFillPattern(FillPatternType.LEAST_DOTS);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);

        return cellStyle;
    }

    private static void writeHeader(Sheet sheet, int rowIndex) {
        CellStyle cellStyle = createCellStyleHeader(sheet);

        //create row
        Row row = sheet.createRow(rowIndex);

        //create cells
        Cell cell = null;

        cell = row.createCell(COLUMN_INDEX_ID);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("ID");

        cell = row.createCell(COLUMN_INDEX_EMAIL);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Email");

        cell = row.createCell(COLUMN_INDEX_FULL_NAME);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Tên");

        cell = row.createCell(COLUMN_INDEX_AGE);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Tuổi");

        cell = row.createCell(COLUMN_INDEX_ADDRESS);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Địa chỉ");

        cell = row.createCell(COLUMN_INDEX_PHONE);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Số điện thoại");

        cell = row.createCell(COLUMN_INDEX_CREATED_AT);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Ngày tạo");

        cell = row.createCell(COLUMN_INDEX_MODIFIED_AT);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Ngày sửa");
    }

    private static void writeUser(UserDTO userDTO, Row row, Sheet sheet) {
        Cell cell = null;

        if (cellStyle == null) {
            // Create font
            Font font = sheet.getWorkbook().createFont();
            font.setFontName("Times New Roman");
            font.setBold(true);
            font.setFontHeightInPoints((short) 11);
            font.setColor(IndexedColors.GREY_50_PERCENT.getIndex());
            //Create CellStyle
            Workbook workbook = row.getSheet().getWorkbook();
            cellStyle = workbook.createCellStyle();
            cellStyle.setFont(font);
            cellStyle.setFillBackgroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
        }

        cell = row.createCell(COLUMN_INDEX_ID);
        cell.setCellValue(userDTO.getId());
        cell.setCellStyle(cellStyle);

        cell = row.createCell(COLUMN_INDEX_EMAIL);
        cell.setCellStyle(cellStyle);
        cell.setCellValue(userDTO.getEmail());

        cell = row.createCell(COLUMN_INDEX_FULL_NAME);
        cell.setCellStyle(cellStyle);
        cell.setCellValue(userDTO.getFullName());

        cell = row.createCell(COLUMN_INDEX_AGE);
        cell.setCellStyle(cellStyle);
        cell.setCellValue(userDTO.getAge());

        cell = row.createCell(COLUMN_INDEX_ADDRESS);
        cell.setCellValue(userDTO.getAddress());
        cell.setCellStyle(cellStyle);

        cell = row.createCell(COLUMN_INDEX_PHONE);
        cell.setCellStyle(cellStyle);
        cell.setCellValue(userDTO.getPhone());

        Workbook workbook = row.getSheet().getWorkbook();
        CreationHelper creationHelper = workbook.getCreationHelper();
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("dd/MM/yyyy HH:mm:ss"));
        cellStyle.setFillBackgroundColor(IndexedColors.LEMON_CHIFFON.getIndex());
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);


        cell = row.createCell(COLUMN_INDEX_CREATED_AT);
        cell.setCellStyle(cellStyle);
        cell.setCellValue(userDTO.getCreatedAt());

        cell = row.createCell(COLUMN_INDEX_MODIFIED_AT);
        cell.setCellStyle(cellStyle);
        cell.setCellValue(userDTO.getModifiedAt());

    }

    private static void writeExcel(List<UserDTO> userDTOS, String excelFilePath) {
        // Create Workbook
        Workbook workbook = getWorkbook(excelFilePath);
        //Create sheet
        Sheet sheet = workbook.createSheet("Danh sách khách hàng");
        int rowIndex = 0;
        //write header
        writeHeader(sheet, rowIndex);
        rowIndex++;
        for (UserDTO userDTO : userDTOS) {
            Row row = sheet.createRow(rowIndex);
            //ghi dữ liệu vào hàng
            writeUser(userDTO, row, sheet);
            rowIndex++;
        }

        autosizeColumn(sheet, rowIndex);

        createOutputFile(workbook, excelFilePath);
    }

    // Auto resize column width
    private static void autosizeColumn(Sheet sheet, int lastColumn) {
        for (int columnIndex = 0; columnIndex < lastColumn; columnIndex++) {
            sheet.autoSizeColumn(columnIndex);
        }
    }

    private static void createOutputFile(Workbook workbook, String excelFilePath) {
        try (OutputStream os = new FileOutputStream(excelFilePath)) {
            workbook.write(os);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
