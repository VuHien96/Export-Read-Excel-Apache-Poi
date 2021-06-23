package com.vuhien.application.controller;

import com.vuhien.application.model.dto.UserDTO;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

public class ReadExcelUserController {

    public static final SimpleDateFormat simpleDateFormat = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
    public static final int COLUMN_INDEX_ID = 0;
    public static final int COLUMN_INDEX_EMAIL = 1;
    public static final int COLUMN_INDEX_FULL_NAME = 2;
    public static final int COLUMN_INDEX_AGE = 3;
    public static final int COLUMN_INDEX_ADDRESS = 4;
    public static final int COLUMN_INDEX_PHONE = 5;
    public static final int COLUMN_INDEX_CREATED_AT = 6;
    public static final int COLUMN_INDEX_MODIFIED_AT = 7;

    public static void main(String[] args) throws IOException {
        final String excelFilePath = "D:/customer.xlsx";
        final List<UserDTO> userDTOS = readExcel(excelFilePath);
        for (UserDTO userDTO : userDTOS) {
            System.out.println(userDTO);
        }
    }

    public static List<UserDTO> readExcel(String excelFilePath) throws IOException {
        List<UserDTO> listBooks = new ArrayList<>();

        // Get file
        InputStream inputStream = new FileInputStream(new File(excelFilePath));

        // Get workbook
        Workbook workbook = getWorkbook(inputStream, excelFilePath);

        // Get sheet
        Sheet sheet = workbook.getSheetAt(0);

        // Get all rows
        Iterator<Row> iterator = sheet.iterator();
        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            if (nextRow.getRowNum() == 0) {
                // Ignore header
                continue;
            }

            // Get all cells
            Iterator<Cell> cellIterator = nextRow.cellIterator();

            // Read cells and set value for book object
            UserDTO userDTO = new UserDTO();
            while (cellIterator.hasNext()) {
                //Read cell
                Cell cell = cellIterator.next();
                Object cellValue = getCellValue(cell);
                if (cellValue == null || cellValue.toString().isEmpty()) {
                    continue;
                }
                // Set value for book object
                int columnIndex = cell.getColumnIndex();
                switch (columnIndex) {
                    case COLUMN_INDEX_ID:
                        userDTO.setId(new BigDecimal((double) cellValue).intValue());
                        break;
                    case COLUMN_INDEX_EMAIL:
                        userDTO.setEmail((String) getCellValue(cell));
                        break;
                    case COLUMN_INDEX_FULL_NAME:
                        userDTO.setFullName((String) getCellValue(cell));
                        break;
                    case COLUMN_INDEX_AGE:
                        userDTO.setAge(new BigDecimal((double) cellValue).intValue());
                        break;
                    case COLUMN_INDEX_ADDRESS:
                        userDTO.setAddress((String) getCellValue(cell));
                        break;
                    case COLUMN_INDEX_PHONE:
                        userDTO.setPhone((String) getCellValue(cell));
                        break;
                    case COLUMN_INDEX_CREATED_AT:
                        userDTO.setCreatedAt((Date) getCellValue(cell));
                        break;
                    case COLUMN_INDEX_MODIFIED_AT:
                        userDTO.setModifiedAt((Date) getCellValue(cell));
                        break;
                    default:
                        break;
                }

            }
            listBooks.add(userDTO);
        }

        workbook.close();
        inputStream.close();

        return listBooks;
    }

    // Get Workbook
    private static Workbook getWorkbook(InputStream inputStream, String excelFilePath) throws IOException {
        Workbook workbook = null;
        if (excelFilePath.endsWith("xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        } else if (excelFilePath.endsWith("xls")) {
            workbook = new HSSFWorkbook(inputStream);
        } else {
            throw new IllegalArgumentException("The specified file is not Excel file");
        }

        return workbook;
    }

    // Get cell value
    private static Object getCellValue(Cell cell) {
        CellType cellType = cell.getCellTypeEnum();
        Object cellValue = null;
        switch (cellType) {
            case BOOLEAN:
                cellValue = cell.getBooleanCellValue();
                break;
            case FORMULA:
                Workbook workbook = cell.getSheet().getWorkbook();
                FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                cellValue = evaluator.evaluate(cell).getNumberValue();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    cellValue = cell.getDateCellValue();
                } else {
                    cellValue = cell.getNumericCellValue();
                }
                break;
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            case _NONE:
            case BLANK:
            case ERROR:
                break;
            default:
                break;
        }

        return cellValue;
    }
}
