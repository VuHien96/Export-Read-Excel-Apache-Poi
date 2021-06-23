package com.vuhien.application.controller;

import com.vuhien.application.model.dto.UserDTO;
import com.vuhien.application.service.UserService;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

@RestController
public class PieChart {

    @Autowired
    private UserService userService;

    public static final int COLUMN_INDEX_ID = 0;
    public static final int COLUMN_INDEX_FULL_NAME = 1;
    public static final int COLUMN_INDEX_AGE = 2;
    private static CellStyle cellStyle = null;
    public final String excelFilePath = "D:/customer.xlsx";

    @GetMapping("/api/users")
    public ResponseEntity<Object> getListUsers() throws IOException {
        List<UserDTO> userDTOS = userService.getListUsers();
        writeExcel(userDTOS, excelFilePath);
        return ResponseEntity.ok(userDTOS);
    }

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

        cell = row.createCell(COLUMN_INDEX_FULL_NAME);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Tên");

        cell = row.createCell(COLUMN_INDEX_AGE);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("Tuổi");

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

        cell = row.createCell(COLUMN_INDEX_FULL_NAME);
        cell.setCellStyle(cellStyle);
        cell.setCellValue(userDTO.getFullName());

        cell = row.createCell(COLUMN_INDEX_AGE);
        cell.setCellStyle(cellStyle);
        cell.setCellValue(userDTO.getAge());

    }

    // Auto resize column width
    private static void autosizeColumn(Sheet sheet, int lastColumn) {
        for (int columnIndex = 0; columnIndex < lastColumn; columnIndex++) {
            sheet.autoSizeColumn(columnIndex);
        }
    }

    private static void writeExcel(List<UserDTO> userDTOS, String excelFilePath) throws IOException {
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

        //Create a canvas
        XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
        //The first four defaults are 0, [0,4]: start from 0 column and 4 rows; [7,20]: width 7 cells, 20 expands down to 20 rows
        //Default width (14-8)*12
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 4, 7, 20);
        //Create a chart object
        XSSFChart chart = drawing.createChart(anchor);
        //Title
        chart.setTitleText("The top seven countries in the region");
        //Whether the title covers the chart
        chart.setTitleOverlay(false);

        //Legend position
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP_RIGHT);

        //CellRangeAddress (starting row number, ending row number, starting column number, ending column number)
        //Classification axis data,
        XDDFDataSource<String> countries = XDDFDataSourcesFactory.fromStringCellRange((XSSFSheet) sheet, new CellRangeAddress(0, 0, 0, 6));
        //Data 1,
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange((XSSFSheet) sheet, new CellRangeAddress(1, 1, 0, 6));
        //XDDFChartData data = chart.createData(ChartTypes.PIE3D, null, null);
        XDDFChartData data = chart.createData(ChartTypes.PIE, null, null);
        //Set to variable color
        data.setVaryColors(true);
        //Chart load data
        data.addSeries(countries, values);

        //Draw
        chart.plot(data);

        // Write output to excel file
        FileOutputStream fileOut = new FileOutputStream(excelFilePath);
        workbook.write(fileOut);
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

}
