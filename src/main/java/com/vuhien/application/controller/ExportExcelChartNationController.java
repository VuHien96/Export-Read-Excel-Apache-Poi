package com.vuhien.application.controller;

import com.vuhien.application.entity.Nation;
import com.vuhien.application.model.dto.UserDTO;
import com.vuhien.application.service.NationService;
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
public class ExportExcelChartNationController {

    @Autowired
    private NationService nationService;

    public static final int COLUMN_INDEX_ID = 0;
    public static final int COLUMN_INDEX_COUNTRY = 1;
    public static final int COLUMN_INDEX_POPULATION = 2;
    public final String excelFilePath = "D:/nation.xlsx";

    @GetMapping("/api/nations")
    public ResponseEntity<Object> getListNations() throws IOException {
        List<Nation> nations = nationService.getListNations();
        writeExcel(nations, excelFilePath);
        return ResponseEntity.ok(nations);
    }

    private static CellStyle createCellStyleHeader(Sheet sheet) {
        // Tạo font
        Font font = sheet.getWorkbook().createFont();
        font.setFontName("Times New Roman");
        font.setBold(true);
        font.setFontHeightInPoints((short) 18);
        font.setColor(IndexedColors.GREY_50_PERCENT.getIndex());

        // Tạo cellStyle
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

    private static void writeHeader(Sheet sheet, List<Nation> nations, int rowIndex) {
        CellStyle cellStyle = createCellStyleHeader(sheet);
        Row row = sheet.createRow(rowIndex);
        Cell cell;
        for (Nation nation : nations) {
            cell = row.createCell(rowIndex);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(nation.getCountry());
            rowIndex++;
        }
    }

    private static void writeExcel(List<Nation> nations, String excelFilePath) throws IOException {
        // Create Workbook
        Workbook workbook = getWorkbook(excelFilePath);
        int index = 0;
        for (int i = 0; i < nations.size() - 1; i++) {
            index++;
        }
        //Create sheet
        Sheet sheet = workbook.createSheet("Danh sách quốc gia");
        int rowIndex = 0;
        writeHeader(sheet, nations, rowIndex);

        rowIndex++;
        writeNation(sheet, nations);

        //Create a canvas
        XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
        //The first four defaults are 0, [0,4]: start from 0 column and 4 rows; [7,20]: width 7 cells, 20 expands down to 20 rows
        //Default width (14-8)*12
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 4, 7, 20);
        //Create a chart object
        XSSFChart chart = drawing.createChart(anchor);

        //Title
        chart.setTitleText("Dân số các nước");
        //Whether the title covers the chart
        chart.setTitleOverlay(false);
        //Legend position
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP_RIGHT);

        //CellRangeAddress (starting row number, ending row number, starting column number, ending column number)
        //Classification axis data,
        XDDFDataSource<String> countries = XDDFDataSourcesFactory.fromStringCellRange((XSSFSheet) sheet, new CellRangeAddress(0, 0, 0, index));
        //Data 1,
        XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange((XSSFSheet) sheet, new CellRangeAddress(1, 1, 0, index));
//        XDDFChartData data = chart.createData(ChartTypes.PIE3D, null, null);
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

    private static void writeNation(Sheet sheet, List<Nation> nations) {
        Cell cell;
        CellStyle cellStyle = createCellStyleHeader(sheet);
        int rowIndex = 1;
        Row row = sheet.createRow(rowIndex);
        rowIndex = 0;
        for (Nation nation : nations) {
            cell = row.createCell(rowIndex);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(nation.getPopulation());
            rowIndex++;
        }
    }
}
