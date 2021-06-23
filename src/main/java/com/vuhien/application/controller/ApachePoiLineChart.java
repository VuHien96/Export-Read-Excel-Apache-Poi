package com.vuhien.application.controller;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Chart;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.PresetLineDash;
import org.apache.poi.xddf.usermodel.XDDFLineProperties;
import org.apache.poi.xddf.usermodel.XDDFPresetLineDash;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;

import java.io.FileOutputStream;
import java.io.IOException;

public class ApachePoiLineChart {
    public static void main(String[] args) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        String sheetName = "Sheet1";
        FileOutputStream fileOut = null;
        try {

            Row row;
            Cell cell;
            XSSFDrawing drawing;
            XSSFClientAnchor anchor;
            XSSFChart chart;
            XDDFChartLegend legend;
            XDDFDataSource<String> countries;


            XSSFSheet sheet = wb.createSheet(sheetName);
            //First line, country name
            row = sheet.createRow(0);
            cell = row.createCell(0);
            cell.setCellValue("Russia");
            cell = row.createCell(1);
            cell.setCellValue("Canada");
            cell = row.createCell(2);
            cell.setCellValue("United States");
            cell = row.createCell(3);
            cell.setCellValue("China");
            cell = row.createCell(4);
            cell.setCellValue("Brazil");
            cell = row.createCell(5);
            cell.setCellValue("Australia");
            cell = row.createCell(6);
            cell.setCellValue("India");
            // second line, rural area
            row = sheet.createRow(1);
            cell = row.createCell(0);
            cell.setCellValue(17098242);
            cell = row.createCell(1);
            cell.setCellValue(9984670);
            cell = row.createCell(2);
            cell.setCellValue(9826675);
            cell = row.createCell(3);
            cell.setCellValue(9596961);
            cell = row.createCell(4);
            cell.setCellValue(8514877);
            cell = row.createCell(5);
            cell.setCellValue(7741220);
            cell = row.createCell(6);
            cell.setCellValue(3287263);
            // Third line, rural population
            row = sheet.createRow(2);
            cell = row.createCell(0);
            cell.setCellValue(14590041);
            cell = row.createCell(1);
            cell.setCellValue(35151728);
            cell = row.createCell(2);
            cell.setCellValue(32993302);
            cell = row.createCell(3);
            cell.setCellValue(14362887);
            cell = row.createCell(4);
            cell.setCellValue(21172141);
            cell = row.createCell(5);
            cell.setCellValue(25335727);
            cell = row.createCell(6);
            cell.setCellValue(13724923);
            // The fourth line, the area is tied
            row = sheet.createRow(3);
            cell = row.createCell(0);
            cell.setCellValue(9435701.143);
            cell = row.createCell(1);
            cell.setCellValue(9435701.143);
            cell = row.createCell(2);
            cell.setCellValue(9435701.143);
            cell = row.createCell(3);
            cell.setCellValue(9435701.143);
            cell = row.createCell(4);
            cell.setCellValue(9435701.143);
            cell = row.createCell(5);
            cell.setCellValue(9435701.143);
            cell = row.createCell(6);
            cell.setCellValue(9435701.143);
            // fourth line, population tie
            row = sheet.createRow(4);
            cell = row.createCell(0);
            cell.setCellValue(22475821.29);
            cell = row.createCell(1);
            cell.setCellValue(22475821.29);
            cell = row.createCell(2);
            cell.setCellValue(22475821.29);
            cell = row.createCell(3);
            cell.setCellValue(22475821.29);
            cell = row.createCell(4);
            cell.setCellValue(22475821.29);
            cell = row.createCell(5);
            cell.setCellValue(22475821.29);
            cell = row.createCell(6);
            cell.setCellValue(22475821.29);

            //Create a canvas
            drawing = sheet.createDrawingPatriarch();
            //The first four default 0, [0,5]: start from 0 column and 5 rows; [7,26]: width 7 cells, 26 expands down to 26 rows
            //Default width (14-8)*12
            anchor = drawing.createAnchor(0, 0, 0, 0, 0, 5, 7, 26);
            //Create a chart object
            chart = drawing.createChart(anchor);
            //Title
            chart.setTitleText("The top seven countries in the region");
            //Title overwrite
            chart.setTitleOverlay(false);

            //Legend position
            legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.TOP);

            //Classification axis (X axis), title position
            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            bottomAxis.setTitle("Country");
            //Value (Y axis) axis, title position
            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
            leftAxis.setTitle("Area and population");

            //CellRangeAddress (starting row number, ending row number, starting column number, ending column number)
            //Categorized axis index (X axis) data, cell range position [0, 0] to [0, 6]
            countries = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(0, 0, 0, 6));
            // XDDFCategoryDataSource countries = XDDFDataSourcesFactory.fromArray(new String[] {"Russia","Canada","United States","China","Brazil","Australia","India"});
            //Data 1, cell range position [1, 0] to [1, 6]
            XDDFNumericalDataSource<Double> area = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, 6));
//			XDDFNumericalDataSource<Integer> area = XDDFDataSourcesFactory.fromArray(new Integer[] {17098242,9984670,9826675,9596961,8514877,7741220,3287263});

            //Data 1, cell range position [2, 0] to [2, 6]
            XDDFNumericalDataSource<Double> population = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(2, 2, 0, 6));

            //LINE: line chart,
            XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);

            //Chart load data, broken line 1
            XDDFLineChartData.Series series1 = (XDDFLineChartData.Series) data.addSeries(countries, area);
            //Line legend title
            series1.setTitle("Area", null);
            //Straight
            series1.setSmooth(false);
            //Set the mark size
            series1.setMarkerSize((short) 6);
            //Set the mark style, stars
            series1.setMarkerStyle(MarkerStyle.STAR);

            //Chart load data, broken line 2
            XDDFLineChartData.Series series2 = (XDDFLineChartData.Series) data.addSeries(countries, population);
            //Line legend title
            series2.setTitle("Population", null);
            //Curve
            series2.setSmooth(true);
            //Set the mark size
            series2.setMarkerSize((short) 6);
            //Set the mark style, square
            series2.setMarkerStyle(MarkerStyle.SQUARE);

            //Chart load data, average line 3
            //Data 1, cell range position [2, 0] to [2, 6]
            XDDFNumericalDataSource<Double> population3 = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(3, 3, 0, 6));
            XDDFLineChartData.Series series3 = (XDDFLineChartData.Series) data.addSeries(countries, population3);
            //Line legend title
            series3.setTitle("Average Area", null);
            //Straight
            series3.setSmooth(false);
            //Set the mark size
            //			series3.setMarkerSize((short) 3);
            //Set the mark style, square
            series3.setMarkerStyle(MarkerStyle.NONE);
            //LineChart
            //	        XDDFSolidFillProperties fill = new XDDFSolidFillProperties(XDDFColor.from(PresetColor.CHARTREUSE));
            XDDFLineProperties line = new XDDFLineProperties();
            //	        line.setFillProperties(fill);
            //	        line.setLineCap(LineCap.ROUND);
            line.setPresetDash(new XDDFPresetLineDash(PresetLineDash.DOT));//dotted line
            //	        XDDFShapeProperties shapeProperties = new XDDFShapeProperties();
            //	        shapeProperties.setLineProperties(line);
            //	        series3.setShapeProperties(shapeProperties);
            series3.setLineProperties(line);

            //Chart load data, average line 3
            //Data 1, cell range position [2, 0] to [2, 6]
            XDDFNumericalDataSource<Double> population4 = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(4, 4, 0, 6));
            XDDFLineChartData.Series series4 = (XDDFLineChartData.Series) data.addSeries(countries, population4);
            //Line legend title
            series4.setTitle("Population Average", null);
            //Straight
            series4.setSmooth(false);
            //Set the mark size
            //			series4.setMarkerSize((short) 3);
            //Set the mark style, square
            series4.setMarkerStyle(MarkerStyle.NONE);
            XDDFLineProperties line4 = new XDDFLineProperties();
            line4.setPresetDash(new XDDFPresetLineDash(PresetLineDash.DOT));//dotted line
            series4.setLineProperties(line);

            //Draw
            chart.plot(data);



            /* ******************************************************************************* */

            sheet = wb.createSheet("Sheet2");


            row = sheet.createRow(0);
            row.createCell(0);
            row.createCell(1).setCellValue("Bars");
            row.createCell(2).setCellValue("Lines");

            for (int r = 1; r < 7; r++) {
                row = sheet.createRow(r);
                cell = row.createCell(0);
                cell.setCellValue("C" + r);
                cell = row.createCell(1);
                cell.setCellValue(new java.util.Random().nextDouble());
                cell = row.createCell(2);
                cell.setCellValue(new java.util.Random().nextDouble() * 10d);
            }

            drawing = sheet.createDrawingPatriarch();
//        ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 4, 0, 11, 15);
            anchor = (XSSFClientAnchor) drawing.createAnchor(0, 0, 0, 0, 4, 0, 11, 15);

            Chart chart1 = drawing.createChart(anchor);

            CTChart ctChart = ((XSSFChart) chart1).getCTChart();
            CTPlotArea ctPlotArea = ctChart.getPlotArea();

            //the bar chart
            CTBarChart ctBarChart = ctPlotArea.addNewBarChart();
            CTBoolean ctBoolean = ctBarChart.addNewVaryColors();
            ctBoolean.setVal(true);
            ctBarChart.addNewBarDir().setVal(STBarDir.COL);

            //the bar series
            CTBarSer ctBarSer = ctBarChart.addNewSer();
            CTSerTx ctSerTx = ctBarSer.addNewTx();
            CTStrRef ctStrRef = ctSerTx.addNewStrRef();
            ctStrRef.setF("Sheet1!$B$1");
            ctBarSer.addNewIdx().setVal(0);
            CTAxDataSource cttAxDataSource = ctBarSer.addNewCat();
            ctStrRef = cttAxDataSource.addNewStrRef();
            ctStrRef.setF("Sheet1!$A$2:$A$7");
            CTNumDataSource ctNumDataSource = ctBarSer.addNewVal();
            CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();
            ctNumRef.setF("Sheet1!$B$2:$B$7");

            //at least the border lines in Libreoffice Calc ;-)
            ctBarSer.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[]{0, 0, 0});

            //telling the BarChart that it has axes and giving them Ids
            ctBarChart.addNewAxId().setVal(123456); //cat axis 1 (bars)
            ctBarChart.addNewAxId().setVal(123457); //val axis 1 (left)

            //the line chart
            CTLineChart ctLineChart = ctPlotArea.addNewLineChart();
            ctBoolean = ctLineChart.addNewVaryColors();
            ctBoolean.setVal(true);

            //the line series
            CTLineSer ctLineSer = ctLineChart.addNewSer();
            ctSerTx = ctLineSer.addNewTx();
            ctStrRef = ctSerTx.addNewStrRef();
            ctStrRef.setF("Sheet1!$C$1");
            ctLineSer.addNewIdx().setVal(1);
            cttAxDataSource = ctLineSer.addNewCat();
            ctStrRef = cttAxDataSource.addNewStrRef();
            ctStrRef.setF("Sheet1!$A$2:$A$7");
            ctNumDataSource = ctLineSer.addNewVal();
            ctNumRef = ctNumDataSource.addNewNumRef();
            ctNumRef.setF("Sheet1!$C$2:$C$7");

            //at least the border lines in Libreoffice Calc ;-)
            ctLineSer.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[]{0, 0, 0});

            //telling the LineChart that it has axes and giving them Ids
            ctLineChart.addNewAxId().setVal(123458); //cat axis 2 (lines)
            ctLineChart.addNewAxId().setVal(123459); //val axis 2 (right)

            //cat axis 1 (bars)
            CTCatAx ctCatAx = ctPlotArea.addNewCatAx();
            ctCatAx.addNewAxId().setVal(123456); //id of the cat axis
            CTScaling ctScaling = ctCatAx.addNewScaling();
            ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
            ctCatAx.addNewDelete().setVal(false);
            ctCatAx.addNewAxPos().setVal(STAxPos.B);
            ctCatAx.addNewCrossAx().setVal(123457); //id of the val axis
            ctCatAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

            //val axis 1 (left)
            CTValAx ctValAx = ctPlotArea.addNewValAx();
            ctValAx.addNewAxId().setVal(123457); //id of the val axis
            ctScaling = ctValAx.addNewScaling();
            ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
            ctValAx.addNewDelete().setVal(false);
            ctValAx.addNewAxPos().setVal(STAxPos.L);
            ctValAx.addNewCrossAx().setVal(123456); //id of the cat axis
            ctValAx.addNewCrosses().setVal(STCrosses.AUTO_ZERO); //this val axis crosses the cat axis at zero
            ctValAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

            //cat axis 2 (lines)
            ctCatAx = ctPlotArea.addNewCatAx();
            ctCatAx.addNewAxId().setVal(123458); //id of the cat axis
            ctScaling = ctCatAx.addNewScaling();
            ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
            ctCatAx.addNewDelete().setVal(true); //this cat axis is deleted
            ctCatAx.addNewAxPos().setVal(STAxPos.B);
            ctCatAx.addNewCrossAx().setVal(123459); //id of the val axis
            ctCatAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

            //val axis 2 (right)
            ctValAx = ctPlotArea.addNewValAx();
            ctValAx.addNewAxId().setVal(123459); //id of the val axis
            ctScaling = ctValAx.addNewScaling();
            ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
            ctValAx.addNewDelete().setVal(false);
            ctValAx.addNewAxPos().setVal(STAxPos.R);
            ctValAx.addNewCrossAx().setVal(123458); //id of the cat axis
            ctValAx.addNewCrosses().setVal(STCrosses.MAX); //this val axis crosses the cat axis at max value
            ctValAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

            //legend
            CTLegend ctLegend = ctChart.addNewLegend();
            ctLegend.addNewLegendPos().setVal(STLegendPos.B);
            ctLegend.addNewOverlay().setVal(false);

            /* ************************************************************************************************* */

            sheet = wb.createSheet("Sheet3");

            row = sheet.createRow(0);
            cell = row.createCell(0);
            cell.setCellValue("Russia");
            cell = row.createCell(1);
            cell.setCellValue("Canada");
            cell = row.createCell(2);
            cell.setCellValue("United States");
            cell = row.createCell(3);
            cell.setCellValue("China");
            cell = row.createCell(4);
            cell.setCellValue("Brazil");
            cell = row.createCell(5);
            cell.setCellValue("Australia");
            cell = row.createCell(6);
            cell.setCellValue("India");
            // second line, rural area
            row = sheet.createRow(1);
            cell = row.createCell(0);
            cell.setCellValue(17098242);
            cell = row.createCell(1);
            cell.setCellValue(9984670);
            cell = row.createCell(2);
            cell.setCellValue(9826675);
            cell = row.createCell(3);
            cell.setCellValue(9596961);
            cell = row.createCell(4);
            cell.setCellValue(8514877);
            cell = row.createCell(5);
            cell.setCellValue(7741220);
            cell = row.createCell(6);
            cell.setCellValue(3287263);
            // Third line, rural population
            row = sheet.createRow(2);
            cell = row.createCell(0);
            cell.setCellValue(14590041);
            cell = row.createCell(1);
            cell.setCellValue(35151728);
            cell = row.createCell(2);
            cell.setCellValue(32993302);
            cell = row.createCell(3);
            cell.setCellValue(14362887);
            cell = row.createCell(4);
            cell.setCellValue(21172141);
            cell = row.createCell(5);
            cell.setCellValue(25335727);
            cell = row.createCell(6);
            cell.setCellValue(13724923);

            //Create a canvas
            drawing = sheet.createDrawingPatriarch();
            //The first four defaults are 0, [0,4]: start from 0 column and 4 rows; [7,20]: width 7 cells, 20 expands down to 20 rows
            //Default width (14-8)*12
            anchor = drawing.createAnchor(0, 0, 0, 0, 0, 4, 7, 20);
            //Create a chart object
            chart = drawing.createChart(anchor);
            //Title
            chart.setTitleText("The top seven countries in the region");
            //Whether the title covers the chart
            chart.setTitleOverlay(false);

            //Legend position
            legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.TOP_RIGHT);

            //CellRangeAddress (starting row number, ending row number, starting column number, ending column number)
            //Classification axis data,
            countries = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(0, 0, 0, 6));
            //Data 1,
            XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, 6));
            //XDDFChartData data = chart.createData(ChartTypes.PIE3D, null, null);
            XDDFChartData xddfChartData = chart.createData(ChartTypes.PIE, null, null);
            //Set to variable color
            xddfChartData.setVaryColors(true);
            //Chart load data
            xddfChartData.addSeries(countries, values);

            //Draw
            chart.plot(xddfChartData);


            // Write output to excel file
            String filename = "D:/z.xlsx";
            fileOut = new FileOutputStream(filename);
            wb.write(fileOut);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            wb.close();
            if (fileOut != null) {
                fileOut.close();
            }
        }


    }
}
