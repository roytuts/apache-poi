package com.roytuts.apache.poi.excel.stacked.bar.column.chart;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.BarDirection;
import org.apache.poi.xddf.usermodel.chart.BarGrouping;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ApachePoiStackedBarColumnChart {

	public static void main(String[] args) throws IOException {
		stackedBarChart();
	}

	public static void stackedBarChart() throws FileNotFoundException, IOException {
		try (XSSFWorkbook wb = new XSSFWorkbook()) {

			String sheetName = "CountryStackedBarChart";// "CountryStackedColumnChart";

			XSSFSheet sheet = wb.createSheet(sheetName);

			// Country Names
			Row row = sheet.createRow((short) 0);

			Cell cell = row.createCell((short) 0);
			cell.setCellValue("Russia");

			cell = row.createCell((short) 1);
			cell.setCellValue("Canada");

			cell = row.createCell((short) 2);
			cell.setCellValue("USA");

			cell = row.createCell((short) 3);
			cell.setCellValue("China");

			cell = row.createCell((short) 4);
			cell.setCellValue("Brazil");

			cell = row.createCell((short) 5);
			cell.setCellValue("Australia");

			cell = row.createCell((short) 6);
			cell.setCellValue("India");

			// Country Area
			row = sheet.createRow((short) 1);

			cell = row.createCell((short) 0);
			cell.setCellValue(17098242);

			cell = row.createCell((short) 1);
			cell.setCellValue(9984670);

			cell = row.createCell((short) 2);
			cell.setCellValue(9826675);

			cell = row.createCell((short) 3);
			cell.setCellValue(9596961);

			cell = row.createCell((short) 4);
			cell.setCellValue(8514877);

			cell = row.createCell((short) 5);
			cell.setCellValue(7741220);

			cell = row.createCell((short) 6);
			cell.setCellValue(3287263);

			// Country Population
			row = sheet.createRow((short) 2);

			cell = row.createCell((short) 0);
			cell.setCellValue(14590041);

			cell = row.createCell((short) 1);
			cell.setCellValue(35151728);

			cell = row.createCell((short) 2);
			cell.setCellValue(32993302);

			cell = row.createCell((short) 3);
			cell.setCellValue(14362887);

			cell = row.createCell((short) 4);
			cell.setCellValue(21172141);

			cell = row.createCell((short) 5);
			cell.setCellValue(25335727);

			cell = row.createCell((short) 6);
			cell.setCellValue(13724923);

			XSSFDrawing drawing = sheet.createDrawingPatriarch();
			XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 4, 7, 26);

			XSSFChart chart = drawing.createChart(anchor);
			chart.setTitleText("Area-wise Top Seven Countries");
			chart.setTitleOverlay(false);

			XDDFChartLegend legend = chart.getOrAddLegend();
			legend.setPosition(LegendPosition.TOP_RIGHT);

			XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
			bottomAxis.setTitle("Country");
			XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
			leftAxis.setTitle("Area & Population");

			XDDFDataSource<String> countries = XDDFDataSourcesFactory.fromStringCellRange(sheet,
					new CellRangeAddress(0, 0, 0, 6));

			XDDFNumericalDataSource<Double> area = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
					new CellRangeAddress(1, 1, 0, 6));

			XDDFNumericalDataSource<Double> population = XDDFDataSourcesFactory.fromNumericCellRange(sheet,
					new CellRangeAddress(2, 2, 0, 6));

			XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
			XDDFChartData.Series series1 = data.addSeries(countries, area);
			series1.setTitle("Area", null);
			XDDFChartData.Series series2 = data.addSeries(countries, population);
			series2.setTitle("Population", null);
			data.setVaryColors(true);
			chart.plot(data);

			// correcting the overlap so bars really are stacked and not side by side
			chart.getCTChart().getPlotArea().getBarChartArray(0).addNewOverlap().setVal((byte) 100);

			// in order to transform a bar chart into a column chart, you just need to
			// change the bar direction
			XDDFBarChartData bar = (XDDFBarChartData) data;
			bar.setBarGrouping(BarGrouping.STACKED);
			bar.setBarDirection(BarDirection.BAR);
			// bar.setBarDirection(BarDirection.COL);

			// Write output to an excel file
			String filename = "bar-chart-top-seven-countries.xlsx";// "column-chart-top-seven-countries.xlsx";
			try (FileOutputStream fileOut = new FileOutputStream(filename)) {
				wb.write(fileOut);
			}
		}
	}

}
