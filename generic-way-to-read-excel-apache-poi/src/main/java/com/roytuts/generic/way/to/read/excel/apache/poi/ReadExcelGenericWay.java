package com.roytuts.generic.way.to.read.excel.apache.poi;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.fasterxml.jackson.databind.exc.InvalidFormatException;
import com.roytuts.generic.way.to.read.excel.apache.poi.enums.ExcelSection;
import com.roytuts.generic.way.to.read.excel.apache.poi.mapper.ExcelFieldMapper;
import com.roytuts.generic.way.to.read.excel.apache.poi.model.ExcelField;
import com.roytuts.generic.way.to.read.excel.apache.poi.model.Order;
import com.roytuts.generic.way.to.read.excel.apache.poi.model.Profit;
import com.roytuts.generic.way.to.read.excel.apache.poi.reader.ExcelFileReader;

public class ReadExcelGenericWay {

	public static void main(String[] args) throws InvalidFormatException {
		try {
			Workbook workbook = ExcelFileReader.readExcel("C:/jee_workspace/order-profit.xlsx");
			Sheet sheet = workbook.getSheetAt(0);
			
			Map<String, List<ExcelField[]>> excelRowValuesMap = ExcelFileReader.getExcelRowValues(sheet);
			
			excelRowValuesMap.forEach((section, rows) -> {
				System.out.println(section);
				System.out.println("==============");
				boolean headerPrint = true;
				for (ExcelField[] evc : rows) {
					if (headerPrint) {
						for (int j = 0; j < evc.length; j++) {
							System.out.print(evc[j].getExcelHeader() + "\t");
						}
						System.out.println();
						System.out.println(
								"------------------------------------------------------------------------------------");
						System.out.println();
						headerPrint = false;
					}
					for (int j = 0; j < evc.length; j++) {
						System.out.print(evc[j].getExcelValue() + "\t");
					}
					System.out.println();
				}
				System.out.println();
			});
			
			List<Order> orders = ExcelFieldMapper.getPojos(excelRowValuesMap.get(ExcelSection.ORDERS.getValue()),
					Order.class);

			List<Profit> profits = ExcelFieldMapper.getPojos(excelRowValuesMap.get(ExcelSection.PROFIT.getValue()),
					Profit.class);

			/*
			 * orders.forEach(o -> { System.out.println(o.getItem()); });
			 * 
			 * profits.forEach(p -> { System.out.println(p.getProfit());
			 * System.out.println(p.getDate()); });
			 */
		} catch (EncryptedDocumentException | IOException e) {
			e.printStackTrace();
		}
	}

}
