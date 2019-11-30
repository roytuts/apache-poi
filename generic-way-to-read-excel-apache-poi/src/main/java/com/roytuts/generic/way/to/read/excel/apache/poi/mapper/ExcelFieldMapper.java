package com.roytuts.generic.way.to.read.excel.apache.poi.mapper;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

import com.roytuts.generic.way.to.read.excel.apache.poi.enums.FieldType;
import com.roytuts.generic.way.to.read.excel.apache.poi.model.ExcelField;

public final class ExcelFieldMapper {

	final static DateTimeFormatter dtf = DateTimeFormatter.ofPattern("dd-MM-yyyy");

	public static <T> List<T> getPojos(List<ExcelField[]> excelFields, Class<T> clazz) {

		List<T> list = new ArrayList<>();
		excelFields.forEach(evc -> {

			T t = null;

			try {
				t = clazz.getConstructor().newInstance();
			} catch (InstantiationException | IllegalAccessException | IllegalArgumentException
					| InvocationTargetException | NoSuchMethodException | SecurityException e1) {
				e1.printStackTrace();
			}

			Class<? extends Object> classz = t.getClass();

			for (int i = 0; i < evc.length; i++) {

				for (Field field : classz.getDeclaredFields()) {
					field.setAccessible(true);

					if (evc[i].getPojoAttribute().equalsIgnoreCase(field.getName())) {

						try {
							if (FieldType.STRING.getValue().equalsIgnoreCase(evc[i].getExcelColType())) {
								field.set(t, evc[i].getExcelValue());
							} else if (FieldType.DOUBLE.getValue().equalsIgnoreCase(evc[i].getExcelColType())) {
								field.set(t, Double.valueOf(evc[i].getExcelValue()));
							} else if (FieldType.INTEGER.getValue().equalsIgnoreCase(evc[i].getExcelColType())) {
								field.set(t, Double.valueOf(evc[i].getExcelValue()).intValue());
							} else if (FieldType.DATE.getValue().equalsIgnoreCase(evc[i].getExcelColType())) {
								field.set(t, LocalDate.parse(evc[i].getExcelValue(), dtf));
							}
						} catch (IllegalArgumentException | IllegalAccessException e) {
							e.printStackTrace();
						}

						break;
					}
				}
			}

			list.add(t);
		});

		return list;
	}

}
