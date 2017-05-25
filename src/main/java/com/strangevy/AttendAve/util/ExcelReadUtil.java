package com.strangevy.AttendAve.util;


import java.io.FileInputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * 读取Excel辅助类
 */
public final class ExcelReadUtil {
	private ExcelReadUtil() {
	}


	/**
	 * 获取Excel数据,返回List<String[]>
	 * 
	 * @param sheetNumber 读取工作表的下标(从1开始).可有可无,默认读取所有表单.
	 */
	public static final List<String[]> excelToArrayList(String filePath, int... sheetNumber) throws Exception {
		List<String[]> resultList = null;
		InputStream is = null;
		try {
			is = new FileInputStream(filePath);
			resultList = excelToArrayList(is, sheetNumber);
		} catch (Exception e) {
			throw e;
		} finally {
			if (is != null) {
				is.close();
			}
		}
		return resultList;
	}


	public  static final  List<String[]> excelToList(InputStream is) throws  Exception{
		ArrayList<String[]> resultList = new ArrayList<String[]>();
		Workbook wb = null;

		wb = WorkbookFactory.create(is);
		Sheet sheet = wb.getSheetAt(0);

	//	System.out.println("lastRowNum--------:"+rows);
		int rows1 = sheet.getPhysicalNumberOfRows();
	//	System.out.println("getPhysicalNumberOfRows-------:"+rows1);
		int column = sheet.getRow(0).getPhysicalNumberOfCells();
	//	System.out.println("column-------------"+column);

		for(int i = 0;i<rows1;i++){
			String[] object = new String[column];
			for(int j=0;j<column;j++) {
				if (null != sheet.getRow(i).getCell(j)) {
					if (sheet.getRow(i).getCell(j).getCellType() == Cell.CELL_TYPE_NUMERIC) {
						if (DateUtil.isCellDateFormatted(sheet.getRow(i).getCell(j))) {
							SimpleDateFormat format = new SimpleDateFormat("yyyy-MM-dd");
							String date = format.format(sheet.getRow(i).getCell(j).getDateCellValue());
							object[j] = date.trim();

						} else {
							sheet.getRow(i).getCell(j).setCellType(Cell.CELL_TYPE_STRING);
							//System.out.println("工时：---------------"+sheet.getRow(i).getCell(j).getStringCellValue());
							object[j] = sheet.getRow(i).getCell(j).getStringCellValue().trim();
						}
					} else {
						sheet.getRow(i).getCell(j).setCellType(Cell.CELL_TYPE_STRING);
						//System.out.println("i的值是:"+i+"j的值是："+j+"type是："+sheet.getRow(i).getCell(j).getCellType());
						object[j] = sheet.getRow(i).getCell(j).getStringCellValue().trim();

					}

				}else{
					object[j] = "";
				}
			}
			resultList.add(object);

		}

		return  resultList;
	}

	/**
	 * 获取Excel数据,返回List<String[]>;
	 * 
	 * @param //sheetNumber 读取工作表的下标(从1开始).可有可无,默认读取所有表单.
	 */
	public static final List<String[]> excelToArrayList(InputStream is, int... sheetNumber) throws Exception {
		ArrayList<String[]> resultList = new ArrayList<String[]>();
		Workbook wb = null;
		try {
			wb = WorkbookFactory.create(is);
			if (sheetNumber == null || sheetNumber.length < 1) {
				int sheetCount = wb.getNumberOfSheets();// 得到所有Excel中页的列表.
				sheetNumber = new int[sheetCount];
				for (int i = 0; i < sheetNumber.length; i++) {
					sheetNumber[i] = i + 1;
				}
			}
			Sheet sheet = null;
			for (int k = 0; k < sheetNumber.length; k++) {// 循环工作表
				sheet = wb.getSheetAt(sheetNumber[k] - 1);
				int rsRows = sheet.getRow(0) == null ? -1 : sheet.getLastRowNum();
				//假设第一行是表头
				int headCount = sheet.getRow(0).getLastCellNum();
				for (int i = 0; i <= rsRows; i++) {// 循环行
					Row row = sheet.getRow(i);
					if (row != null) {
						int cellCount = row.getLastCellNum();
						if (cellCount > 0) {
							/*String[] objects = new String[cellCount];
							for (int j = 0; j < cellCount; j++) {// 读取单元格
								objects[j] = getCellValue(row.getCell(j));
							}
							resultList.add(objects);*/
							String[] objects = new String[headCount];
							for (int j = 0; j < headCount; j++) {// 读取单元格
								objects[j] = getCellValue(row.getCell(j));
							}
							resultList.add(objects);
						}
					}
				}
			}
		} catch (Exception e) {
			throw e;
		}
		return resultList;
	}

	/**
	 * 读取Excel中的数据.将这些数据放入到一个三维数组中.
	 * 
	 * @param filePath 文件路径.
	 * @param sheetNumber 读取工作表的下标(从1开始).可有可无,默认读取所有表单.
	 * @deprecated:读取Excel中的数据将它放入到ArrayList数组中(此为三维数组).
	 */
	public static final ArrayList<ArrayList<Object>> readExcel(String filePath, int... sheetNumber) throws Exception {
		ArrayList<ArrayList<Object>> subdata = null;
		InputStream is = null;
		try {
			is = new FileInputStream(filePath);
			subdata = readExcel(is, sheetNumber);
		} catch (Exception e) {
			throw e;
		} finally {
			if (is != null) {
				is.close();
			}
		}
		return subdata;
	}


	/**
	 * 读取Excel中的数据.将这些数据放入到一个三维数组中.
	 * 
	 * @param sheetNumber 读取工作表的下标(从1开始).可有可无,默认读取所有表单.
	 */
	public static final ArrayList<ArrayList<Object>> readExcel(InputStream is, int... sheetNumber) throws Exception {
		ArrayList<ArrayList<Object>> subdata = new ArrayList<ArrayList<Object>>();
		Workbook wb = null;
		wb = WorkbookFactory.create(is);
		if (sheetNumber == null || sheetNumber.length < 1) {
			int sheetCount = wb.getNumberOfSheets();// 得到所有Excel中页的列表.
			sheetNumber = new int[sheetCount];
			for (int i = 0; i < sheetNumber.length; i++) {
				sheetNumber[i] = i + 1;
			}
		}
		Sheet sheet = null;
		ArrayList<ArrayList<String>> alList = null;
		ArrayList<String> tablenames = null;
		ArrayList<Object> tableAndContents = null;
		for (int a = 0; a < sheetNumber.length; a++) {
			alList = new ArrayList<ArrayList<String>>();
			tablenames = new ArrayList<String>();
			tableAndContents = new ArrayList<Object>();
			String tablename = wb.getSheetName(sheetNumber[a] - 1).trim();
			int b = 0;
			sheet = wb.getSheetAt(sheetNumber[a] - 1);
			int rsRows = sheet.getRow(0) == null ? -1 : sheet.getLastRowNum();
			for (int i = 1; i <= rsRows; i++) {
				ArrayList<String> al = new ArrayList<String>();
				Row row = sheet.getRow(i);
				int cellCount = row.getLastCellNum();
				for (int j = 0; j < cellCount; j++) {
					// 通用的获取cell值的方式,返回字符串
					String strc00 = getCellValue(row.getCell(j));
					// 获得cell具体类型值的方式得到内容.
					al.add(j, strc00);
				}
				alList.add(b, al);
				b++;
			}
			tablenames.add(tablename);
			tableAndContents.add(0, tablenames);
			tableAndContents.add(1, alList);
			subdata.add(a, tableAndContents);
		}
		return subdata;
	}

	/** 根据Cell类型设置数据 */
	private static String getCellValue(Cell cell) {
		String cellvalue = "";
		if (cell != null) {
			// 判断当前Cell的Type
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_BLANK:
				cellvalue = "";
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				cellvalue = Boolean.toString(cell.getBooleanCellValue());
				break;
			// 数值
			case Cell.CELL_TYPE_NUMERIC:
				if ( DateUtil.isCellDateFormatted(cell)) {
					cellvalue = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(cell.getDateCellValue());
				} else {
					cell.setCellType(Cell.CELL_TYPE_STRING);
					String temp = cell.getStringCellValue();
					// 判断是否包含小数点，如果不含小数点，则以字符串读取，如果含小数点，则转换为Double类型的字符串
					if (temp.indexOf(".") > -1) {
						cellvalue = String.valueOf(new Double(temp)).trim();
					} else {
						cellvalue = temp.trim();
					}
				}
				break;
			case Cell.CELL_TYPE_STRING:
				cellvalue = cell.getStringCellValue().trim();
				break;
			case Cell.CELL_TYPE_ERROR:
				cellvalue = "";
				break;
			case Cell.CELL_TYPE_FORMULA:
				cell.setCellType(Cell.CELL_TYPE_STRING);
				cellvalue = cell.getStringCellValue();
				if (cellvalue != null) {
					cellvalue = cellvalue.replaceAll("#N/A", "").trim();
				}
				break;
			default:
				cellvalue = "";
				break;
			}
		} else {
			cellvalue = "";
		}
		return cellvalue;
	}

	/**
	 * 判断是否空行
	 * @param rowData
	 * @return
	 */
	public static boolean isBlankRow(String[] rowData){
		boolean isBlank = true;
		if(rowData != null) {
			for (int i = 0; i < rowData.length; i++) {
				if (null!=rowData[i]&&!"".equals(rowData[i].trim())) {
					isBlank = false;
					break;
				}
			}
		}
		return isBlank;
	}

}
