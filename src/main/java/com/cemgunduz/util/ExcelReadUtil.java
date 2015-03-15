package com.cemgunduz.util;

import com.cemgunduz.model.ExcelReadError;
import com.cemgunduz.model.ExcelReadErrorType;
import com.cemgunduz.model.ExcelReadResponse;
import com.cemgunduz.model.ExcelSheet;
import com.cemgunduz.model.annotation.ExcelMapping;
import com.cemgunduz.model.annotation.ExcelMappingKey;
import com.cemgunduz.model.annotation.ExcelMappings;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.lang.reflect.Field;
import java.util.*;

/**
 * Created by cgunduz on 3/12/14.
 */

public class ExcelReadUtil {

	//private static Logger logger = LoggerFactory.getLogger(com.netas.aydes.util.ExcelReadUtil.class);

	/**
	 * Get a list of sheets, wrapped as a excel sheet object from spreadsheet
	 *
	 * @param inputStream
	 * @return
	 * @throws java.io.IOException
	 * @throws InvalidFormatException
	 */
	public static List<ExcelSheet> getContentMatriceFromExcelFile(InputStream inputStream) throws IOException, InvalidFormatException {

		Workbook workbook = WorkbookFactory.create(inputStream);
		int totalSheets = workbook.getNumberOfSheets();

		List<ExcelSheet> contentMatrice = new ArrayList<ExcelSheet>();

		for(int i = 0; totalSheets > i; i++){
			contentMatrice.add(getContentMatrixBySheet(workbook.getSheetAt(i)));
		}

		return contentMatrice;
	}

	public static List<ExcelSheet> getContentMatriceFromExcelFile(File file) throws IOException, InvalidFormatException {

		FileInputStream fileInputStream = new FileInputStream(file);

		List<ExcelSheet> excelSheets;
		try{
			excelSheets = getContentMatriceFromExcelFile(fileInputStream);
		}
		finally{
			fileInputStream.close();
		}

		return excelSheets;
	}

	/**
	 * Get entity list from spreadsheet file with each desired variable mapped with excelMappable
	 * Default sheet no is used as 0 (first)
	 *
	 * @param inputStream
	 * @param T
	 * @param <T>
	 * @return
	 * @throws java.io.IOException
	 * @throws InvalidFormatException
	 */
	public static <T> ExcelReadResponse<T> getEntityListFromExcelFile(InputStream inputStream, Class T) throws IOException, InvalidFormatException {

		return getEntityListFromExcelFile(inputStream, 0, true, T);
	}

	public static <T> ExcelReadResponse<T> getEntityListFromExcelFile(File file, Class T) throws IOException, InvalidFormatException {

		FileInputStream fileInputStream = new FileInputStream(file);

		ExcelReadResponse<T> excelReadResponse;
		try{
			excelReadResponse = getEntityListFromExcelFile(fileInputStream, 0, true, T);
		}
		finally{
			fileInputStream.close();
		}

		return excelReadResponse;

	}

	/**
	 * Get entity list from spreadsheet file with each desired variable mapped with excelMappable
	 *
	 * @param inputStream
	 * @param sheetNo
	 * @param T
	 * @param <T>
	 * @return
	 * @throws java.io.IOException
	 * @throws InvalidFormatException
	 */
	public static <T> ExcelReadResponse<T> getEntityListFromExcelFile(InputStream inputStream, int sheetNo, boolean useHeader, Class T) throws IOException, InvalidFormatException {

		ExcelReadResponse<T> excelReadResponse = getContentMatrixBySheetAsReadResponse(inputStream, useHeader, T);
		List<ExcelSheet> excelSheetList = excelReadResponse.getExcelSheetList();

		int totalRemovedSheets = 0;
		for(Integer emptySheetNo : excelReadResponse.getEmptySheets()){
			if(emptySheetNo < sheetNo) totalRemovedSheets++;
			else if(emptySheetNo == sheetNo){
                excelReadResponse.setSuccessful(false);
				excelReadResponse.getExcelReadErrorList().add(new ExcelReadError(ExcelReadErrorType.EMPTY_SHEET));
				return excelReadResponse;
			}
		}

		sheetNo -= totalRemovedSheets;
		ExcelSheet excelSheet = excelSheetList.get(sheetNo);

		List<T> entityList = new ArrayList<T>();
		for(int i = 0; i < excelSheet.getTotalRows(); i++){
			// Skip faulty rows
			boolean isFaultyRow = false;
			for(ExcelReadError excelReadError : excelReadResponse.getExcelReadErrorList())
				if(excelReadError.getRowNumber().equals(i)){
					isFaultyRow = true;
					break;
				}

			// Skip header
			if(isFaultyRow || (useHeader && i == 0)) continue;

			int coloumnNo = -1;
			try{
				entityList.add((T) instantiateObject(excelSheet, i, T));
			}
			catch (IllegalStateException e){
				ExcelReadError excelReadError = new ExcelReadError(ExcelReadErrorType.ILLEGAL_TYPE, i, coloumnNo);
				excelReadResponse.getExcelReadErrorList().add(excelReadError);
				excelReadResponse.setSuccessful(false);
			}
			catch (IllegalAccessException e){
				ExcelReadError excelReadError = new ExcelReadError(ExcelReadErrorType.UNEXPECTED, i, coloumnNo);
				excelReadResponse.getExcelReadErrorList().add(excelReadError);
				excelReadResponse.setSuccessful(false);
			}
			catch (InstantiationException e){
				ExcelReadError excelReadError = new ExcelReadError(ExcelReadErrorType.CLASS_INSTANTIATION_EXCEPTION, i, coloumnNo);
				excelReadResponse.getExcelReadErrorList().add(excelReadError);
				excelReadResponse.setSuccessful(false);
				return excelReadResponse;
			}

		}

		excelReadResponse.setEntityList(entityList);
		return excelReadResponse;
	}

	private static Object instantiateObject(ExcelSheet excelSheet, int rowNum, Class type) throws IllegalAccessException, InstantiationException, IllegalStateException {

		return instantiateObject(excelSheet, rowNum, type, "default");
	}

	private static Object instantiateObject(ExcelSheet excelSheet, int rowNum, Class type, String key) throws IllegalAccessException, InstantiationException, IllegalStateException {

		int coloumnNo = -1;
		Object reflectionObject = type.newInstance();
		for(Field field : type.getDeclaredFields()){
			field.setAccessible(true);

			Class fieldType = field.getType();

			if(!isASupportedKnownClass(fieldType)){
				ExcelMappingKey excelMappingKey = field.getAnnotation(ExcelMappingKey.class);
				if(excelMappingKey == null || !excelMappingKey.lock().equals(key)) continue;

				field.set(reflectionObject, instantiateObject(excelSheet, rowNum, fieldType, excelMappingKey.key()));
			}
			else{
				ExcelMapping excelMapping = null;
				ExcelMappings excelMappings = field.getAnnotation(ExcelMappings.class);
				if(excelMappings != null){
					for(ExcelMapping mapping : excelMappings.list())
						if(mapping.key().equals(key)){
							excelMapping = mapping;
							break;
						}
				}

				if(excelMapping == null){
					excelMapping = field.getAnnotation(ExcelMapping.class);
					if(excelMapping == null || !excelMapping.key().equals(key)){
						continue;
					}
				}

				coloumnNo = excelMapping.coloumnNo();

				Cell cell = excelSheet.getCell(rowNum, coloumnNo);

				if(fieldType.equals(Integer.class) || fieldType.equals(Double.class) || fieldType.equals(Float.class)) field.set(reflectionObject, cell.getNumericCellValue());
                else if(fieldType.equals(Long.class)) field.set(reflectionObject, (long) cell.getNumericCellValue());
				else if(fieldType.equals(Date.class) || fieldType.equals(Calendar.class)) field.set(reflectionObject, cell.getDateCellValue());
				else if(fieldType.equals(Boolean.class)) field.set(reflectionObject, cell.getStringCellValue().equals("true") || cell.getStringCellValue().equals("t")
						|| cell.getStringCellValue().equals("T"));
				else{
					if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC) field.set(reflectionObject, String.valueOf(cell.getNumericCellValue()));
					else field.set(reflectionObject, cell.getStringCellValue());
				}

			}
		}

		return reflectionObject;
	}

	private static boolean isASupportedKnownClass(Class type) {
		return type.equals(Integer.class) || type.equals(Double.class) || type.equals(Float.class) || type.equals(Date.class) || type.equals(Calendar.class)
				|| type.equals(Boolean.class) || type.equals(String.class) || type.equals(Long.class);
	}

	public static <T> ExcelReadResponse<T> getEntityListFromExcelFile(File file, int sheetNo, boolean useHeader, Class T) throws IOException, InvalidFormatException {

		FileInputStream fileInputStream = new FileInputStream(file);

		ExcelReadResponse<T> excelReadResponse;
		try{
			excelReadResponse = getEntityListFromExcelFile(new FileInputStream(file), sheetNo, useHeader, T);
		}
		finally{
			fileInputStream.close();
		}

		return excelReadResponse;
	}

	public static ExcelReadResponse<Object> getContentMatrixBySheetAsReadResponse(InputStream inputStream) {
		return getContentMatrixBySheetAsReadResponse(inputStream, true, Object.class);
	}

	public static ExcelReadResponse<Object> getContentMatrixBySheetAsReadResponse(File file) throws FileNotFoundException {

		FileInputStream fileInputStream = new FileInputStream(file);

		ExcelReadResponse<Object> excelReadResponse;
		try{
			excelReadResponse = getContentMatrixBySheetAsReadResponse(fileInputStream, true, Object.class);
		}
		finally {
			try{
				fileInputStream.close();
			}
			catch (IOException e){
				//logger.error(e.getMessage());
			}
		}

		return excelReadResponse;
	}

	private static <T> ExcelReadResponse<T> getContentMatrixBySheetAsReadResponse(InputStream inputStream, boolean useHeader, Class T) {
		ExcelReadResponse<T> excelReadResponse = new ExcelReadResponse<T>();
		List<ExcelSheet> excelSheetList = null;
		try{
			excelSheetList = getContentMatriceFromExcelFile(inputStream);
		}
		catch (IOException e){

			ExcelReadError excelReadError = new ExcelReadError();
			excelReadError.setExcelReadErrorMessage(ExcelReadErrorType.IO_EXCEPTION);
			excelReadResponse.getExcelReadErrorList().add(excelReadError);

			excelReadResponse.setSuccessful(false);

		}
		catch (InvalidFormatException e){

			ExcelReadError excelReadError = new ExcelReadError();
			excelReadError.setExcelReadErrorMessage(ExcelReadErrorType.INVALID_FORMAT_EXCEPTION);
			excelReadResponse.getExcelReadErrorList().add(excelReadError);

			excelReadResponse.setSuccessful(false);
		}

		for(int i = 0; i < excelSheetList.size(); i++){
			if(excelSheetList.get(i).getCellMatrice() == null || excelSheetList.get(i).getCellMatrice().size() == 0){
				excelSheetList.remove(excelSheetList.get(i));
				excelReadResponse.getEmptySheets().add(i + excelReadResponse.getEmptySheets().size());
				i--;
			}
		}

		int modeWeight = 0;
		int mode = 0;

		for(ExcelSheet excelSheet : excelSheetList){
			if(useHeader){
				mode = excelSheet.getTotalCellsInARow(0);
			}
			else{
				Map<Integer, Integer> coloumnWeightMap = new HashMap<Integer, Integer>();
				for(int i = 0; i < excelSheet.getTotalRows(); i++)
					addOrIncrement(coloumnWeightMap, excelSheet.getTotalCellsInARow(i));

				for(Integer key : coloumnWeightMap.keySet()){
					if(coloumnWeightMap.get(key) > modeWeight){
						modeWeight = coloumnWeightMap.get(key);
						mode = key;
					}
				}
			}

			int rowNum = 0;
			for(List<Cell> cellMatrice : excelSheet.getCellMatrice()){
				if(cellMatrice.size() < mode){
					ExcelReadError excelReadError = new ExcelReadError(ExcelReadErrorType.NOT_ENOUGH_INPUTS, rowNum);
					excelReadResponse.getExcelReadErrorList().add(excelReadError);
					excelReadResponse.setSuccessful(false);
				}
				else if(cellMatrice.size() > mode){
					ExcelReadError excelReadError = new ExcelReadError(ExcelReadErrorType.TOO_MANY_INPUTS, rowNum);
					excelReadResponse.getExcelReadErrorList().add(excelReadError);
					excelReadResponse.setSuccessful(false);
				}
				rowNum++;
			}
		}

		excelReadResponse.setExcelSheetList(excelSheetList);
		return excelReadResponse;
	}

	private static void addOrIncrement(Map<Integer, Integer> map, Integer key) {
		if(map.containsKey(key)) map.put(key, map.get(key) + 1);
		else map.put(key, 1);
	}

	private static ExcelSheet getContentMatrixBySheet(Sheet sheet) {

		Iterator<Row> rowIterator = sheet.iterator();
		ExcelSheet excelSheet = new ExcelSheet();

		int rowNum = 0;
		while(rowIterator.hasNext()){
			Row row = rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			while(cellIterator.hasNext()){
				Cell cell = cellIterator.next();
				addByOrder(excelSheet, cell, rowNum);
			}

			rowNum++;
		}

		return excelSheet;
	}

	private static void addByOrder(ExcelSheet excelSheet, Cell cell, int rowNum) {
		while(excelSheet.getCellMatrice().size() <= rowNum){
			excelSheet.getCellMatrice().add(new ArrayList<Cell>());
		}

		excelSheet.getCellMatrice().get(rowNum).add(cell);
	}
}
