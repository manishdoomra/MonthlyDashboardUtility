/**
 * 
 */
package com.nexenta.utilities.monthly.dashboard.generator;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author manish.doomra
 *
 */
public class MonthlyDashboardUtilityImpl implements IMonthlyDashboardUtility{
	
	/***
	 * 
	 * Reads Section And Categories Excel File - This Excel file contains section names distributed amongst various sheets where each sheet is the category
	 * 
	 * @param sectionCategoriesExcelFile
	 * @return
	 * @throws IOException
	 */
	private Map<String, String> readSectionAndCategoriesFromFile(String sectionCategoriesExcelFile) throws IOException{
		FileInputStream flatFile = null;
		try {
			Map<String, String> sectionAndCategories = new HashMap<String, String>();
			flatFile = new FileInputStream(sectionCategoriesExcelFile);
			Workbook flatFileWorkbk = WorkbookFactory.create(flatFile);
			if(flatFileWorkbk!=null){
				int numberOfSheets = flatFileWorkbk.getNumberOfSheets();
				for(int i=0; i<numberOfSheets;i++){
					Sheet categorySheet = flatFileWorkbk.getSheetAt(i);
					String sheetName = categorySheet.getSheetName();
					int numberOfRows = categorySheet.getPhysicalNumberOfRows();
					
					for(int j=0; j<numberOfRows;j++){
						Row row = categorySheet.getRow(j);
						String cellValue = row.getCell(0).toString();
						if(!cellValue.trim().equalsIgnoreCase("Section Name")){								
							sectionAndCategories.put(cellValue, sheetName);
						}
					}
				}
				return sectionAndCategories;
			}else{
				System.out.println("ERROR : The section And Category Excel File is null");
				return null;
			}
		} catch (InvalidFormatException | IOException e) {			
			e.printStackTrace();
		} finally{
			if(flatFile!=null){
				flatFile.close();
			}
		}
		return null;
	}
	
	/***
	 * 
	 * Read Testrail excel file raw data exported
	 * 
	 * Pre-requisites:- 
	 * a) There would be only 1 sheet or the data is populated in the first sheet
	 * b) Where the excel file is populated with 5 columns namely :-
	 * 		ID	Title	Automation?	Created On	Section
	 * c) Automation column is populated in third column and Section is populated in 5th column
	 * 
	 * It will return the data structure as {CLI={Manual Only=5, Cucumber Automated=6, Needs Review=3, Needs Automation=1}}  
	 * 
	 * @param testrailExcelFile
	 * @return
	 * @throws IOException 
	 */
	private Map<String, Map<String, Integer>> readTestRailRawData(String testrailExcelFile) throws IOException {
		FileInputStream flatFile = null;
		try {
			Map<String, Map<String, Integer>> testRailData = new HashMap<String, Map<String, Integer>>();
			flatFile = new FileInputStream(testrailExcelFile);
			Workbook flatFileWorkbk = WorkbookFactory.create(flatFile);
			
			if (flatFileWorkbk != null) {
				
				Sheet firstSheet = flatFileWorkbk.getSheetAt(0);				
				int numberOfRows = firstSheet.getPhysicalNumberOfRows();				
				for(int i=0; i<numberOfRows;i++){
					if(i!=0){ // Ignore first row, since it represents header
						Row row = firstSheet.getRow(i);
						Cell automationCell = row.getCell(2);
						String automationFlag= "Blank-Automation-Flag";
						if(automationCell!=null){
							automationFlag  = automationCell.toString();
						}
						String sectionName = row.getCell(4).toString();
						sectionName = sectionName.trim();
						if(testRailData.containsKey(sectionName)){
							Map<String, Integer> automationFlagMapForASection = testRailData.get(sectionName);
							
							if(automationFlagMapForASection.containsKey(automationFlag.trim())){
								Integer oldVal = automationFlagMapForASection.get(automationFlag.trim());
								automationFlagMapForASection.put(automationFlag.trim(), oldVal + 1);
							}
							testRailData.put(sectionName, automationFlagMapForASection);
							
						}else{ // putting in map for the first time
							Map<String, Integer> automationFlagMapForASection = populateDefaultValuesForInternalMap();
							if(automationFlagMapForASection.containsKey(automationFlag.trim())){								
								automationFlagMapForASection.put(automationFlag.trim(), 1);
							}
							testRailData.put(sectionName, automationFlagMapForASection);
							
						}
					}
				}
				return testRailData;
			}
		} catch (InvalidFormatException | IOException e) {
			e.printStackTrace();
		} finally {
			if (flatFile != null) {
				flatFile.close();
			}
		}
		return null;
	}
	
	private Map<String, Integer> populateDefaultValuesForInternalMap(){
		Map<String, Integer> automationFlagMapForASection = new HashMap<String, Integer>();
		automationFlagMapForASection.put("Manual Only", new Integer(0));
		automationFlagMapForASection.put("Cucumber Automated", new Integer(0));
		automationFlagMapForASection.put("Needs Review", new Integer(0));
		automationFlagMapForASection.put("Needs Automation", new Integer(0));
		automationFlagMapForASection.put("GUI Automated", new Integer(0));
		automationFlagMapForASection.put("Automated", new Integer(0));	
		automationFlagMapForASection.put("Blank-Automation-Flag", new Integer(0));	
		return automationFlagMapForASection;
	}
	
	
	private Map<String, Map<String, Integer>> categorizeSections(Map<String, String> sectionAndCategories, Map<String, Map<String, Integer>> testRailRawData){
		Map<String, Map<String, Integer>> condensedForm = new HashMap<String, Map<String,Integer>>();
		for(Map.Entry<String, Map<String, Integer>> entry : testRailRawData.entrySet()){
			if(sectionAndCategories.containsKey(entry.getKey())){ // if section name is present in excel file
				String section = sectionAndCategories.get(entry.getKey());
				if(!condensedForm.containsKey(section)){ // if section is being populated for the first time
					condensedForm.put(section, testRailRawData.get(entry.getKey()));
				}else{ // Already present in the map
					Map<String,Integer> automationFlagMapForASection = condensedForm.get(section);
					Map<String,Integer> automationFlagMapForASectionNew  = entry.getValue();
					for(Map.Entry<String, Integer> oldEntry : automationFlagMapForASection.entrySet()){
						Integer oldVal = oldEntry.getValue();
						String key = oldEntry.getKey();
						automationFlagMapForASectionNew.put(key, oldVal + automationFlagMapForASectionNew.get(key)); 
					}
					condensedForm.put(section, automationFlagMapForASectionNew);
				}
			}else{
				String section = entry.getKey();
				if(!condensedForm.containsKey(section)){ 
					condensedForm.put(section, testRailRawData.get(section));
				}else{ // Already present in the map
					Map<String,Integer> automationFlagMapForASection = condensedForm.get(section);
					Map<String,Integer> automationFlagMapForASectionNew  = entry.getValue();
					for(Map.Entry<String, Integer> oldEntry : automationFlagMapForASection.entrySet()){
						Integer oldVal = oldEntry.getValue();
						String key = oldEntry.getKey();
						automationFlagMapForASectionNew.put(key, oldVal + automationFlagMapForASectionNew.get(key)); 
					}
					condensedForm.put(section, automationFlagMapForASectionNew);
				}
			}
		}
		return condensedForm;
	}
	
	/*public static void main(String[] args){
		MonthlyDashboardUtilityImpl util = new MonthlyDashboardUtilityImpl();
		try {
			Map<String, String> sectionAndCategories = util.readSectionAndCategoriesFromFile("C:/Manish/Nexenta/source_31_Aug_new/F-TAF_Tests_5.0/utils-5.0/src/test/resources/SectionCategories.xls");
			Map<String, Map<String, Integer>> testRailRawData = util.readTestRailRawData("C:/Manish/Nexenta/source_31_Aug_new/F-TAF_Tests_5.0/utils-5.0/src/test/resources/TestRailRawData_30May.xls");
			Map<String, Map<String, Integer>>  map = util.categorizeSections(sectionAndCategories, testRailRawData);
			System.out.println();
			XSSFWorkbook workbook = new XSSFWorkbook(); 
			XSSFSheet spreadsheet = workbook.createSheet("Sheet Name");			
			XSSFRow row;
            Cell cell;            
            row = spreadsheet.createRow(0);
            Cell cell1 = row.createCell(0);
            cell1.setCellValue("Section Name");
            Set<String> keys = util.populateDefaultValuesForInternalMap().keySet();
            int col = 1;
            for(String key :  keys){
            	row.createCell(col++).setCellValue(key);
            }
            
            
            Set<Entry<String, Map<String, Integer>>> set = map.entrySet();
            int rowNum = 1;
            for(Entry<String, Map<String, Integer>> entry : set){
            	String sectionName = entry.getKey();
            	Row sectionRow = spreadsheet.createRow(rowNum++);
            	sectionRow.createCell(0).setCellValue(sectionName);
            	Set<Entry<String, Integer>> rows = entry.getValue().entrySet();
            	int colNum =1;
            	for(Entry<String,Integer> rowPlaceholder : rows){
            		Integer colVal = rowPlaceholder.getValue();
            		System.out.println(colVal+" --- "+ rowPlaceholder.getKey());
            		sectionRow.createCell(colNum++).setCellValue(colVal); 
            	}
            	
            }
            
			FileOutputStream out = new FileOutputStream(new File("C:/Manish/Nexenta/source_31_Aug_new/F-TAF_Tests_5.0/utils-5.0/src/test/resources/createworkbook_May.xlsx"));
			workbook.write(out);
		    out.close();
		    System.out.println("createworkbook.xlsx written successfully");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}*/

	@Override
	public void generateMonthlyDashboardTestRailData(String rawDataTestrailExcelFile,String testRailSectionAndItsCategoriesExcelFile, String generatedExcelFile) {
		System.out.println(rawDataTestrailExcelFile);
		System.out.println(testRailSectionAndItsCategoriesExcelFile);
		System.out.println(generatedExcelFile);
		XSSFWorkbook generatedWorkbook = null;
		try{
			Map<String, String> sectionAndCategories = readSectionAndCategoriesFromFile(testRailSectionAndItsCategoriesExcelFile);
			Map<String, Map<String, Integer>> testRailRawData = readTestRailRawData(rawDataTestrailExcelFile);
			Map<String, Map<String, Integer>>  map = categorizeSections(sectionAndCategories, testRailRawData);
			generatedWorkbook = new XSSFWorkbook(); 
			XSSFSheet spreadsheet = generatedWorkbook.createSheet("Sheet Name");			
			XSSFRow row;
            row = spreadsheet.createRow(0);
            Cell cell1 = row.createCell(0);
            cell1.setCellValue("Section Name");
            Set<String> keys = populateDefaultValuesForInternalMap().keySet();
            int col = 1;
            for(String key :  keys){
            	row.createCell(col++).setCellValue(key);
            }
            
            
            Set<Entry<String, Map<String, Integer>>> set = map.entrySet();
            int rowNum = 1;
            for(Entry<String, Map<String, Integer>> entry : set){
            	String sectionName = entry.getKey();
            	Row sectionRow = spreadsheet.createRow(rowNum++);
            	sectionRow.createCell(0).setCellValue(sectionName);
            	Set<Entry<String, Integer>> rows = entry.getValue().entrySet();
            	int colNum =1;
            	for(Entry<String,Integer> rowPlaceholder : rows){
            		Integer colVal = rowPlaceholder.getValue();
            		sectionRow.createCell(colNum++).setCellValue(colVal); 
            	}
            	
            }
            
			FileOutputStream out = new FileOutputStream(new File(generatedExcelFile));
			generatedWorkbook.write(out);
		    out.close();
		    System.out.println(generatedExcelFile+" written successfully");
		}catch(IOException e){
			System.err.println("Problem in reading/writing excel file");
		}finally{
			try {
				generatedWorkbook.close();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}

}
