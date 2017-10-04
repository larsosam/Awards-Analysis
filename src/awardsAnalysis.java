import  java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Scanner;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.monitorjbl.xlsx.StreamingReader;

public class awardsAnalysis {

	//The column the groupings are assigned to
	static int groupingCol;
	
	//The column the element codes are in
	static int elementCol;
	
	//The column the reference codes begin in
	static int referenceCol = 200;
	
	//Hashmaps with values of element codes and their groupings
	static Map<String, ArrayList<String>> ECCodes = null;
	static Map<String, ArrayList<String>> PRCodes = null;
	
	public static void main(String[] args) {
		//String[] arg = { "/Users/slarson22/Desktop/Site Based Database.xlsx", "/Users/slarson22/Desktop/NSF Codebook 20170711.xls", "/Users/slarson22/Desktop/"};
	
		if (args.length != 3) {
			System.out.println("Usage: java -jar AwardsAnalysis.jar "
					+ "[Site Based Database Path] [Codebook Path] [Output Path]");
			System.exit(1);
		}
		
		String siteBasedPath = args[0];
		String codebookPath = args[1];
		String outputPath = args[2];
		
		if (siteBasedPath.endsWith(File.separator)) {
			siteBasedPath = siteBasedPath.substring(0, siteBasedPath.length());
		}
		if (codebookPath.endsWith(File.separator)) {
			codebookPath = codebookPath.substring(0, codebookPath.length());
		}
		if (!outputPath.endsWith(File.separator)) {
			outputPath += File.separator;
		}
		
		//Checks site based database
		if (!siteBasedPath.endsWith("xlsx")) {
			System.out.println("Incorrect file extension for site based database. Please use .xlsx");
			System.exit(1);
		}
		
		if (!new File(siteBasedPath).exists()) {
			System.out.println("Site based database cannot be found or opened");
			System.exit(1);
		}
		
		//Checks codebook
		if (!(codebookPath.endsWith("xlsx") || codebookPath.endsWith(".xls"))) {
			System.out.println("Incorrect file extension for codebook. Please use .xlsx or .xls");
			System.exit(1);
		}
		
		if (!new File(codebookPath).exists()) {
			System.out.println("Codebook cannot be found or opened");
			System.exit(1);
		}
		
		//Checks outputPath
		if (!new File(outputPath).exists()) {
			System.out.println("Output path cannot be found or opened");
		}
		
		if (!new File(outputPath).isDirectory()) {
			System.out.println("Output Path is not directory");
		}
		
		if (new File(outputPath + "Grouped Site Based Database.xlsx").exists()) {
			System.out.println("Output file already exists. Overwrite? y/n");
			boolean hasChoice = false;
			while (hasChoice == false) {
				Scanner scnr = new Scanner(System.in);
				String choice = scnr.next();
				if (choice.equals("y")) {
					hasChoice = true;
					scnr.close();
				} else if (choice.equals("n")) {
					System.out.println("Exiting Program");
					System.exit(0);
				} else {
					System.out.println("Invalid Input. y for yes, n for no");
				}
			}
		}
		
		//Opening the Site Based Database
		System.out.println("Opening Site Based Database");
		
		Workbook data = null;
		InputStream is = null;
		
		//Creates the site based workbook stream
		try {
			is = new FileInputStream(new File(siteBasedPath));
			data = StreamingReader.builder()
			        .rowCacheSize(100)    // number of rows to keep in memory (defaults to 10)
			        .bufferSize(4096)     // buffer size to use when reading InputStream to file (defaults to 1024)
			        .open(is);            // InputStream or File for XLSX file (required)
	
		} catch (Exception e) {
			System.out.println("Database either could not be opened, or could not be found");
			System.exit(1);
		}

		
		//Creates the output stream
		SXSSFWorkbook output = new SXSSFWorkbook(10);
		SXSSFSheet outputSheet = output.createSheet();
		SXSSFRow outputHeader = outputSheet.createRow(0);
		
		Sheet dataSheet = data.getSheetAt(0);
		Iterator<Row> itr = dataSheet.rowIterator();
		Row header = itr.next();
		for (Cell cell : header) {
			if (cell.getColumnIndex() == 0) {
				outputHeader.createCell(0).setCellValue(cell.getStringCellValue());
			} else {
				outputHeader.createCell(outputHeader.getLastCellNum()).setCellValue(cell.getStringCellValue());
			}
			if (cell.getStringCellValue().equals("Program Element Code")) {
				elementCol = cell.getColumnIndex();
			}
			if (cell.getStringCellValue().equals("Program Reference Code") &&
					cell.getColumnIndex() < referenceCol) {
				referenceCol = cell.getColumnIndex();
			}
		}
		groupingCol = header.getLastCellNum();
		outputHeader.createCell(groupingCol).setCellValue("Grouping");
		

		//Creating the HashMaps
		
		//Get keywords/element codes
		System.out.println("Opening Codebook");
		Workbook codebook = openWorkbook(codebookPath);
		
		//The Important Sheets
		Sheet ECSheet = null;
		Sheet PRCSheet = null;
		try {
			ECSheet = codebook.getSheet("Element Codes");
			PRCSheet = codebook.getSheet("Program Reference Codes");
		} catch (IllegalArgumentException ex) {
			System.out.println("Please format the codebook so one sheet is named 'Element Codes'"
					+ " and another 'Program Reference Codes'");
			System.exit(1);
		}
		
		ECCodes = new HashMap<String, ArrayList<String>>();
		PRCodes = new HashMap<String, ArrayList<String>>();
		
		//Making arrays of the hits we want in the HashMap
		for (Integer i = 1; i < 26; i++) {
			ECCodes.put(i.toString(), new ArrayList<String>());
			PRCodes.put(i.toString(), new ArrayList<String>());
		}
		
		//Assigning set values for the column the codes and groupings are in
		int ECGroupingCol = 1000;
		int ECCodeCol = 1000;
		for (Cell cell : ECSheet.getRow(0)) {
			if (cell.toString().equals("Code")) {
				ECCodeCol = cell.getColumnIndex();
			}
			if (cell.toString().equals("Grouping")) {
				ECGroupingCol = cell.getColumnIndex();
			}
		}
		
		//Incorrectly formatted element codebook
		if (ECGroupingCol == 1000) {
			System.out.println("Please title the grouping column for the element codes as 'Grouping'");
			System.exit(1);
		}
		if (ECCodeCol == 1000) {
			System.out.println("Please title the code column for the element codes as 'Code'");
			System.exit(1);
		}
		
		int PRCGroupingCol = 1000;
		int PRCCodeCol = 1000;
		for (Cell cell : PRCSheet.getRow(0)) {
			if (cell.toString().equals("Prog_Code")) {
				PRCCodeCol = cell.getColumnIndex();
			}
			if (cell.toString().equals("Grouping")) {
				PRCGroupingCol = cell.getColumnIndex();
			}
		}
		
		//Incorrectly formatted codebook
		if (PRCGroupingCol == 1000) {
			System.out.println("Please title the grouping column for the reference codes as 'Grouping'");
			System.exit(1);
		}
		if (PRCCodeCol == 1000) {
			System.out.println("Please title the code column for the reference codes as 'Prog_Code'");
			System.exit(1);
		}
		
		//For every row in the codebook, take the coding and put it in the key
		for (Row row : ECSheet) {
			try {
				if (ECCodes.containsKey(row.getCell(ECGroupingCol).toString().
						substring(0, row.getCell(ECGroupingCol).toString().length()-2))) {
					
					//Add the code to the hashmap
					ECCodes.get(row.getCell(ECGroupingCol).toString().
							substring(0, row.getCell(ECGroupingCol).toString().length()-2)).
						add(row.getCell(ECCodeCol).toString().replace(".0", ""));
				}
			} catch (NumberFormatException ex) {
				System.out.println(row.getRowNum());
			} catch (NullPointerException ex) {
				
			} catch (StringIndexOutOfBoundsException ex) {

			}
		}
		for (Row row : PRCSheet) {
			try {
				if (PRCodes.containsKey(row.getCell(PRCGroupingCol).toString())) {
					
					//Add the code to the hashmap
					PRCodes.get(row.getCell(PRCGroupingCol).toString()).
						add(row.getCell(PRCCodeCol).toString().replace(".0", ""));
				}
			} catch (NumberFormatException ex) {
				System.out.println(row.getRowNum());
			} catch (NullPointerException ex) {
				
			}
		}

		//Groupings!		
		System.out.println("Assigning Groupings");
		for (Row row : dataSheet) {
			assignGrouping(row, outputSheet.createRow(row.getRowNum()));
		}
		
		try {
			
			Sheet codebookSheet = data.getSheet("Codebook");
			Sheet outputCodebook = output.createSheet("Codebook");
			for (Row row : codebookSheet) {
				Row outputRow = outputCodebook.createRow(outputCodebook.getLastRowNum() + 1);
				for (Cell cell : row) {
					if (outputRow.getLastCellNum() == -1) {
						outputRow.createCell(0).setCellValue(cell.getStringCellValue());
					}
					else {
						outputRow.createCell(outputRow.getLastCellNum()).setCellValue(cell.getStringCellValue());
					}
				}
			}
		} catch (NullPointerException ex) {
			System.out.println("No codebook");
		}

		//Writes the output
		System.out.println("Writing Output");
		try {
			FileOutputStream fileOut = new FileOutputStream(outputPath + "Grouped Site Based Database.xlsx");
			output.write(fileOut);
			fileOut.close();
			System.out.println("Output successfully written to " + outputPath + "Grouped Site Based Database.xlsx");
			output.close();
		} catch (IOException e) {
			System.out.println("Somehow the file you want to write to no longer exists. Cmon.");
		}
		if (true) {
			return;
		}
		
	}
	
	/**
	 * Opens the workbook given as an argument, returns the workbook
	 * @param filename The workbook to open
	 * @return
	 */
	public static Workbook openWorkbook(String filename) {
		Workbook workbook = null;
		
		try {
			workbook = WorkbookFactory.create(new FileInputStream(filename));			
		} catch (IOException ex){
			System.out.println("The file cannot be found");
		} catch (EncryptedDocumentException e) {
			System.out.println("The file is encrypted");
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			System.out.println("The file is of invalid format");
		}
		
		return workbook;
	}
	
	public static void assignGrouping(Row dataRow, SXSSFRow outputRow) {
		for (int i = 0; i < dataRow.getLastCellNum(); i++) {
			outputRow.createCell(i).setCellValue(dataRow.getCell(i).getStringCellValue());
		}
		
		for (Integer i = 1; i < 26; i++) {
			
			//ECCodes!
			if (ECCodes.get(i.toString()).contains(dataRow.getCell(elementCol).getStringCellValue())) {
				
				//Creates the cell if it doesn't exist yet
				if (outputRow.getCell(groupingCol) == null) {
					outputRow.createCell(groupingCol).setCellValue(i);
				}
				
				//Adds the grouping to the end if others exist
				else {							
					boolean same = false;
					for (String str : outputRow.getCell(groupingCol).toString().split(", ")) {
						if (str.replace(".0", "").equals(i.toString())) {
							same = true;
						}
					}
					if (!same) {
						outputRow.getCell(groupingCol).setCellValue(
								outputRow.getCell(groupingCol).toString()
								.replace(".0", "") + ", " + i.toString());
						
						//Set the multiple value indicator
						outputRow.createCell(groupingCol + 1).
						setCellValue("Multiple References");
					}
				}
			}
			
			//Reference Codes
			for (Integer j = referenceCol; j < groupingCol; j += 2) {
				
				//Checks to make sure there is actually a code there
				if (dataRow.getCell(j) != null) {
					
					//If the code is in the specified grouping
					if (PRCodes.get(i.toString()).contains(dataRow.getCell(j).getStringCellValue())) {
						//Creates the cell if it doesn't exist yet
						if (outputRow.getCell(groupingCol) == null) {
							outputRow.createCell(groupingCol).setCellValue(i);
						}
						else {							
							boolean same = false;
							for (String str : outputRow.getCell(groupingCol).toString().split(", ")) {
								if (str.replace(".0", "").equals(i.toString())) {
									same = true;
								}
							}
							if (!same) {
								outputRow.getCell(groupingCol).setCellValue(
										outputRow.getCell(groupingCol).toString()
										.replace(".0", "") + ", " + i.toString());
								
								//Set the multiple value indicator
								outputRow.createCell(groupingCol + 1).
								setCellValue("Multiple References");
							}
						}
					}
				}
			}
		}
	}
}
