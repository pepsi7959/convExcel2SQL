package sqlconvertor.pepsi7959.github.com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App 
{
	public static  Map<Integer, String> header = null;
	
	public static Iterator<Cell> skipCell(Row row, int skip_col) {
		Iterator<Cell> cellIterator = row.cellIterator();
		for(int i = 0; i < skip_col; i++) {
			if(cellIterator.hasNext())
				cellIterator.next();
		}
		return cellIterator;
	}
	
    private static String getHeaderValue(int i) {

		return App.header.get(i);
	}

	private static LinkedList<String> cloneLinkedlist(LinkedList<String> listFixedCell) {
    		LinkedList<String> newLinkedlist = new LinkedList<String>();
    		for (String object : listFixedCell) {
				newLinkedlist.add(object);
			}
		return newLinkedlist;
	}

	public static Iterator<Row> skipRow(Sheet sheet, int skip_row) {
		Iterator<Row> rowIterator = sheet.iterator();
		for(int i = 0; i < skip_row; i++) {
			if( rowIterator.hasNext() )
				rowIterator.next();
			else
				System.out.println("Warning: cannot skip row ");
		}
		return rowIterator;
	}
	
	public static String toString(Cell cell) {
		String value = "";
		
		switch( cell.getCellType() ) {
			case Cell.CELL_TYPE_NUMERIC :
				value = (int)cell.getNumericCellValue()+"";
				break;
			case Cell.CELL_TYPE_STRING:
				value = cell.getStringCellValue();
			default:
				//System.out.println("Enum: " + cell.getCellType() );
				value = cell.toString();
		}
		
		return value;
	}
	
	/* version 0.0.1 */
	public static Map<Integer, String> getHeader(String fileLocation) {
		Map<Integer, String> header = new HashMap<>();
		FileInputStream file;
		try {
			file = new FileInputStream(new File(fileLocation ));
			Workbook workbook = new XSSFWorkbook(file);
			Sheet sheet = workbook.getSheetAt(0);
			int i = 0;
			
			Iterator<Row> rowIterator = skipRow(sheet, 1);
			
			while(rowIterator.hasNext()) {
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator  = skipCell(row, 0);
				
				int idxCell = 0;
			    while(cellIterator.hasNext()) {
			    		Cell cell = cellIterator.next();
			        switch ( cell.getCellTypeEnum() ) {
			            case STRING:
			            		System.out.print(cell.getStringCellValue() + " | ");
			            		header.put(idxCell, cell.getStringCellValue());
			            		break;
			            case NUMERIC:
			            		System.out.print(cell.getNumericCellValue() + " | ");
			            		header.put(idxCell, cell.getNumericCellValue()+"");
			            		break;
			            default:
			            		System.out.print("Nan | ");
			        }
			        
			        idxCell++;
			    }
			    System.out.println();
			    if( true ) break;
			    i++;
			}
		
		}catch (FileNotFoundException e) {
			e.printStackTrace();
		}catch (IOException e) {
			e.printStackTrace();
		}
		return header;
	}
	
	public static Map<Integer, LinkedList<String>> readHeaderConf(String fileLocation, int sheetId,  int rowBegin, int cellBegin, int cellRange) {
		
		FileInputStream file;
		Workbook workbook;
		Sheet sheet;
		Map<Integer, LinkedList<String>> header = new HashMap<>();

		try {
			
			file = new FileInputStream(new File(fileLocation ));
			workbook = new XSSFWorkbook(file);
			sheet = workbook.getSheetAt(sheetId);
			
			int idxRow = 0;
			Iterator<Row> rowIterator = skipRow(sheet, rowBegin);
			
			while( rowIterator.hasNext() ) {
					
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator  = skipCell(row, cellBegin);
				
				int idxCell = 0;
				LinkedList<String> headerData = new LinkedList<>();

			    while( cellIterator.hasNext()  && (idxCell < cellRange) ) {
			    		Cell cell = cellIterator.next();
			    		String value = cell.toString();
	            		System.out.print(value + " | ");
	            		headerData.add(idxCell, value);
			        idxCell++;
			    }
			    
			    System.out.println();
			    if( headerData.size() > 0)
			    header.put(idxRow, headerData);
			    idxRow++;
			}

			workbook.close();
			
		}catch (FileNotFoundException e) {
			e.printStackTrace();
		}catch (IOException e) {
			e.printStackTrace();
		}finally {
		}
		
		return header;
	}

	public static ArrayList<LinkedList<String>> convertRowHeaderToCell(String fileLocation, int fixedColumn) {
		ArrayList<LinkedList<String>> data = new ArrayList<>();
		FileInputStream file;
		
		try {
			
			file = new FileInputStream(new File(fileLocation ));
			Workbook workbook = new XSSFWorkbook(file);
			Sheet sheet = workbook.getSheetAt(0);
			Iterator<Row> rowIterator = skipRow(sheet, 2);
			
			while(rowIterator.hasNext()) {
				
				LinkedList<String> listFixedCell = new LinkedList<>();
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator  = skipCell(row, 0);
				int idxCell = 0;

			    while(cellIterator.hasNext()) {

			    		Cell cell = cellIterator.next();
			    		
			    		if( idxCell < fixedColumn) {
				        switch ( cell.getCellTypeEnum() ) {
				            case STRING:
				            		System.out.print(cell.getStringCellValue() + " | ");
				            		listFixedCell.add( cell.getStringCellValue() );
				            		break;
				            case NUMERIC:
				            		System.out.print(cell.getNumericCellValue() + " | ");
				            		listFixedCell.add( cell.getNumericCellValue()+"" );
				            		break;
				            default:
				            		System.out.print("Nan | ");
				        }
			    		}else{
			    				String value = "";
			    				String headerValue = "";
			    				LinkedList<String> newRecord = null;
			    				
					        switch ( cell.getCellTypeEnum() ) {
				            case STRING:
				            		System.out.print(cell.getStringCellValue() + " | ");
				            		value = cell.getStringCellValue();
				            		newRecord = cloneLinkedlist(listFixedCell);
				            		if( value.equals("1") ) {
				            			headerValue = getHeaderValue(idxCell-fixedColumn);
				            			newRecord.add(headerValue);
				            			data.add(newRecord);
				            		}else {
				            			//do nothing
				            		}
				            		break;
				            case NUMERIC:
				            		System.out.print(cell.getNumericCellValue() + " | ");
				            		value = cell.getNumericCellValue()+"";
				            		newRecord = cloneLinkedlist(listFixedCell);
				            		
				            		if( value.equals("1.0") ) {
				            			headerValue = getHeaderValue(idxCell-fixedColumn);
				            			newRecord.add(headerValue);
				            			data.add(newRecord);
				            			System.out.println("data size: " + data.size());
				            		}else {
				            			//do nothing
				            		}
				            		break;
				            default:
				            		System.out.print("Nan | ");
				        }//end switch
					        
			    		}
			        idxCell++;
			    }
			    System.out.println();
			}
			workbook.close();
		}catch (FileNotFoundException e) {
			e.printStackTrace();
		}catch (IOException e) {
			e.printStackTrace();
		}
		System.out.println("data size : " + data.size());
		return data;
	}
	
	public static ArrayList<LinkedList<String>> convertRowToCell(Map<Integer, LinkedList<String>> headerConfig, String fileLocation, int sheetId,  int fixedCell, int rowBegin, int rowRange, int cellBegin, int cellRange) {
		
		FileInputStream file;
		Workbook workbook;
		Sheet sheet;
		ArrayList<LinkedList<String>> data = new ArrayList<>();

		try {
			
			file = new FileInputStream(new File(fileLocation ));
			workbook = new XSSFWorkbook(file);
			sheet = workbook.getSheetAt(sheetId);
			
			System.out.println("Number of Rows: " + cellRange);
			
			int idxRow = 0;
			double percentage = 0;
			Iterator<Row> rowIterator = skipRow(sheet, rowBegin);
			int rowToProcess = (rowRange == -1)?sheet.getLastRowNum()-rowBegin:rowRange;
					
			while( rowIterator.hasNext() && (idxRow < rowRange) ) {
					
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator  = skipCell(row, cellBegin);
				
				int idxCell = 0;
				LinkedList<String> reservedCell = new LinkedList<>();
				LinkedList<String> record = new LinkedList<>();
				
			    while( cellIterator.hasNext()  && (idxCell < cellRange) ) {
			    	
		    			Cell cell = cellIterator.next();
			    		String value = cell.toString();
			    		
			    		if( idxCell <= fixedCell ) {
				    		reservedCell.add(idxCell, value);
			    		}else {
			    			if( value.equals("1.0") ) {
				    			record = cloneLinkedlist(reservedCell);
				    			record.add(headerConfig.get(idxCell - fixedCell).get(0));
				    			record.add(headerConfig.get(idxCell - fixedCell).get(1));
				    			data.add(record);
			    			}
			    		}
			    		
			        idxCell++;
			    }
			    
			    
			   percentage = ((double)idxRow/(double)rowToProcess) * 100.00;
			   //if( (int)percentage % 25 == 0)
			    		System.out.println("job status: " + percentage +"%");
		    		idxRow++;
			}

			workbook.close();
			
		}catch (FileNotFoundException e) {
			e.printStackTrace();
		}catch (IOException e) {
			e.printStackTrace();
		}finally {
		}
		
		return data;
	}

	
	/* version 0.0.2 */

	public static Map<String, LinkedList<String>> readDoumentMaster(String fileLocation, int sheetId,  int rowBegin, int cellBegin, int cellRange) {
		
		System.out.println("\n****************************** start reading document master  ******************************" );
		
		FileInputStream file;
		Workbook workbook;
		Sheet sheet;
		Map<String, LinkedList<String>> header = new HashMap<>();

		try {
			
			file = new FileInputStream(new File(fileLocation ));
			workbook = new XSSFWorkbook(file);
			sheet = workbook.getSheetAt(sheetId);
			
			int idxRow = 0;
			Iterator<Row> rowIterator = skipRow(sheet, rowBegin);
			
			while( rowIterator.hasNext() ) {
					
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator  = skipCell(row, cellBegin);
				
				int idxCell = 0;
				LinkedList<String> headerData = new LinkedList<>();
				String key = "";
				
				/* read key */
				Cell cell = cellIterator.next();
				if( cellIterator.hasNext() ) {
					key =  toString(cell);
				}
				
				if( key.isEmpty() ) {
					System.out.println("Warning: skipping empty key at (" + cell.getAddress() + ") ");
					break;
				}else {
					System.out.print("Row("+ idxRow +") > ");
					System.out.print(key + " | ");
					headerData.add(key);
					idxCell++;
					
					/* read value */
				    while( cellIterator.hasNext()  && (idxCell < cellRange) ) {
				    		cell = cellIterator.next();
				    		String value = cell.toString();
		            		System.out.print(value + " | ");
		            		headerData.add(value);
				        idxCell++;
				    }
			    
				    System.out.println();
				    if( headerData.size() > 0)
				    header.put(key, headerData);
				}
			    idxRow++;
			}

			workbook.close();
			
		}catch (FileNotFoundException e) {
			e.printStackTrace();
		}catch (IOException e) {
			e.printStackTrace();
		}finally {
		}
		
		return header;
	}
	
	public static Map<String, String> readHeaderDocument(String fileLocation, int sheetId, int rowBegin, int rowRange, int skipCell, int cellRange) {
		
		System.out.println("\n****************************** start reading header document  ******************************" );
		
		FileInputStream file;
		Workbook workbook;
		Sheet sheet;
		Map<String, String> header = new HashMap<>();

		try {
			
			file = new FileInputStream(new File(fileLocation ));
			workbook = new XSSFWorkbook(file);
			sheet = workbook.getSheetAt(sheetId);
			
			int idxRow = 0;
			Iterator<Row> rowIterator = skipRow(sheet, rowBegin);
			
			while( rowIterator.hasNext() ) {
					
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator  = skipCell(row, skipCell);
				
				int idxCell = 0;
				LinkedList<String> headerData = new LinkedList<>();

				
			    while( cellIterator.hasNext()  && (idxCell < cellRange) ) {
			    		Cell cell = cellIterator.next();
			    		String value = toString(cell);
	            		System.out.print(value + " | ");
	            		header.put(cell.getColumnIndex()+"", value);
			        idxCell++;
			    }
			    
			    System.out.println();
			    idxRow++;
			    break;
			}

			workbook.close();
			
		}catch (FileNotFoundException e) {
			e.printStackTrace();
		}catch (IOException e) {
			e.printStackTrace();
		}finally {
		}
		
		return header;
	}

	public static ArrayList<LinkedList<String>> convertProcedure(Map<String, LinkedList<String>> documentmaster, Map<String, String> header, String fileLocation, int sheetId,  int fixedCell, int rowBegin, int rowRange, int cellBegin, int cellRange) {
		
		System.out.println("\n****************************** start converting proceure  ******************************" );
		
		FileInputStream file;
		Workbook workbook;
		Sheet sheet;
		ArrayList<LinkedList<String>> data = new ArrayList<>();
		int fixedCellSize = fixedCell + 1;
		int countUseDocuments = 0;

		try {
			
			file = new FileInputStream(new File(fileLocation ));
			workbook = new XSSFWorkbook(file);
			sheet = workbook.getSheetAt(sheetId);
			
			System.out.println("Number of Rows: " + cellRange);
			
			int idxRow = 0;
			double percentage = 0;
			Iterator<Row> rowIterator = skipRow(sheet, rowBegin);
			int rowToProcess = (rowRange == -1)?sheet.getLastRowNum()-rowBegin:rowRange;
					
			while( rowIterator.hasNext() && (idxRow < rowToProcess) ) {
					
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator  = skipCell(row, cellBegin);
				
				int idxCell = 0;
				LinkedList<String> reservedCell = new LinkedList<>();
				LinkedList<String> record = new LinkedList<>();
				System.out.println("Row ("+ idxRow +") : ");
				
			    while( cellIterator.hasNext()  && (idxCell < cellRange) ) {
			    	
		    			Cell cell = cellIterator.next();
		    			int cellAddr = cell.getColumnIndex();
			    		String value = toString(cell);
			    		System.out.println("    - "+"("+cell.getAddress()+")("+cellAddr+")" + value +" | ");
			    		
			    		if( cellAddr <= fixedCell ) {
			    			while( reservedCell.size() < cellAddr ) {
			    				reservedCell.add("");
			    			}
				    		reservedCell.add(value);
			    		}else {
			    			while( reservedCell.size() < fixedCellSize ) {
			    				reservedCell.add("");
			    			}
			    			if( value.equals("1") ){
		    					countUseDocuments++;
			    				String docId = header.get(cellAddr+"");
			    				System.out.println("        - cell("+ cellAddr +") DocId(" + docId + ") | ");
			    				
			    				if( documentmaster.get(docId) != null) 	//header 
			    				{
			    					String docName = documentmaster.get(docId).get(1);
			    					String docOwner = documentmaster.get(docId).get(2);
			    					
			    					record = cloneLinkedlist(reservedCell);
			    					record.add(docId);
			    					record.add(docName);
			    					record.add(docOwner);
			    					data.add(record);
			    				}else {
			    					System.out.println("Error: Not found document master");
			    					System.exit(-1);
			    				}
				    			
			    			}else {
			    				if(!value.equals("")) {
			    					System.out.println("Warning: unexpected value ("+ value +")");
			    				}
			    			}
			    		}
			        idxCell++;
			    }
			    
			    
			    percentage = (int)(((double)idxRow/(double)rowToProcess) * 100.00);
			    	System.out.println("job status: " + percentage +"%");
		    		idxRow++;
			}

			workbook.close();
			
		}catch (FileNotFoundException e) {
			e.printStackTrace();
		}catch (IOException e) {
			e.printStackTrace();
		}finally {
		}
		
		System.out.println("\n****************************** Number of documents : "+ countUseDocuments + "  ******************************");
		
		return data;
	}
	
	public static ArrayList<LinkedList<String>> convertProcedureForBenz(Map<String, LinkedList<String>> documentmaster, Map<String, String> header, String fileLocation, int sheetId,  int fixedCell, int rowBegin, int rowRange, int cellBegin, int cellRange) {
		
		System.out.println("\n****************************** start converting proceure  ******************************" );
		
		FileInputStream file;
		Workbook workbook;
		Sheet sheet;
		ArrayList<LinkedList<String>> data = new ArrayList<>();
		int fixedCellSize = fixedCell + 1;

		try {
			
			file = new FileInputStream(new File(fileLocation ));
			workbook = new XSSFWorkbook(file);
			sheet = workbook.getSheetAt(sheetId);
			
			System.out.println("Number of Rows: " + cellRange);
			
			int idxRow = 0;
			double percentage = 0;
			Iterator<Row> rowIterator = skipRow(sheet, rowBegin);
			int rowToProcess = (rowRange == -1)?sheet.getLastRowNum()-rowBegin:rowRange;
					
			while( rowIterator.hasNext() && (idxRow < rowToProcess) ) {
					
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator  = skipCell(row, cellBegin);
				
				int idxCell = 0;
				LinkedList<String> reservedCell = new LinkedList<>();
				System.out.println("Row ("+ idxRow +") : ");
				
			    while( cellIterator.hasNext()  && (idxCell < cellRange) ) {
			    	
		    			Cell cell = cellIterator.next();
		    			int cellAddr = cell.getColumnIndex();
			    		String value = toString(cell);
			    		System.out.println("    - "+"("+cell.getAddress()+")("+cellAddr+")" + value +" | ");
			    		
			    		if( cellAddr <= fixedCell ) {
			    			while( reservedCell.size() < cellAddr ) {
			    				reservedCell.add("");
			    			}
				    		reservedCell.add(value);
			    		}else {
			    			
			    			while( reservedCell.size() < fixedCellSize ) {
			    				reservedCell.add("");
			    			}
			    			
			    			/* lists of compound documentName::ownerDocument */
			    			String values[] = value.split(";;|;");
			    			System.out.println("    \t- all(" + value + ")");
			    			
			    			/* documentName::ownerDocument */
			    			for(int i = 0; i < values.length; i++) {
			    				
			    				System.out.println("    \t- " + values[i]);
			    				
			    				LinkedList<String> newRecord = cloneLinkedlist(reservedCell);
			    				String nameAndOwner[] = values[i].trim().split("::"); 
			    				
			    				/* Document Name */
			    				if( nameAndOwner.length > 1 || !nameAndOwner[0].isEmpty() ) {
			    					newRecord.add(nameAndOwner[0].trim());
			    				}else {
			    					System.out.println("Error: Not found \"Document Name\"");
			    				}
			    				
			    				/* Owner name */
			    				if( nameAndOwner.length == 2 && !nameAndOwner[1].isEmpty() ) {
			    					newRecord.add(nameAndOwner[1].trim());
			    				}else {
			    					System.out.println("Warning: Not found \"Owner Name\"");
			    				}
			    				
			    				data.add(newRecord);

			    			}
			    			
			    		}
			        idxCell++;
			        
			    }
			    
			    
			    percentage = (int)(((double)idxRow/(double)rowToProcess) * 100.00);
			    	System.out.println("job status: " + percentage +"%");
		    		idxRow++;
			}

			workbook.close();
			
		}catch (FileNotFoundException e) {
			e.printStackTrace();
		}catch (IOException e) {
			e.printStackTrace();
		}finally {
		}
		
		return data;
	}
	

	/* utility */
	
	public static int cellAddressToInt(String cellStr) {

		int numOfAphabets = 26;
		int len = cellStr.length();
		cellStr = cellStr.toUpperCase();
		
		if( len == 1 ) {
			/* A - Z */
			return cellStr.charAt(0) - 'A'  ;
		}else if( len == 2){
			/* AA - AA */
			/*index0:index1*/
			int index_0 = cellStr.charAt(0) - 'A' + 1;
			int index_1 = cellStr.charAt(1) - 'A';
			return (numOfAphabets*index_0) + index_1;
		}else {
			System.out.println("Error: not support cell size more than 2 characters");
		}
		return -1;
	}
	
	public static void exportToCSV(String destFilePath, ArrayList<LinkedList<String>> listofRows, String[] header) {
		
		System.out.println( "\n****************************** exporting to Excel ***************************** ");
		String sheetName = "output";
		FileOutputStream outstream;
		Workbook workbook;
		Sheet sheet;
		
		
		try {
			
			outstream = new FileOutputStream(new File(destFilePath ));
			workbook = new XSSFWorkbook();
			
			sheet = workbook.createSheet(sheetName);
			int idxRow = 0;
			
			/* Write header */
			if( header != null ) {
				
				Row row = sheet.createRow(idxRow++);
				int idxCell = 0;
				
				for( int i = 0; i < header.length; i++) {
					Cell cell = row.createCell(idxCell++);
					//System.out.println("> "+value);
					cell.setCellValue(header[i].trim());
				}
			}
				
			for(LinkedList<String> listofCols : listofRows) {
				Row row = sheet.createRow(idxRow++);
				Iterator<String> iterator = listofCols.iterator();
				int idxCell = 0;
				
				while( iterator.hasNext() ) {
					Cell cell = row.createCell(idxCell++);
					String value = iterator.next();
					//System.out.println("> "+value);
					cell.setCellValue(value.trim());
				}
				
			}
			
			workbook.write(outstream);
			outstream.close();
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		System.out.println( "\n----------------------------- exporting to "+ destFilePath +" -----------------------------");
	}

	public static void exportToSQL(String destFilePath, ArrayList<LinkedList<String>> listofRows, String metadata, int selectedIndex[]){
		
		System.out.println( "\n****************************** exporting to SQL ***************************** ");
		
		FileWriter fileWriter;
		File file;
		
		try {
			
			int rowCount = 0;
			file = new File(destFilePath);
			fileWriter = new FileWriter(file);
			
			for(LinkedList<String> listofCols : listofRows) {
				
				Iterator<String> iterator = listofCols.iterator();
				int idxCell = 0;

				if( rowCount%1000 == 0 ) {
					if ( rowCount > 0 )
						fileWriter.write(";\n");
					fileWriter.write(metadata);
				}
				
				fileWriter.write('(');
				StringBuilder temp = new StringBuilder();
				
				while( iterator.hasNext() ) {
					String value = iterator.next();
					if( selectedIndex[idxCell] == 1 ) {
						temp.append("'" + value.trim() + "',");
					}
					idxCell++;
				}
				
				if( temp.length() > 0 ) {
					temp.deleteCharAt(temp.length()-1);
				}
				
				fileWriter.write(temp.toString());
				fileWriter.write(")");
				
				rowCount++;
				
				if( rowCount%1000 != 0) {
					fileWriter.write(",\n");
				}

				
			}
			
			fileWriter.close();
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		System.out.println( "\n----------------------------- exporting to "+ destFilePath +" -----------------------------");
	}
	
	private static void display(ArrayList<LinkedList<String>> listofRows) {
		System.out.println("\n****************************** display *****************************" );
		for(LinkedList lists : listofRows) {
			Iterator it = lists.iterator();
			while( it.hasNext() ) {
				System.out.print(it.next().toString() + '|');
			}
			System.out.println();
		}
	}
	
	private static void displayHeader(Map<Integer, LinkedList<String>> headerConf) {
		for (int i = 0; i < headerConf.size() ; i++) {
			LinkedList<String> cellLists = headerConf.get(i);
			if( cellLists != null && cellLists.get(0) != null && cellLists.get(1) != null) {
				System.out.println("row(" + i + ")" + cellLists.get(0) +" "+ cellLists.get(1));
			}
		}
	}
	
	private static void displayDocumentMaster(Map<String, LinkedList<String>> documents) {
		for(LinkedList<String> document: documents.values()) {
			if( document != null && document.get(0) != null && document.get(1) != null) {
				System.out.println(document.get(0) +" "+ document.get(1));
			}
		}
	}
	
	
	/* Convert */
	public static void convertForPPoom() {
		
		int sheetId = 0;
		int beginRow = 1;
		int beginCell = 0;
		int rowRange = 0;
		int cellRange = 3;
		int fixedCell = 0;
    		Map<String, LinkedList<String>>  doumentMaster = readDoumentMaster("/Users/narongsak.mala/Documents/GDX/DocumentMaster-V.1.0.0.xlsx", sheetId, beginRow, beginCell, cellRange);
    		displayDocumentMaster(doumentMaster);
    		
		sheetId = 0;
		beginRow = 3;
		beginCell = 9;
		cellRange = 444;
		Map<String, String>  headerDocumentMap = readHeaderDocument("/Users/narongsak.mala/Documents/GDX/procedureToConvert/CleansingFromPPoom.xlsx", sheetId, beginRow, rowRange, beginCell, cellRange);
		int i = 0;
		for(String header: headerDocumentMap.values()) {
			System.out.println("Header (" + i++ +") : " + header);
		}
		
    		sheetId = 0;
    		beginRow = 4;
    		rowRange = 408;//317;
    		fixedCell = cellAddressToInt("I");
    		beginCell = 0;
    		cellRange = cellAddressToInt("JQ")+1;
    		ArrayList<LinkedList<String>> listOfProcedures =  convertProcedure(doumentMaster, headerDocumentMap, "/Users/narongsak.mala/Documents/GDX/procedureToConvert/CleansingFromPPoom.xlsx", sheetId, fixedCell, beginRow, rowRange, beginCell, cellRange);
    		display(listOfProcedures);
    		
    		String header[] = {"ประเภทหน่วยงาน", "ชื่อกระทรวง", "ชื่อหน่วยงาน", "ProcedureID", "Procedure Grouping", "ชื่อบริการ/ชื่อกระบวนงาน", "ซื่อเล่น", "ความถี่", "แบบคำขอ", "เลขเอกสาร", "ชื่อเอกสาร", "เจ้าของเอกสาร"};
    		exportToCSV("/Users/narongsak.mala/Documents/GDX/out/PPoom.xlsx", listOfProcedures, header);
    		
    		String metadata = "INSERT INTO [dbo].[GovTech_Publish]([MinistryName], [DepartmentName],[ProcedureID], [ProcedureName], [CitizenGuideDocumentName], [OwnerDocumentOrgName])\n" + "VALUES";
    		int selectedIndex[] = {0, 1, 1, 1, 0, 1, 0, 0, 0, 0, 1, 1};
    		exportToSQL("/Users/narongsak.mala/Documents/GDX/out/PPoom.sql", listOfProcedures, metadata, selectedIndex);
    		System.out.println("Info: Task is complete.");
    		
	}
		
	public static void convertForPPoom2() {
		
		int sheetId = 3;
		int beginRow = 1;
		int beginCell = 0;
		int rowRange = 0;
		int cellRange = 3;
		int fixedCell = 0;
    		Map<String, LinkedList<String>>  doumentMaster = readDoumentMaster("/Users/narongsak.mala/Documents/GDX/DocumentMaster-V.1.3.0.xlsx", sheetId, beginRow, beginCell, cellRange);
    		displayDocumentMaster(doumentMaster);
    		
    		System.exit(0);
    		
		sheetId = 1;
		beginRow = 3;
		beginCell = 7;
		cellRange = 113;
		Map<String, String>  headerDocumentMap = readHeaderDocument("/Users/narongsak.mala/Documents/GDX/procedureToConvert/CleansingFromPPoom2.xlsx", sheetId, beginRow, rowRange, beginCell, cellRange);
		int i = 0;
		for(String header: headerDocumentMap.values()) {
			System.out.println("Header (" + i++ +") : " + header);
		}
		

		
    		sheetId = 1;
    		beginRow = 4;
    		rowRange = 31;//317;
    		fixedCell = cellAddressToInt("G");
    		beginCell = 0;
    		cellRange = cellAddressToInt("DP")+1;
    		ArrayList<LinkedList<String>> listOfProcedures =  convertProcedure(doumentMaster, headerDocumentMap, "/Users/narongsak.mala/Documents/GDX/procedureToConvert/CleansingFromPPoom2.xlsx", sheetId, fixedCell, beginRow, rowRange, beginCell, cellRange);
    		display(listOfProcedures);
    		
    		String header[] = {"ประเภทหน่วยงาน", "ชื่อกระทรวง", "ชื่อหน่วยงาน", "ProcedureID", "Procedure Grouping", "ชื่อบริการ/ชื่อกระบวนงาน", "แบบคำขอ", "เลขเอกสาร", "ชื่อเอกสาร", "เจ้าของเอกสาร"};
    		exportToCSV("/Users/narongsak.mala/Documents/GDX/out/PPoom2.xlsx", listOfProcedures, header);
    		String metadata = "INSERT INTO [dbo].[GovTech_Publish]([MinistryName], [DepartmentName],[ProcedureID], [ProcedureName], [CitizenGuideDocumentName], [OwnerDocumentOrgName])\n" + "VALUES";
    		int selectedIndex[] = {0, 1, 1, 1, 0, 1, 0, 0, 1, 1};
    		exportToSQL("/Users/narongsak.mala/Documents/GDX/out/PPoom2.sql", listOfProcedures, metadata, selectedIndex);
    		System.out.println("Info: Task is complete.");
    		
	}
	
	public static void convertForPBee() {
		
		int sheetId = 0;
		int beginRow = 1;
		int beginCell = 0;
		int rowRange = 0;
		int cellRange = 3;
		int fixedCell = 0;
    		Map<String, LinkedList<String>>  doumentMaster = readDoumentMaster("/Users/narongsak.mala/Documents/GDX/DocumentMaster-V.1.0.0.xlsx", sheetId, beginRow, beginCell, cellRange);
    		displayDocumentMaster(doumentMaster);
    		
		sheetId = 0;
		beginRow = 2;
		beginCell = 9;
		cellRange = 376;
		Map<String, String>  headerDocumentMap = readHeaderDocument("/Users/narongsak.mala/Documents/GDX/procedureToConvert/CleansingFromPBee.xlsx", sheetId, beginRow, rowRange, beginCell, cellRange);
		int i = 0;
		for(String header: headerDocumentMap.values()) {
			System.out.println("Header (" + i++ +") : " + header);
		}
		
    		sheetId = 0;
    		beginRow = 3;
    		rowRange = 711;//715;
    		fixedCell = cellAddressToInt("I");
    		beginCell = 0;
    		cellRange = 385;
    		ArrayList<LinkedList<String>> listOfProcedures =  convertProcedure(doumentMaster, headerDocumentMap, "/Users/narongsak.mala/Documents/GDX/procedureToConvert/CleansingFromPBee.xlsx", sheetId, fixedCell, beginRow, rowRange, beginCell, cellRange);
    		display(listOfProcedures);
    		
    		String header[] = {"ประเภทหน่วยงาน", "ชื่อกระทรวง", "ชื่อหน่วยงาน", "Procedure ID", "Procedure Grouping", "ชื่อบริการ/ชื่อกระบวนงาน","ชื่อเล่น", "ความถี่", "แบบคำขอ", "ชื่อเอกสาร", "เจ้าของเอกสาร"};
    		exportToCSV("/Users/narongsak.mala/Documents/GDX/out/PBee.xlsx", listOfProcedures, header);
    		
    		String metadata = "INSERT INTO [dbo].[GovTech_Publish]([MinistryName], [DepartmentName],[ProcedureID], [ProcedureName], [CitizenGuideDocumentName], [OwnerDocumentOrgName])\n" + "VALUES";
    		int selectedIndex[] = {0, 1, 1, 1, 0, 1, 0, 0, 0, 0, 1, 1};
    		exportToSQL("/Users/narongsak.mala/Documents/GDX/out/PBee.sql", listOfProcedures, metadata, selectedIndex);
    		
    		System.out.println("Info: Task is complete.");
    		
	}
	
	public static void convertForPBoon() {
		
		int sheetId = 1;
		int beginRow = 1;
		int beginCell = 0;
		int rowRange = 0;
		int cellRange = 3;
		int fixedCell = 0;
    		Map<String, LinkedList<String>>  doumentMaster = readDoumentMaster("/Users/narongsak.mala/Documents/GDX/DocumentMaster-V.1.1.1.xlsx", sheetId, beginRow, beginCell, cellRange);
    		displayDocumentMaster(doumentMaster);
    		
		sheetId = 2;
		beginRow = 2;
		beginCell = 9;
		cellRange = 444;
		Map<String, String>  headerDocumentMap = readHeaderDocument("/Users/narongsak.mala/Documents/GDX/procedureToConvert/CleansingFromPBoon.xlsx", sheetId, beginRow, rowRange, beginCell, cellRange);
		int i = 0;
		for(String header: headerDocumentMap.values()) {
			System.out.println("Header (" + i++ +") : " + header);
		}
		
    		sheetId = 3;
    		beginRow = 3;
    		rowRange = 314;//317;
    		fixedCell = cellAddressToInt("I");
    		beginCell = 0;
    		cellRange = 444;
    		ArrayList<LinkedList<String>> listOfProcedures =  convertProcedure(doumentMaster, headerDocumentMap, "/Users/narongsak.mala/Documents/GDX/procedureToConvert/CleansingFromPBoon.xlsx", sheetId, fixedCell, beginRow, rowRange, beginCell, cellRange);
    		display(listOfProcedures);
    		
    		String header[] = {"ประเภทหน่วยงาน", "กระทรวง", "หน่วยงาน", "กลุ่มกระบวนงาน", "เลขกระบวนงาน", "กระบวนงาน", "ชื่อเล่น", "ความถี่", "แบบคำขอ", "เลขเอกสาร", "ชื่อเอกสาร", "เจ้าของเอกสาร"};
    		exportToCSV("/Users/narongsak.mala/Documents/GDX/out/PBoon.xlsx", listOfProcedures, header);
    		System.out.println("Info: Task is complete.");
    		
    		String metadata = "INSERT INTO [dbo].[GovTech_Publish]([MinistryName], [DepartmentName],[ProcedureID], [ProcedureName], [CitizenGuideDocumentName], [OwnerDocumentOrgName])\n" + "VALUES";
    		int selectedIndex[] = {0, 1, 1, 0, 1, 1, 0, 0, 0, 0, 1, 1};
    		exportToSQL("/Users/narongsak.mala/Documents/GDX/out/PBoon.sql", listOfProcedures, metadata, selectedIndex);
    		System.out.println("Info: Task is complete.");
    		
	}

	public static void convertForPBenz() {
		
		int sheetId = 2;
		int beginRow = 1;
		int beginCell = 0;
		int rowRange = 0;
		int cellRange = 3;
		int fixedCell = 0;
    		Map<String, LinkedList<String>>  doumentMaster = readDoumentMaster("/Users/narongsak.mala/Documents/GDX/DocumentMaster-V.1.2.0.xlsx", sheetId, beginRow, beginCell, cellRange);
    		displayDocumentMaster(doumentMaster);
    		
    		
		sheetId = 7;
		beginRow = 3;
		beginCell = 9;
		cellRange = 440;
		Map<String, String>  headerDocumentMap = readHeaderDocument("/Users/narongsak.mala/Documents/GDX/procedureToConvert/CleansingFromPBenz.xlsx", sheetId, beginRow, rowRange, beginCell, cellRange);
		int i = 0;
		for(String header: headerDocumentMap.values()) {
			System.out.println("Header (" + i++ +") : " + header);
		}
		

    		sheetId = 7;
    		beginRow = 4;
    		rowRange = 43;
    		fixedCell = cellAddressToInt("I");
    		beginCell = 0;
    		cellRange = 449;
    		ArrayList<LinkedList<String>> listOfProcedures =  convertProcedure(doumentMaster, headerDocumentMap, "/Users/narongsak.mala/Documents/GDX/procedureToConvert/CleansingFromPBenz.xlsx", sheetId, fixedCell, beginRow, rowRange, beginCell, cellRange);
    		display(listOfProcedures);
    		
    		String header[] = {"ประเภทหน่วยงาน", "กระทรวง", "หน่วยงาน",  "เลขกระบวนงาน", "กลุ่มกระบวนงาน", "กระบวนงาน", "ชื่อเล่น", "ความถี่", "แบบคำขอ", "เลขเอกสาร", "ชื่อเอกสาร", "เจ้าของเอกสาร"};
    		exportToCSV("/Users/narongsak.mala/Documents/GDX/out/PBenz.xlsx", listOfProcedures, header);
    		System.out.println("Info: Task is complete.");
    		
    		String metadata = "INSERT INTO [dbo].[GovTech_Publish]([MinistryName], [DepartmentName],[ProcedureID], [ProcedureName], [CitizenGuideDocumentName], [OwnerDocumentOrgName])\n" + "VALUES";
    		int selectedIndex[] = {0, 1, 1, 1, 0, 1, 0, 0, 0, 0, 1, 1};
    		exportToSQL("/Users/narongsak.mala/Documents/GDX/out/PBenz.sql", listOfProcedures, metadata, selectedIndex);
    		System.out.println("Info: Task is complete.");
    		
	}

	/* main */
	
	public static void main( String[] args )
    {
		//convertForPPoom();
		//convertForPPoom();
		//convertForPBoon();
		//convertForPBee();
		convertForPBenz();
		//convertForPPoom2();
		
    }
	

}
