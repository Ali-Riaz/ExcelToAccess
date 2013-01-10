/*
 * Class Name: TestRead.java
 * By: Ali Riaz
 * Class Description: The goal of this file is to read all the data cells within the excel sheet, and add them within an ArrayList <String> 'strCells'
 * 					  along the way. The way it reads the data is row by row. That is, once it's finished reading an entire row, it moves on to the
 * 					  next one. For flexibility, the user has been given the option to select which row he/she wants to start reading from, and similarly,
 * 					  which row he/she wants to end the reading.
 */

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

import jxl.Cell;
import jxl.CellType;
import jxl.LabelCell;
import jxl.NumberCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class TestRead {

	private String inputFile;													// This variable will contain the pathname for the excel file
	String stringCell = null;													// This variable will contain the contents of each cell
	ArrayList <String> strCells = new ArrayList <String> ();					// The ArrayList that's going to be used to store all the data in the Excel file
	
	public void setInputFile(String inputFile){
		this.inputFile = inputFile;
	}
	
	public void addEmptyStrings(ArrayList <String> str, int count){				// count for how many times you want to add an empty String
																				// In other words, how many column cells do you want to skip because of the heading
		for(int i = 0; i<count; i++){
			str.add("");
		}
	}
	
	public ArrayList <String> read() throws IOException {
		File inputWorkbook = new File(inputFile);
		Workbook w;
		
		try {
			w = Workbook.getWorkbook(inputWorkbook);							// Gets the MS Excel file in which you enter and store related data.
			Sheet sheet = w.getSheet(0);										// Gets the worksheet (also known as spreadsheet) of the Excel file where you store and manipulate the data
			
			int counter = 11;													// Zero index based, meaning, the program will start reading from the 12th row
			int row_end = 302;													// Same situation as above, that is, the program will end reading the excel file after the 301st row.
			
			/*
			 *  Use 'counter' to start from whichever row you want
			 *  And 'row_end' to stop at whichever row you want
			 */
			
			while(counter < row_end){											
			
			/*
			 *  Use 'int i' to start from whichever column you want, keep in mind that this is also zero index based
			 *  For example i=0 will make it read from the beginning column
			 *  And i=5 will make it read from the 6th column	
			 */
			
				for(int i = 0; i< sheet.getColumns(); i++){							
				
					Cell cell = sheet.getCell(i, counter);						// Gets the cell at position (i,counter)..Where 'i' is the column, and 'counter' is the row
					CellType type = cell.getType();								// Gets info on what type of cell it is, for instance, NUMBER? LABEL? or just EMPTY?
					
					/*
					 * Case 1: Maybe the cell retrieved is a LABEL type, meaning it contains a String value. If that's the case, then get the contents of the cell.
					 * And finally, add the String to the ArrayList <String> strCells.
					 * 
					 * NOTE: This case might not apply to your excel sheet:
					 * 1) There was a LABEL with the value 'Don't Know' within the excel sheet, and that apostrophe within the 'Don't' word was giving me problems 
					 * with the syntax in the next file when I was trying to put it into the Database. So, I just wanted to get rid of that apostrophe, and change
					 * the word to 'Dont Know'
					 */
					
					if(type == CellType.LABEL){
						LabelCell lc = (LabelCell) cell;
						stringCell = lc.getContents();							
						
						if(stringCell.equalsIgnoreCase("Don't Know")){
							stringCell = "Dont Know";
						}
											
						strCells.add(stringCell);
					}
					
					/*
					 * Case 2: Maybe the cell retrieved is a NUMBER type, meaning it contains a numeric value. If that's the case, then get the contents of the
					 * cell, type-cast it to a String type, and finally add it to the ArrayList <String> strCells
					 */
					
					else if(type == CellType.NUMBER){
						NumberCell nc = (NumberCell) cell;
						stringCell = (String) nc.getContents();	
						strCells.add(stringCell);
					}
					
					/*
					 * Case 3: NOTE: These are some additional cases that I introduced to read my version of the excel file, you might not necessarily need them
					 * In the excel file that I was trying to read, I would sometimes get an entirely empty row, so in order to keep the 
					 * format of the data intact, I had to skip to the next row by adding 14 (in this case, 14 is the total number of
					 * columns in a row) to 'i'
					 */
					
					else if(type == CellType.EMPTY && i == 0){						// Check if it's the 1st column and if it's empty, then skip to next row
						i += 14;
					}
					
					/*
					 * Case 4: NOTE: These are some additional cases that I introduced to read my version of the excel file, you might not necessarily need them.
					 * In the excel file that I was trying to read, I would sometimes get a row with exactly one column and the rest empty, so in order to keep 
					 * the format of the data intact, I had to skip to the next row by adding 13 (in this case, 13 is the remaining number of columns in the row)
					 * to 'i'.
					 */
					
					else if(type == CellType.EMPTY && i == 1){						// Check if it's the 2nd column and if it's empty, then add 13 empty Strings
																					// using the user defined method 'addEmptyString()'
						addEmptyStrings(strCells, 13);
						i += 13;
					}
				}
				
				counter++;															// Once the program gets finished with all the columns in one row, time to move 
																					// on to the next row.
				
			}																		// End of while method
			
			w.close();																// Closing the workBook
		}																			// End of try method
		
		catch(BiffException e){
			e.printStackTrace();
		}
		
		return strCells;															// At the end, an ArrayList <String> strCells is returned. (Filled with all the data)
	
	}																				// End of method read()
	
}																					// End of class TestRead.java