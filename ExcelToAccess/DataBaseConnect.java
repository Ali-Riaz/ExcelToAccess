/*
 * Class Name: DataBaseConnect.java
 * By: Ali Riaz
 * Class Description: The goal of this file is to establish a connection to the MS Database, and start putting the ArrayList data (from the class TestRead.java) into it
 */

import java.sql.Connection;
import java.sql.DatabaseMetaData;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;

public class DataBaseConnect {
	
	static final String JDBC_DRIVER = "sun.jdbc.odbc.JdbcOdbcDriver";																// Name of the JDBC:ODBC Bridge Driver provided by sun that makes a connection from Java to any Open DataBase Connectivity possible
	static final String DB_URL = "jdbc:odbc:Driver={Microsoft Access Driver (*.mdb)};DBQ=MyDSN.mdb";								// Address that points to your DataBase
	String TABLE_NAME = null;																										// The name of the Table that's going to be created within the Database
	String inputFile = "C:/ExcelToAccess_Stuff/HC2.2 Structural and Geographic by OwnerRenter.xls";									// The pathname where my Excel sheet is located; you'll have a different one
	ArrayList <String> strCells;																									// ArrayList <String> strCells will contain all the data read by the class TestRead.java
	String [] strColumnHeadings = {"A","B","C","D","E","F","G","H","I","J","K","L","M","N"};										// The user has been given the flexibility to write the column names to whatever he/she wishes. Just remember, NO Spaces, NO Apostrophe's within your words..
																																	// Otherwise, you will get an error.
	
	/*
	 * This method sets the name of the Table (to be created within the Database), and is used by the main class ReadAndInsert.java
	 * It cannot have any spaces, or dots(.), try to give it a simple name, or use ( _ ) to separate multiple words
	 */
	
	public void setTableName(String TABLE_NAME){
		this.TABLE_NAME = TABLE_NAME;
	}
	
	/*
	 * This method creates the String "PARAMETER_STRING" used in the SQL Statement "CREATE TABLE_NAME(PARAMETER_STRING)"
	 * Basically, it takes each element of the above declared String Array 'strColumnHeadings', and appends it with " VARCHAR(255)," meaning
	 * each column declared within the Table can have data of type char (up to 255 characters)
	 * 
	 * For instance, a column by the name "A" would become "A VARCHAR(255)," etc...The String "PARAMETER_STRING" would keep on being appended till we've reached
	 * the final parameter of the strColumnHeadings. Once that's done, the final String will be returned.
	 * 
	 * e.g. (A VARCHAR(255),B VARCHAR(255),C VARCHAR(255),D VARCHAR(255),E VARCHAR(255),F VARCHAR(255),G VARCHAR(255),H VARCHAR(255),I VARCHAR(255),J VARCHAR(255),K VARCHAR(255),L VARCHAR(255),M VARCHAR(255),N VARCHAR(255))
	 */
	
	public String createHeadings(String [] strColumnHeadings){
		StringBuilder strBuilder = new StringBuilder();
		int length = strColumnHeadings.length;
		for(int i=0; i< length;){
			strBuilder.append(strColumnHeadings[i]);
			strBuilder.append(" VARCHAR(255)");
			if(i < (length-1)){
				strBuilder.append(',');
			}
			i++;
		}
		return strBuilder.toString();
	}
	
	/*
	 * This method creates the String "COLUMNS_PARAMETER" used in the SQL Statement "INSERT INTO TABLE_NAME(COLUMNS_PARAMETER) VALUES(RESPECTIVE_DATA)"
	 * Basically, it takes each element of the above declared String Array 'strColumnHeadings', and appends it with a comma (,)
	 * For instance, a column by the name "A" would become "A," etc...The String would keep on being appended till we've reached
	 * the final parameter of the strColumnHeadings. Once that's done, the final String will be returned. e.g (A,B,C,D,E,F,G,H,I,J,K,L,M,N)
	 */
		
	public String columnNames(String [] strColumns){
		StringBuilder strBuilder = new StringBuilder();
		int length = strColumns.length;
		for(int i=0; i<length;){
			strBuilder.append(strColumns[i]);
			if(i < (length-1)){
				strBuilder.append(',');
			}
			i++;
		}
		return strBuilder.toString();
	}
	
	/*
	 * This method creates the String "RESPECTIVE_DATA" used in the SQL Statement "INSERT INTO TABLE_NAME(COLUMNS_PARAMETER) VALUES(RESPECTIVE_DATA)"
	 * Basically, it takes each element of the above declared ArrayList <String> 'strColumns', and appends it with single apostrophee's, and a comma ('',)
	 * For instance, a value by the name "Data1" would become 'Data1,' etc...The String would keep on being appended till we've reached the end of the column
	 * Once that's done, the final String will be returned. e.g ('Data1','Data2','Data3','Data4','Data5','Data6','Data7','Data8','Data9','Data10','Data11','Data12','Data13','Data14')
	 */
	
	public String buildString(ArrayList <String> strColumns, int start, int stop){
		StringBuilder strBuilder = new StringBuilder();
		for(int i = (start-1); i < (stop+1-1);){
			strBuilder.append("'" + strColumns.get(i) + "'");
			if(i < (stop+1-1-1)){
				strBuilder.append(',');
			}
			i++;
		}
		return strBuilder.toString();
	}
	
	/*
	 * This method does two things:
	 * 1) Establish a connection to the MS Access Database
	 * 2) Execute SQL Statements like DROP, CREATE, INSERT
	 */
	
	public void getConnection(){
		Connection conn = null;																// The connection to the Database
		Statement stmt = null;																// The statement is used to perform actions on the Database
		TestRead tr = new TestRead();														// Creating an instance of the TestRead class
		tr.setInputFile(inputFile);															// Using that instance to set the pathname for the MS Excel file.
		
		try{
			
			strCells = tr.read();															// Stores the Excel file data into the ArrayList <String> strCells
			
			System.out.println("Initializing JDBC Driver...");
			Class.forName(JDBC_DRIVER);
			
			System.out.println("Establishing connection to DataBase...");
			conn = DriverManager.getConnection(DB_URL);
			
			System.out.println("Creating Statement...");
			stmt = conn.createStatement();
			
			String sql = "DROP TABLE " + TABLE_NAME;
			
			DatabaseMetaData dbm = conn.getMetaData();										// Used to get more information on the Database
		    ResultSet rs = dbm.getTables(null, null, TABLE_NAME, null);
		    
		    /*
		     * This portion of the code checks if any Table with the name TABLE_NAME exists or not
		     * If it does, then the program DROPS the current Table, and then creates a new one
		     * If it doesn't, then it simply creates the Table.
		     */
		    
		    if (rs.next()) {
		      System.out.println("Table exists");
		      stmt.execute(sql);
		    } 
		    else {
		      System.out.println("Table does not exist");
		    }
		    
		    String separatedColumns = createHeadings(strColumnHeadings);
		    System.out.println("The parameters for creating the table are: " + separatedColumns);
		    
		    String columns = columnNames(strColumnHeadings);
		    System.out.println("The COLUMN_PARAMETERS for INSERT method are: " + columns);
		    
		    sql = "CREATE TABLE " + TABLE_NAME + "(" + separatedColumns + ")";
		    stmt.execute(sql);
		    
		    int start = 1;																	// Beginning of the row
		    int stop = 14;																	// Ending of the row
		    
		    while(stop < strCells.size()){													// Is this the final row? If yes, stop. If not, carry on.
			    
		    	String parameters_str = buildString(strCells, start, stop);
		    	sql = "INSERT INTO " + TABLE_NAME + " (" + columns + ")" + " VALUES (" + parameters_str + ")";
			    System.out.println(sql);
			    start += 14;																// This value depends on the number of columns you have in the excel sheet
			    stop += 14;																	// That is, if the excel sheet has 14 columns, then each row is going to have 14 columns
			    																			// So, now we set the start and end values for the new row
			    stmt.executeUpdate(sql);													// Execute the SQL statement above																				
		    }		    

			stmt.close();																	// Closing the statement
			conn.close();																	// Closing the connection
		}
		
		catch(SQLException se){																// Handle errors for JDBC
		      se.printStackTrace();
		   }
		
		catch(Exception e){																	// Handle errors for Class.forName
		      e.printStackTrace();
		   }
			
		finally{																			// Finally block used to close resources if they're still open
		      try{
		         if(stmt!=null)
		            stmt.close();
		      }
		      catch(SQLException se2){														// Nothing we can do
		    	  
		      }
		      
		      try{
		         if(conn!=null)
		            conn.close();
		      }
		      
		      catch(SQLException se){
		         se.printStackTrace();
		      
		      }
		   }																				// End of finally
	}																						// End of getConnection() method
}																							// End of the class DataBaseConnect.java