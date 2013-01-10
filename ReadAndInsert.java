/*
 * Class Name: ReadAndInsert.java
 * By: Ali Riaz
 * Class Description: This class acts as the starter of the entire program. It does only 4 things:
 * 					1) Sets the name of the Table that you're going to create.
 * 					2) Creates a new instance of the class DataBaseConnect called 'dbc' within the main method
 * 					3) Uses that instance to call two methods, the first one being setTableName(). For more details on this method, look at DataBaseConnect.java file.
 * 					4) Finally, it calls the second method getConnection(), where it establishes a connection with the MS Access Database, and starts putting data into it.
 * 	  				For more details on this method, look at the DataBaseConnect.java file
 */


import java.io.IOException;


public class ReadAndInsert {

	static final String TABLE_NAME = "Structural_and_Geographic_by_OwnerRenter";
	
	public static void main(String[] args) throws IOException {
		
		DataBaseConnect dbc = new DataBaseConnect();
		dbc.setTableName(TABLE_NAME);
		dbc.getConnection();
	}

}
