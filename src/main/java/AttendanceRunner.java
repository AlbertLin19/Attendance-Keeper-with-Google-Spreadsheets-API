

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.*;
import com.google.api.services.sheets.v4.Sheets;

import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Scanner;

//altered the Quickstart template given by Google API tutorials to use with
//my attendance keeping program - Albert Lin
public class AttendanceRunner {
    /** Application name. */
    private static final String APPLICATION_NAME =
        "Attendance Keeper";

    /** Directory to store user credentials for this application. */
    private static final java.io.File DATA_STORE_DIR;

    /** Global instance of the {@link FileDataStoreFactory}. */
    private static FileDataStoreFactory DATA_STORE_FACTORY;

    /** Global instance of the JSON factory. */
    private static final JsonFactory JSON_FACTORY =
        JacksonFactory.getDefaultInstance();

    /** Global instance of the HTTP transport. */
    private static HttpTransport HTTP_TRANSPORT;

    /** Global instance of the scopes required by this quickstart.
     *
     * If modifying these scopes, delete your previously saved credentials
     * at ~/.credentials/sheets.googleapis.com-java-quickstart
     */
    private static final List<String> SCOPES =
        Arrays.asList(SheetsScopes.SPREADSHEETS);

    static {
    	if (System.getProperty("os.name").equals("Linux")) {
    		System.out.println("Hopefully, this is Linux.");
    		DATA_STORE_DIR = new java.io.File(
    		        "/home/pi/Desktop", ".credentials/sheets.googleapis.com-java-quickstart");
        } else {
        	DATA_STORE_DIR = new java.io.File(
        	        "C:\\Users\\Albert Lin\\git\\FRC Team Charging Champions 2017 Pre Season Attendance Keeper", ".credentials/sheets.googleapis.com-java-quickstart");
        }
        try {
            HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
            DATA_STORE_FACTORY = new FileDataStoreFactory(DATA_STORE_DIR);
        } catch (Throwable t) {
            t.printStackTrace();
            System.exit(1);
        }
    }

    /**
     * Creates an authorized Credential object.
     * @return an authorized Credential object.
     * @throws IOException
     */
    public static Credential authorize() throws IOException {
    	
        // Load client secrets.
        InputStream in =
            AttendanceRunner.class.getResourceAsStream("/client_secret.json");
        GoogleClientSecrets clientSecrets =
            GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow =
                new GoogleAuthorizationCodeFlow.Builder(
                        HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                .setDataStoreFactory(DATA_STORE_FACTORY)
                .setAccessType("offline")
                .build();
        Credential credential = new AuthorizationCodeInstalledApp(
            flow, new LocalServerReceiver()).authorize("user");
        System.out.println(
                "Credentials saved to " + DATA_STORE_DIR.getAbsolutePath());
        return credential;
    }

    /**
     * Build and return an authorized Sheets API client service.
     * @return an authorized Sheets API client service
     * @throws IOException
     */
    public static Sheets getSheetsService() throws IOException {
        Credential credential = authorize();
        return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }

    public static void main(String[] args) throws IOException {
    	
    	ArrayList<Member> roster = new ArrayList<>();
    	String currentColumn = "";
    	
        // Build a new authorized API client service.
        Sheets service = getSheetsService();

        // Get Old Date and Old Column
        String spreadsheetId = "1L9D1xp9WsVt0UkRq9ggEhWJged5NRxlaRmQOdIE8clg";
        String oldDateRange = "Attendance Sheet!A2:A2";
        String oldColumnRange = "Attendance Sheet!A5:A5";
        ValueRange oldDateList = service.spreadsheets().values()
            .get(spreadsheetId, oldDateRange)
            .execute();
        List<List<Object>> dateValue = oldDateList.getValues();
        ValueRange oldColumnList = service.spreadsheets().values()
            .get(spreadsheetId, oldColumnRange)
            .execute();
        List<List<Object>> columnValue = oldColumnList.getValues();
        String oldDate = (String) dateValue.get(0).get(0);
        String oldColumn = (String) columnValue.get(0).get(0);
        currentColumn = oldColumn;
        
        //get the current date
        Date date = new Date();
		DateFormat dateFormat = new SimpleDateFormat("MM/dd/yy");
		String newDate = dateFormat.format(date);
		
		
		if (!newDate.equals(oldDate)) {
			
			List<List<Object>> values = Arrays.asList(
				Arrays.asList((Object) newDate)
					// Additional rows ...
					);
					ValueRange requestBody = new ValueRange()
					.setValues(values);
	    	Sheets.Spreadsheets.Values.Update request =
	    	service.spreadsheets().values().update(spreadsheetId, "Attendance Sheet!A2:A2", requestBody);
	    	request.setValueInputOption("USER_ENTERED");

	    	UpdateValuesResponse response = request.execute();

	    	// TODO: Change code below to process the `response` object:
	    	System.out.println(response);
			String newColumn = "";
	    	if (!oldColumn.substring(oldColumn.length()-1, oldColumn.length()).equalsIgnoreCase("Z") || !oldColumn.substring(oldColumn.length()-1, oldColumn.length()).equalsIgnoreCase("Y")) {
	    		int charValue = oldColumn.charAt(oldColumn.length()-1);
	    		if (oldColumn.length()>1) {
	    			newColumn = oldColumn.substring(0, oldColumn.length()-1) + String.valueOf((char) (charValue + 2));
	    			currentColumn = newColumn;
	    		} else {
	    			newColumn = String.valueOf((char) (charValue + 2));
	    			currentColumn = newColumn;
	    		}
	    	} else {
	    		if (!oldColumn.substring(oldColumn.length()-1, oldColumn.length()).equalsIgnoreCase("Y")) {
	    			newColumn = "AA";
	    			currentColumn = newColumn;
	    		} else {
	    			newColumn = "AA";
	    			currentColumn = newColumn;
	    		}
	    		
	    	}
	    	List<List<Object>> columnVal = Arrays.asList(
					Arrays.asList((Object) (String) (newColumn))
						// Additional rows ...
						);
						ValueRange requestBodyForCol = new ValueRange()
						.setValues(columnVal);
		    	Sheets.Spreadsheets.Values.Update requestForCol =
		    	service.spreadsheets().values().update(spreadsheetId, "Attendance Sheet!A5:A5", requestBodyForCol);
		    	requestForCol.setValueInputOption("USER_ENTERED");

		    	UpdateValuesResponse responseCol = requestForCol.execute();
		    	
		    	Sheets.Spreadsheets.Values.Update request2 =
		    	    	service.spreadsheets().values().update(spreadsheetId, "Attendance Sheet!" + newColumn + "1:" + newColumn + "1", requestBody);
		    	    	request2.setValueInputOption("USER_ENTERED");

		    	    	UpdateValuesResponse response2 = request2.execute();

		    	    	// TODO: Change code below to process the `response` object:
		    	    	System.out.println(response2);
		}
		
		//refresh any that needs to be refreshed
		String refreshIndicatorRange = "Attendance Sheet!B2:B";
		ValueRange refreshList = service.spreadsheets().values()
	            .get(spreadsheetId, refreshIndicatorRange)
	            .execute();
	        List<List<Object>> refreshIndicatorValues = refreshList.getValues();
	        if (refreshIndicatorValues == null || refreshIndicatorValues.size() == 0) {
	            System.out.println("No refresh needed.");
	        } else {
	          for (int i = 0; i <refreshIndicatorValues.size(); i++) {
	            //refresh
	          }
	          
	        }

		
		//create an arrayList of members based on spreadsheet
        String nameListRange = "Attendance Sheet!C3:C";
        ValueRange nameList = service.spreadsheets().values()
            .get(spreadsheetId, nameListRange)
            .execute();
        List<List<Object>> nameValues = nameList.getValues();
        if (nameValues == null || nameValues.size() == 0) {
            System.out.println("No data found.");
        } else {
          for (int i = 0; i <nameValues.size(); i++) {
            roster.add(new Member((String) nameValues.get(i).get(0), i+3));
          }
          
        }
        
        printRoster(roster);
        
        //taking roll
        boolean takingRoll = true;
        Scanner input = new Scanner(System.in);
        while (takingRoll) {
        	
        	if (input.hasNext())
        	{
        	String nameInput = input.next().trim();
        	for (Member mem : roster) {
        		if (nameInput.equalsIgnoreCase(mem.getName())) {
        			String timeListRange = "Attendance Sheet!"+currentColumn+mem.getRowNumber()+":"+currentColumn+mem.getRowNumber();
        	        ValueRange timeList = service.spreadsheets().values()
        	            .get(spreadsheetId, timeListRange)
        	            .execute();
        	        List<List<Object>> timeValues = timeList.getValues();
        	        if (timeValues == null || timeValues.size() == 0) {
        	        	Date time = new Date();
        	    		DateFormat dateFormatTime = new SimpleDateFormat("HH:mm");
        	    		String newTime = dateFormatTime.format(time);
        	    		List<List<Object>> timeVal = Arrays.asList(
        						Arrays.asList((Object) (String) (newTime))
        							// Additional rows ...
        							);
        							ValueRange requestBodyForTime = new ValueRange()
        							.setValues(timeVal);
        			    	Sheets.Spreadsheets.Values.Update requestForTime =
        			    	service.spreadsheets().values().update(spreadsheetId, "Attendance Sheet!"+currentColumn+mem.getRowNumber()+":"+currentColumn+mem.getRowNumber(), requestBodyForTime);
        			    	requestForTime.setValueInputOption("USER_ENTERED");

        			    	UpdateValuesResponse responseTime = requestForTime.execute();
        			    	
        	    		
        	        } else {
        	        	String nextColumn = "";
        	        	if (currentColumn.substring(currentColumn.length()-1, currentColumn.length()).equalsIgnoreCase("Z")) {
        	        		nextColumn = "AA";
        	        	} else {
        	        		int charValue = currentColumn.charAt(currentColumn.length()-1);
            		    	if (currentColumn.length()>1) {
            		    		nextColumn = currentColumn.substring(0, currentColumn.length()-1) + String.valueOf((char) (charValue + 1));
            		    		} else {
            		    			nextColumn = String.valueOf((char) (charValue + 1));
            		    		}
        	        	}
        		    
           	        	Date time = new Date();
        	    		DateFormat dateFormatTime = new SimpleDateFormat("HH:mm");
        	    		String newTime = dateFormatTime.format(time);
        	    		List<List<Object>> timeVal = Arrays.asList(
        						Arrays.asList((Object) (String) (newTime))
        							// Additional rows ...
        							);
        							ValueRange requestBodyForTime = new ValueRange()
        							.setValues(timeVal);
        			    	Sheets.Spreadsheets.Values.Update requestForTime =
        			    	service.spreadsheets().values().update(spreadsheetId, "Attendance Sheet!"+nextColumn+mem.getRowNumber()+":"+nextColumn+mem.getRowNumber(), requestBodyForTime);
        			    	requestForTime.setValueInputOption("USER_ENTERED");

        			    	UpdateValuesResponse responseTime = requestForTime.execute();
        			    	
        	        	
        	        }
        		}
        	}
        	
        	if (nameInput.equalsIgnoreCase("quit")) {
        		takingRoll = false;
        	}
        	}
        }
        input.close();
    }
    
    public static void printRoster(ArrayList<Member> roster) {
    	for (Member mem : roster) {
    		System.out.println("Member name: " + mem.getName() + "\t Member row number: " + mem.getRowNumber());
    	}
    }


}