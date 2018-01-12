

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
import java.util.StringTokenizer;

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

    /** Global instance of the scopes required.
     *
     * If modifying these scopes, delete your previously saved credentials
     * 
     */
    private static final List<String> SCOPES =
        Arrays.asList(SheetsScopes.SPREADSHEETS);
    
    //list of Members
    static ArrayList<Member> roster = new ArrayList<>();
    
    //row number of the official meeting times
    static int officialRow = 2;
    
    //column letter of important first columns
    static String attendanceStartCol;
    static String hourStartCol;
    
    static String currentColumn;
    static String nextCol;
    
    static String attendanceRateCol = "D";
    
    static String hourRateCol = "E";
    
    static String numOfDaysCol = "F";
    
    static String numOfHoursCol = "G";
    
    //sheet service reference for entire class to use
    static Sheets sheetService;
    
    static String sheetSpreadsheetId;
    
    //ArrayList to hold values of the columns of a Google Spreadsheet until "BZ"
    static ArrayList<String> columnArray;
    
    //the row the first member (including member "Team") is off from the top
    static int rowNumOffset = 2;

    static {
    	columnArray = getColumnArray();
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

    public static void main(String[] args) throws IOException, InterruptedException {
    	
        // Build a new authorized API client service.
        Sheets service = getSheetsService();
        
        //give the reference to the class field for outside use
        sheetService = service;

        // Get Old Date and Old Column
        String spreadsheetId = "125gAQVShcHUiMHXDQ1nYFQnweH91lPYE9aOgPs0mAVE";
        //give the ID to class field for further use
        sheetSpreadsheetId = spreadsheetId;
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
        
        //getting start column for attendance and hour rates
        String attendanceStartColRange = "Attendance Sheet!A8:A8";
        String hourStartColRange = "Attendance Sheet!A11:A11";
        ValueRange attendanceStartColList = service.spreadsheets().values()
            .get(spreadsheetId, attendanceStartColRange)
            .execute();
        List<List<Object>> attendanceStartColValue = attendanceStartColList.getValues();
        ValueRange hourStartColList = service.spreadsheets().values()
            .get(spreadsheetId, hourStartColRange)
            .execute();
        List<List<Object>> hourStartColValue = hourStartColList.getValues();
        attendanceStartCol = (String) attendanceStartColValue.get(0).get(0);
        hourStartCol = (String) hourStartColValue.get(0).get(0);
        
        //get the current date
        Date date = new Date();
		DateFormat dateFormat = new SimpleDateFormat("M/d/yy");
		String newDate = dateFormat.format(date);
		
		//writing new date if it is a new day and setting and writing new column
		//also giving new column value to currentColumn variable for further use
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

	    	System.out.println(response);
	    	System.out.println("Creating a new column for a new day!");
	    	currentColumn = columnArray.get(columnArray.indexOf(oldColumn)+2);
	    	
	    	//writing out the new column
	    	List<List<Object>> columnVal = Arrays.asList(
					Arrays.asList((Object) (String) (currentColumn))
						// Additional rows ...
						);
						ValueRange requestBodyForCol = new ValueRange()
						.setValues(columnVal);
		    	Sheets.Spreadsheets.Values.Update requestForCol =
		    	service.spreadsheets().values().update(spreadsheetId, "Attendance Sheet!A5:A5", requestBodyForCol);
		    	requestForCol.setValueInputOption("USER_ENTERED");

		    	UpdateValuesResponse responseCol = requestForCol.execute();
		    	System.out.println(responseCol);
		    	
		    	Sheets.Spreadsheets.Values.Update request2 =
		    	    	service.spreadsheets().values().update(spreadsheetId, "Attendance Sheet!" + currentColumn + "1:" + currentColumn + "1", requestBody);
		    	    	request2.setValueInputOption("USER_ENTERED");

		    	    	UpdateValuesResponse response2 = request2.execute();

		    	    	// TODO: Change code below to process the `response` object:
		    	    	System.out.println(response2);
		}
		
		//getting the next column
		nextCol = columnArray.get(columnArray.indexOf(currentColumn)+1);
		
		//create an arrayList of members based on spreadsheet
        String nameListRange = "Attendance Sheet!C" + rowNumOffset + ":C";
        ValueRange nameList = service.spreadsheets().values()
            .get(spreadsheetId, nameListRange)
            .execute();
        List<List<Object>> nameValues = nameList.getValues();
        if (nameValues == null || nameValues.size() == 0) {
            System.out.println("No data found.");
        } else {
          for (int i = 0; i <nameValues.size(); i++) {
            roster.add(new Member((String) nameValues.get(i).get(0), i+rowNumOffset));
          }
          
        }
        
        printRoster();
		
		//refresh any that needs to be refreshed
		String refreshIndicatorRange = "Attendance Sheet!B" + rowNumOffset + ":B";
		ValueRange refreshList = service.spreadsheets().values()
	            .get(spreadsheetId, refreshIndicatorRange)
	            .execute();
	        List<List<Object>> refreshIndicatorValues = refreshList.getValues();
	        if (refreshIndicatorValues == null || refreshIndicatorValues.size() == 0) {
	            System.out.println("No refresh needed.");
	        } else {
	          for (int i = 0; i <refreshIndicatorValues.size(); i++) {
	        	  if (refreshIndicatorValues.get(i).size() > 0 && ((String)(refreshIndicatorValues.get(i).get(0))).length()!=0) {
	        		  System.out.println("Refreshing " + roster.get(i).getName());
	        		  refresh(roster.get(i).getRowNumber());
	        		  
	        		//reset the cell to blank
	        		  List<List<Object>> blankList = Arrays.asList(
      						Arrays.asList((Object) ""));
	    	          
	    	  					ValueRange requestBodyForBlanks = new ValueRange()
	    	  					.setValues(blankList);
	    	  	    	Sheets.Spreadsheets.Values.Update requestForBlanks =
	    	  	    	sheetService.spreadsheets().values().update(sheetSpreadsheetId, "Attendance Sheet!B"+roster.get(i).getRowNumber()+":B"+roster.get(i).getRowNumber(), requestBodyForBlanks);
	    	  	    	requestForBlanks.setValueInputOption("USER_ENTERED");
	    	  	    	UpdateValuesResponse responseBlanks = requestForBlanks.execute();
	    	  	    	System.out.println(responseBlanks);
	        		
	        	  }
	            
	          }
	          
	        }
        
        //taking roll
	    //enter a time for a person by typing his or her name followed by a return key
	    //refresh all members' statistics by typing "refresh" followed by a return key
	    //refresh a particular member's statistics by typing in "refresh" followed by his or her name, followed by an enter key
	    //quit by typing "quit"
        boolean takingRoll = true;
        Scanner input = new Scanner(System.in);
        System.out.println("Log in/out by typing your first name as listed on the attendance sheet (not case-sensitive) on the Google spreadsheet and hitting the enter key: ");
        while (takingRoll) {
        	if (input.hasNext()) {
        	String nameInput = input.nextLine().trim();
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
        			    	
        			    	System.out.println("Entering a time in for " + mem.getName());

        			    	UpdateValuesResponse responseTime = requestForTime.execute();
        			    	System.out.println(responseTime);
        			    	
        			    	
        	    		
        	        } else {
        	        	
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
        			    	service.spreadsheets().values().update(spreadsheetId, "Attendance Sheet!"+nextCol+mem.getRowNumber()+":"+nextCol+mem.getRowNumber(), requestBodyForTime);
        			    	requestForTime.setValueInputOption("USER_ENTERED");
        			    	
        			    	System.out.println("Entering a time out for " + mem.getName());

        			    	UpdateValuesResponse responseTime = requestForTime.execute();
        			    	System.out.println(responseTime);
        			    	
        			    	System.out.println("Refreshing " + mem.getName());
        			    	refresh(mem.getRowNumber());
        			    	
        	        	
        	        }
        		}
        	}
        	
        	if (nameInput.length()>=7 && nameInput.substring(0, 7).equalsIgnoreCase("refresh")) {
        		if (nameInput.length()==7) {
        			System.out.println("Refreshing all...");
        			for (Member mem : roster) {
        				refresh(mem.getRowNumber());
        				//the below does not seem to help - this is necessary to avoid google api from causing a resource exhausted error (when the rate of writing to sheets is exceeded)
        				//Thread.sleep(500);
        			}
        		} else {
        			for (Member mem : roster) {
        				if (mem.getName().equalsIgnoreCase(nameInput.substring(8))) {
        					System.out.println("Refreshing " + mem.getName());
        					refresh(mem.getRowNumber());
        				}
        			}
        		}
        	}
        	
        	if (nameInput.equalsIgnoreCase("quit")) {
        		System.out.println("Quitting...");
        		takingRoll = false;
        	}
        	
        	if (takingRoll) {
        		System.out.println("Log in/out by typing your first name as listed on the attendance sheet (not case-sensitive) on the Google spreadsheet and hitting the enter key: ");
        	}
        }
        }
        input.close();
    }
    
    public static void printRoster() {
    	for (Member mem : roster) {
    		System.out.println("Member name: " + mem.getName() + "\t Member row number: " + mem.getRowNumber());
    	}
    }
    
    public static void refresh(int rowNum) throws IOException {
    	double attendanceRate = ((double) getTotalDays(rowNum))/getTotalDays(officialRow);
    	double hourRate = ((double) getTotalHours(rowNum))/getTotalHours(officialRow);
    	String attendanceRateString = attendanceRate*100 + "%";
    	String hourRateString = hourRate*100 + "%";
    	int numOfDays = getTotalDays(rowNum);
    	double numOfHours = getTotalHours(rowNum);
    	
    	System.out.println("Printing out stats for row number " + rowNum + "...");
    	List<List<Object>> stats = Arrays.asList(
				Arrays.asList((Object)attendanceRateString, (Object)hourRateString, (Object)numOfDays, (Object)numOfHours)
					// Additional rows ...
					);
					ValueRange requestBodyForStats = new ValueRange()
					.setValues(stats);
	    	Sheets.Spreadsheets.Values.Update requestForStats =
	    	sheetService.spreadsheets().values().update(sheetSpreadsheetId, "Attendance Sheet!"+attendanceRateCol+rowNum+":"+numOfHoursCol+rowNum, requestBodyForStats);
	    	requestForStats.setValueInputOption("USER_ENTERED");
	    	UpdateValuesResponse responseStats = requestForStats.execute();
	    	System.out.println(responseStats);
	    	
    	
    }
    
    public static int getTotalDays(int rowNum) throws IOException {
    	ValueRange dayList = sheetService.spreadsheets().values()
	            .get(sheetSpreadsheetId, "Attendance Sheet!"+attendanceStartCol+rowNum+":"+nextCol+rowNum)
	            .execute();
	        List<List<Object>> dayValues = dayList.getValues();
	        int dayNum = 0;
	        if (dayValues != null && dayValues.size()!=0) {
	        	for (List<Object> list : dayValues) {
	        		for (int i = 0; i < list.size(); i+=2) {
	        			if (((String) list.get(i)).length()!=0) {
	        				dayNum++;
	        			}
	        		}
	        	}
	        	return dayNum;
	        } else {
	        	System.out.println("No data found for row number " + rowNum);
	        	return 0;
	        }
    	
    }
    
    public static double getTotalHours(int rowNum) throws IOException {
    	ValueRange hourList = sheetService.spreadsheets().values()
	            .get(sheetSpreadsheetId, "Attendance Sheet!"+hourStartCol+rowNum+":"+nextCol+rowNum)
	            .execute();
	        List<List<Object>> hourValues = hourList.getValues();
	        double hourNum = 0;
	        if (hourValues != null && hourValues.size()!=0) {
	        	for (List<Object> list : hourValues) {
		        	for (int i = 0; i+1 < list.size(); i+=2) {
		        		if (((String) list.get(i)).length()!=0 && ((String) list.get(i+1)).length()!=0) {
		        			String timeIn = (String) list.get(i);
		        			String timeOut = (String) list.get(i+1);
		        			
		        			StringTokenizer str1 = new StringTokenizer(timeIn, ":");
		        			StringTokenizer str2 = new StringTokenizer(timeOut, ":");
		        			int fullHours = 0;
		        			double partialHours = 0;
		        			try {
		        				fullHours = Integer.parseInt(str2.nextToken())-Integer.parseInt(str1.nextToken());
		        				partialHours = Double.parseDouble(str2.nextToken())/60-Double.parseDouble(str1.nextToken())/60;
		        			} catch (NumberFormatException e) {
		        				System.out.println("An NumberFormatException has occurred trying to parse time entry in hour rate slot " + i+1 + "...");
		        				System.out.println("Will count the hours in the " + i+1 + " hour rate slot to be 0!");
		        				fullHours = 0;
		        				partialHours = 0;
		        			}
		        			
		        			hourNum+=fullHours;
		        			hourNum+=partialHours;
		        		}
		        	}
		        }
	        	return hourNum;
	        } else {
	        	System.out.println("No data found for row number " + rowNum);
	        	return 0;
	        }
    	
    }
    
    public static ArrayList<String> getColumnArray() {
    	ArrayList<String> columnArray = new ArrayList<String>();
    	for (int i = 65; i <= 90; i++) {
    		columnArray.add(Character.toString((char)i));
    	}
    	for (int i = 65; i <= 67; i++) {
    		for (int k = 65; k <= 90; k++) {
    			columnArray.add(Character.toString((char)i) + Character.toString((char)k));
    		}
    	}
    	return columnArray;
    	
    }


}