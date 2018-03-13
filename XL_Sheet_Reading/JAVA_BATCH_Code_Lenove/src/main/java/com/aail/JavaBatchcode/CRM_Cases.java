package com.aail.JavaBatchcode;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Properties;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class CRM_Cases {
	
   public static final String SAMPLE_XLSX_FILE_PATH = "CRM Cases - 26-Feb-2018 12-58-19.xlsx";


    public static void main(String[] args) throws IOException, InvalidFormatException, SQLException, ClassNotFoundException {
    	Connection c = null;
		Statement stmt = null;


    	/*while(true)
    	{*/
    	try
    	{
		
    	Class.forName("org.postgresql.Driver");  

		Properties prop = new Properties();
		File file = new File("Jdbcconnection.properties");
		// System.out.println(file.getAbsolutePath());
		FileInputStream input = new FileInputStream(file.getAbsolutePath());
		prop.load(input);
		String ip_address = prop.getProperty("ip_address");
		String database_name = prop.getProperty("database_name");
		String user_name=prop.getProperty("user_name");
		//String password=prop.getProperty("password");
		String db_url="jdbc:postgresql://"+ip_address+"/"+database_name+"?user="+user_name;//+"&password="+password;
		input.close();

		System.out.println("Opened database successfully");
		
		c = DriverManager.getConnection(db_url);
		c.setAutoCommit(false);
    	
		// Statement stmt=c.createStatement();
	        PreparedStatement ps=null;
	        String sql="Insert into cc_dcg_thinkserver_premier(serial_number,machine_type,warranty_status,warranty_start_date,warranty_end_date,so_number,think_agile,premier_support,assigned_to,complaint_code,severity,subject,company,repair_action_status,ticket_number,azure_subscription,microsoft_id,ir_number,date_so_opened,date_so_closed,file_name) values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
	        ps=c.prepareStatement(sql);
		
	        System.out.println("This is main");
        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));
        System.out.println("This is main33");

        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

        String file_name=SAMPLE_XLSX_FILE_PATH;

        // use a for-each loop
        System.out.println("Retrieving Sheets using for-each loop");
        for(Sheet sheet: workbook) {
            System.out.println("=> " + sheet.getSheetName());
        }

   
        /*
           ==================================================================
           Iterating over all the rows and columns in a Sheet (Multiple ways)
           ==================================================================
        */

        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);

        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();


        String  WarrantyStartDate="",SerialNumber="",MachineType="",WarrantyStatus="",warranty_end_date="",so_number="",think_agile="";
        String  premier_support="",assigned_to="",complaint_code="",severity="",subject="",company="",repair_action_status="",ticket_number="";
        String azure_subscription="",microsoft_id="",ir_number="",date_so_opened="",date_so_closed="";
        
	for (Row row : sheet) { // For each Row.
            if(row.getRowNum() > 0)
            {
	
	     if (row.getCell(0)==null){
		SerialNumber=" ";
	     }
	     else {
             SerialNumber = row.getCell(0).getStringCellValue(); // Get the Cell at the Index / Column you want.
            System.out.print(SerialNumber +"\t");
            }

	     if (row.getCell(1)==null){
		MachineType=" ";
	     }
	     else {	
            Cell cell = row.getCell(1);
             MachineType= dataFormatter.formatCellValue(cell); 
            System.out.print(MachineType+"\t");
            }

	     if (row.getCell(2)==null){
		WarrantyStatus=" ";
	     }
	     else {
             WarrantyStatus= row.getCell(2).getStringCellValue();
            System.out.print(WarrantyStatus+"\t");
	    }


            if (row.getCell(3)==null){
             WarrantyStartDate=" ";
            }
            else
            {
            	WarrantyStartDate = row.getCell(3).getStringCellValue();
            System.out.print(WarrantyStartDate+"\t");
            }
            if (row.getCell(4)==null){
            	warranty_end_date=" ";
               }
               else
               {
            	   warranty_end_date = row.getCell(4).getStringCellValue();
               System.out.print(warranty_end_date+"\t");
               }
            if (row.getCell(5)==null){
            	so_number=" ";
               }
               else
               {
            	   so_number = row.getCell(5).getStringCellValue();
               System.out.print(so_number+"\t");
               }    
            if (row.getCell(6)==null){
            	think_agile=" ";
               }
               else
               {
            	   think_agile = row.getCell(6).getStringCellValue();
               System.out.print(think_agile+"\t");
               }       
            if (row.getCell(7)==null){
            	premier_support=" ";
               }
               else
               {
            	   premier_support = row.getCell(7).getStringCellValue();
               System.out.print(premier_support+"\t");
               } 
            if (row.getCell(8)==null){
            	assigned_to=" ";
               }
               else
               {
            	   assigned_to = row.getCell(8).getStringCellValue();
               System.out.print(assigned_to+"\t");
               } 
            if (row.getCell(9)==null){
            	complaint_code=" ";
               }
               else
               {
            	   complaint_code = row.getCell(9).getStringCellValue();
               System.out.print(complaint_code+"\t");
               } 
            if (row.getCell(10)==null){
            	severity=" ";
               }
               else
               {
            	   severity = row.getCell(10).getStringCellValue();
               System.out.print(severity+"\t");
               } 
            if (row.getCell(11)==null){
            	subject=" ";
               }
               else
               {
            	   subject = row.getCell(11).getStringCellValue();
               System.out.print(subject+"\t");
               }  
            if (row.getCell(12)==null){
            	company=" ";
               }
               else
               {
            	   company = row.getCell(12).getStringCellValue();
               System.out.print(company+"\t");
               } 
            if (row.getCell(13)==null){
            	repair_action_status=" ";
               }
               else
               {
            	   repair_action_status = row.getCell(13).getStringCellValue();
               System.out.print(repair_action_status+"\t");
               } 

	    if (row.getCell(14)==null){	
		ticket_number=" ";
	    }
	    else {	
            Cell cell1 = row.getCell(14);
            ticket_number= dataFormatter.formatCellValue(cell1); 
           System.out.print(ticket_number+"\t");
            }


	    if (row.getCell(15)==null){	
		azure_subscription=" ";
	    }
	    else {
           Cell cell3 = row.getCell(15);
           azure_subscription= dataFormatter.formatCellValue(cell3); 
          System.out.print(azure_subscription+"\t");
           }


	    if (row.getCell(16)==null){	
		microsoft_id=" ";
	    }
	    else {     
            Cell cell2 = row.getCell(16);
            microsoft_id= dataFormatter.formatCellValue(cell2); 
           System.out.print(microsoft_id+"\t");
	   }


           if (row.getCell(17)==null){
            	ir_number=" ";
               }
               else
               {
            	   ir_number = row.getCell(17).getStringCellValue();
               System.out.print(ir_number+"\t");
               } 
            if (row.getCell(18)==null){
            	date_so_opened=" ";
               }
               else
               {
            	   date_so_opened = row.getCell(18).getStringCellValue();
               System.out.print(date_so_opened+"\t");
               } 
            if (row.getCell(19)==null){
            	date_so_closed=" ";
               }
               else
               {
            	   date_so_closed = row.getCell(19).getStringCellValue();
               System.out.print(date_so_closed+"\t");
               } 
            
            ps.setString(1, SerialNumber);
            ps.setString(2, MachineType);
            ps.setString(3,WarrantyStatus);
            ps.setString(4, WarrantyStartDate);
            ps.setString(5, warranty_end_date);
            ps.setString(6, so_number);
            ps.setString(7, think_agile);
            ps.setString(8, premier_support);
            ps.setString(9, assigned_to);
            ps.setString(10, complaint_code);
            ps.setString(11, severity);
            ps.setString(12, subject);
            ps.setString(13, company);
            ps.setString(14, repair_action_status);
            ps.setString(15, ticket_number);
            ps.setString(16, azure_subscription);
            ps.setString(17, microsoft_id);
            ps.setString(18, ir_number);
            ps.setString(19, date_so_opened);
            ps.setString(20, date_so_closed);
            ps.setString(21, file_name);

        ps.executeUpdate();
      //  System.out.println("Values Inserted Successfully");
        c.commit();
		//c.close();  
            }
            System.out.println();

        }

        System.out.println("Values Inserted Successfully");

        workbook.close();
    	}
    	catch(Exception e)
    	{
    		System.out.println("Exeption is occured:"+e);
    	}

    } //main close

} //class close
