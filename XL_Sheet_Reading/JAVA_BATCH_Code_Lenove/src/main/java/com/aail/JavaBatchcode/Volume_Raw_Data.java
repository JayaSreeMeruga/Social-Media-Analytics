package com.aail.JavaBatchcode;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
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

import com.monitorjbl.xlsx.StreamingReader;

public class Volume_Raw_Data {

   public static final String SAMPLE_XLSX_FILE_PATH = "MBG Contact Volume  Warranty Install Base - 20-Feb-2018 00-53-03.xlsx";


    public static void main(String[] args) throws IOException, InvalidFormatException, SQLException, ClassNotFoundException {
    	Connection c = null;
		Statement stmt = null;
        DataFormatter dataFormatter = new DataFormatter();

		InputStream is = new FileInputStream(new File("MBG Contact Volume  Warranty Install Base - 20-Feb-2018 00-53-03.xlsx"));
		StreamingReader reader = StreamingReader.builder()
		        .rowCacheSize(100)    // number of rows to keep in memory (defaults to 10)
		        .bufferSize(4096)     // buffer size to use when reading InputStream to file (defaults to 1024)
		        .sheetIndex(2)        // index of sheet to use (defaults to 0)
		        .read(is);            // InputStream or File for XLSX file (required)

		int year = 0;
		String month="",vendor="",site="",region="",country="",nature="",product="",support_channel="";
		int volume_handled_l1_only=0;
		 
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
	        String sql="Insert into cc_mbg_contactcenterdata(year,month,vendor,site,region,country,nature,product,support_channel,volume_handled_l1_only) values(?,?,?,?,?,?,?,?,?,?)";
	        ps=c.prepareStatement(sql);
		
       
	        for (Row r : reader) {
				if(r.getRowNum()==0) 
				{
					continue;
				}
				
								
				 if ( r.getCell(0)==null){
					 year=0;
		            }
				 else
				 {
				
				year = (int) r.getCell(0).getNumericCellValue();// Get the Cell at the Index / Column you want.
	            System.out.print(year +"\t");
				 }
				 
				 if (r.getCell(1)==null){
					 month=" ";
		               }
		               else
		               {
		            	   month = r.getCell(1).getStringCellValue();
		               System.out.print(month+"\t");
		               } 
				 
				 if (r.getCell(2)==null){
					 vendor=" ";
		               }
		               else
		               {
		            	   vendor = r.getCell(2).getStringCellValue();
		               System.out.print(vendor+"\t");
		               } 
				 
				 
				 
				 if (r.getCell(3)==null){
					 site=" ";
		               }
		               else
		               {
		            	   site = r.getCell(3).getStringCellValue();
		               System.out.print(site+"\t");
		               } 
				 if (r.getCell(4)==null){
					 region=" ";
		               }
		               else
		               {
		            	   region = r.getCell(4).getStringCellValue();
		               System.out.print(region+"\t");
		               } 
				 
				 
				 if (r.getCell(5)==null){
					 country=" ";
		               }
		               else
		               {
		            	   country = r.getCell(5).getStringCellValue();
		               System.out.print(country+"\t");
		               } 
				 if (r.getCell(6)==null){
					 nature=" ";
		               }
		               else
		               {
		            	   nature = r.getCell(6).getStringCellValue();
		               System.out.print(nature+"\t");
		               } 
				 
				 if (r.getCell(7)==null){
					 product=" ";
		               }
		               else
		               {
		            	   product = r.getCell(7).getStringCellValue();
		               System.out.print(product+"\t");
		               } 
				 if (r.getCell(8)==null){
		            	   support_channel=" ";
		               }
		               else
		               {
		            	   support_channel = r.getCell(8).getStringCellValue();
		               System.out.print(support_channel+"\t");
		               } 
				 
				 
				 if ( r.getCell(9)==null){
					 volume_handled_l1_only=0;
		            }
				 else
				 {
				
				volume_handled_l1_only = (int) r.getCell(0).getNumericCellValue();// Get the Cell at the Index / Column you want.
	            System.out.print(volume_handled_l1_only +"\t");
				 }
		            ps.setInt(1, year);
		            ps.setString(2, month);
		            ps.setString(3, vendor);
		            ps.setString(4, site);
		            ps.setString(5, region);
		            ps.setString(6, country);
		            ps.setString(7, nature);
		            ps.setString(8, product);
		            ps.setString(9, support_channel);
		            ps.setInt(10, volume_handled_l1_only);
		           

		            ps.executeUpdate();
		            //  System.out.println("Values Inserted Successfully");
		              c.commit();
		      		//c.close(); 

			}     

            
           }

        	
    	catch(Exception e)
    	{
    		System.out.println("Exeption is occured:"+e);
    	}

    } //main close

} //class close
