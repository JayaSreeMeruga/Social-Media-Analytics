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

public class Warranty_Install_Base {

   public static final String SAMPLE_XLSX_FILE_PATH = "MBG Contact Volume  Warranty Install Base - 20-Feb-2018 00-53-03.xlsx";


    public static void main(String[] args) throws IOException, InvalidFormatException, SQLException, ClassNotFoundException {
 
	Connection c = null;
	Statement stmt = null;
        DataFormatter dataFormatter = new DataFormatter();

	InputStream is = new FileInputStream(new File("MBG Contact Volume  Warranty Install Base - 20-Feb-2018 00-53-03.xlsx"));
		StreamingReader reader = StreamingReader.builder()
		        .rowCacheSize(100)    // number of rows to keep in memory (defaults to 10)
		        .bufferSize(4096)     // buffer size to use when reading InputStream to file (defaults to 1024)
		        .sheetIndex(1)        // index of sheet to use (defaults to 0)
		        .read(is);            // InputStream or File for XLSX file (required)


	String region="";String product="";
	
	int april_2015=0,may_2015,june_2015,july_2015,august_2015,septemeber_2015,october_2015,november_2015,december_2015;
	int january_2016=0,february_2016=0,march_2016=0,april_2016=0,may_2016=0,june_2016=0,july_2016=0,august_2016=0;
	int septemeber_2016=0,october_2016=0,november_2016=0,december_2016=0,january_2017=0,february_2017=0,march_2017=0;
	int april_2017=0,may_2017=0,june_2017=0,july_2017=0,august_2017=0,septemeber_2017=0,october_2017=0,november_2017=0;
	int december_2o17=0;
	

	try
    	{
		
    	Class.forName("org.postgresql.Driver");  

		Properties prop = new Properties();
		File file = new File("Jdbcconnection.properties");
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

	        
	        String sql="insert into CC_MBG_WarrantyInstallBase(region,product,"+"\"01-04-2015\""+", "+"\"01-05-2015\""+", "+"\"01-06-2015\""+", "+"\"01-07-2015\""+", "+"\"01-08-2015\""+", "+"\"01-09-2015\""+", "+"\"01-10-2015\""+", "+"\"01-11-2015\""+", "+"\"01-12-2015\""+", "+"\"01-01-2016\""+", "+"\"01-02-2016\""+", "+"\"01-03-2016\""+", "+"\"01-04-2016\""+", "+"\"01-05-2016\""+", "+"\"01-06-2016\""+", "+"\"01-07-2016\""+", "+"\"01-08-2016\""+", "+"\"01-09-2016\""+", "+"\"01-10-2016\""+", "+"\"01-11-2016\""+", "+"\"01-12-2016\""+", "+"\"01-01-2017\""+", "+"\"01-02-2017\""+", "+"\"01-03-2017\""+", "+"\"01-04-2017\""+", "+"\"01-05-2017\""+", "+"\"01-06-2017\""+", "+"\"01-07-2017\""+", "+"\"01-08-2017\""+", "+"\"01-09-2017\""+", "+"\"01-10-2017\""+", "+"\"01-11-2017\""+", "+"\"01-12-2017\""+") values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";
	        ps=c.prepareStatement(sql);
		
       
	        for (Row r : reader) {
				if(r.getRowNum()==0||r.getRowNum()==9||r.getRowNum()==18||r.getRowNum()==19||r.getRowNum()==20||r.getRowNum()==21||r.getRowNum()==22||r.getRowNum()==23||r.getRowNum()==24||r.getRowNum()==25) 
				{
					continue;
				}
			
				
				 if ( r.getCell(0)==null){
					 region=" ";
		            	}
				 else
				 {
				
				region = r.getCell(0).getStringCellValue();// Get the Cell at the Index / Column you want.
	            System.out.print(region +"\t");
				 }
				 

				 if (r.getCell(1)==null){
					 product=" ";
		               }
		               else
		               {
                           
		            	   product = r.getCell(1).getStringCellValue();
		               System.out.print(product+"\t");
                           
		               } 
				 

				 
				 if (r.getCell(2)==null){
					 april_2015=0;
		               }
		               else
		               {
		            	   april_2015 = (int) r.getCell(2).getNumericCellValue();
		               System.out.print("******** : "+april_2015+"\t");
		               } 
				 
				 
				 
				 if (r.getCell(3)==null){
					 may_2015=0;
		               }
		               else
		               {
		            	   may_2015 = (int) r.getCell(3).getNumericCellValue();
		               System.out.print(may_2015+"\t");
		               } 
				 
				
				 if (r.getCell(4)==null){
					 june_2015=0;
		               }
		               else
		               {
		            	   june_2015 = (int) r.getCell(4).getNumericCellValue();
		               System.out.print(june_2015+"\t");
		               } 

				 

				 if (r.getCell(5)==null){
					 july_2015=0;
		               }
		               else
		               {
		            	   july_2015 = (int) r.getCell(5).getNumericCellValue();
		               System.out.print(july_2015+"\t");
		               } 
				 if (r.getCell(6)==null){
					 august_2015=0;
		               }
		               else
		               {
		            	   august_2015 = (int) r.getCell(6).getNumericCellValue();
		               System.out.print(august_2015+"\t");
		               } 
				 if (r.getCell(7)==null){
					 septemeber_2015=0;
		               }
		               else
		               {
		            	   septemeber_2015 = (int) r.getCell(7).getNumericCellValue();
		               System.out.print(septemeber_2015+"\t");
		               } 
				 
				 
				 if (r.getCell(8)==null){
					 october_2015=0;
		               }
		               else
		               {
		            	   october_2015 = (int) r.getCell(8).getNumericCellValue();
		               System.out.print(october_2015+"\t");
		               } 
				 if (r.getCell(9)==null){
					 november_2015=0;
		               }
		               else
		               {
		            	   november_2015 = (int) r.getCell(9).getNumericCellValue();
		               System.out.print(november_2015+"\t");
		               } 
				 if (r.getCell(10)==null){
					 december_2015=0;
		               }
		               else
		               {
		            	   december_2015 = (int) r.getCell(10).getNumericCellValue();
		               System.out.print(december_2015+"\t");
		               } 
				
				 if (r.getCell(11)==null){
					 january_2016=0;
		               }
		               else
		               {
		            	   january_2016 = (int) r.getCell(11).getNumericCellValue();
		               System.out.print(january_2016+"\t");
		               } 
				 
				 if (r.getCell(12)==null){
					 february_2016=0;
		               }
		               else
		               {
		            	   february_2016 = (int) r.getCell(12).getNumericCellValue();
		               System.out.print(february_2016+"\t");
		               }  
				 
				 if (r.getCell(13)==null){
					 march_2016=0;
		               }
		               else
		               {
		            	   march_2016 = (int) r.getCell(13).getNumericCellValue();
		               System.out.print(march_2016+"\t");
		               }  
				 if (r.getCell(14)==null){
					 april_2016=0;
		               }
		               else
		               {
		            	   april_2016 = (int) r.getCell(14).getNumericCellValue();
		               System.out.print(april_2016+"\t");
		               } 
				 
				 if (r.getCell(15)==null){
					 may_2016=0;
		               }
		               else
		               {
		            	   may_2016 = (int) r.getCell(15).getNumericCellValue();
		               System.out.print(may_2016+"\t");
		               }  
				 
				 if (r.getCell(16)==null){
					 june_2016=0;
		               }
		               else
		               {
		            	   june_2016 = (int) r.getCell(16).getNumericCellValue();
		               System.out.print(june_2016+"\t");
		               } 
				 if (r.getCell(17)==null){
					 july_2016=0;
		               }
		               else
		               {
		            	   july_2016 = (int) r.getCell(17).getNumericCellValue();
		               System.out.print(july_2016+"\t");
		               }  
				 
				 if (r.getCell(18)==null){
					 august_2016=0;
		               }
		               else
		               {
		            	   august_2016 = (int) r.getCell(18).getNumericCellValue();
		               System.out.print(august_2016+"\t");
		               } 
				 
				 if (r.getCell(19)==null){
					 septemeber_2016=0;
		               }
		               else
		               {
		            	   septemeber_2016 = (int) r.getCell(19).getNumericCellValue();
		               System.out.print(septemeber_2016+"\t");
		               } 
				 if (r.getCell(20)==null){
					 october_2016=0;
		               }
		               else
		               {
		            	   october_2016 = (int) r.getCell(20).getNumericCellValue();
		               System.out.print(october_2016+"\t");
		               } 
				 
				 
				 
				 if (r.getCell(21)==null){
					 november_2016=0;
		               }
		               else
		               {
		            	   november_2016 = (int) r.getCell(21).getNumericCellValue();
		               System.out.print(november_2016+"\t");
		               } 
				 if (r.getCell(22)==null){
					 december_2016=0;
		               }
		               else
		               {
		            	   december_2016 = (int) r.getCell(22).getNumericCellValue();
		               System.out.print(december_2016+"\t");
		               } 
				 
				 if (r.getCell(23)==null){
					 january_2017=0;
		               }
		               else
		               {
		            	   january_2017 = (int) r.getCell(23).getNumericCellValue();
		               System.out.print(january_2017+"\t");
		               }  
				 if (r.getCell(24)==null){
					 february_2017=0;
		               }
		               else
		               {
		            	   february_2017 = (int) r.getCell(24).getNumericCellValue();
		               System.out.print(february_2017+"\t");
		               }  
				 if (r.getCell(25)==null){
					 march_2017=0;
		               }
		               else
		               {
		            	   march_2017 = (int) r.getCell(25).getNumericCellValue();
		               System.out.print(march_2017+"\t");
		               } 
				 
				 if (r.getCell(26)==null){
					 april_2017=0;
		               }
		               else
		               {
		            	   april_2017 = (int) r.getCell(26).getNumericCellValue();
		               System.out.print(april_2017+"\t");
		               } 
				 if (r.getCell(27)==null){
					 may_2017=0;
		               }
		               else
		               {
		            	   may_2017 = (int) r.getCell(27).getNumericCellValue();
		               System.out.print(may_2017+"\t");
		               } 
				 if (r.getCell(28)==null){
					 june_2017=0;
		               }
		               else
		               {
		            	   june_2017 = (int) r.getCell(28).getNumericCellValue();
		               System.out.print(june_2017+"\t");
		               } 
				 if (r.getCell(29)==null){
					 july_2017=0;
		               }
		               else
		               {
		            	   july_2017 = (int) r.getCell(29).getNumericCellValue();
		               System.out.print(july_2017+"\t");
		               } 
				 if (r.getCell(30)==null){
					 august_2017=0;
		               }
		               else
		               {
		            	   august_2017 = (int) r.getCell(30).getNumericCellValue();
		               System.out.print(august_2017+"\t");
		               } 
				 if (r.getCell(31)==null){
					 septemeber_2017=0;
		               }
		               else
		               {
		            	   septemeber_2017 = (int) r.getCell(31).getNumericCellValue();
		               System.out.print(septemeber_2017+"\t");
		               }
				 if (r.getCell(32)==null){
					 october_2017=0;
		               }
		               else
		               {
		            	   october_2017 = (int) r.getCell(32).getNumericCellValue();
		               System.out.print(october_2017+"\t");
		               } 
				 if (r.getCell(33)==null){
					 november_2017=0;
		               }
		               else
		               {
		            	   november_2017 = (int) r.getCell(33).getNumericCellValue();
		               System.out.print(november_2017+"\t");
		               } 
				 
				 if (r.getCell(34)==null){
					 december_2o17=0;
		               }
		               else
		               {
		            	   december_2o17 = (int) r.getCell(34).getNumericCellValue();
		               System.out.print(december_2o17+"\t");
		               } 
				 
		            ps.setString(1, region);
		            ps.setString(2, product);
		            ps.setInt(3, april_2015);
		            ps.setInt(4, may_2015);
		            ps.setInt(5, june_2015);
		            ps.setInt(6, july_2015);
		            ps.setInt(7, august_2015);
		            ps.setInt(8, septemeber_2015);
		            ps.setInt(9, october_2015);
		            ps.setInt(10, november_2015);
		            ps.setInt(11, december_2015);
		            ps.setInt(12, january_2016);
		            ps.setInt(13, february_2016);
		            ps.setInt(14, march_2016);
		            ps.setInt(15, april_2016);
		            ps.setInt(16, may_2016);
		            ps.setInt(17, june_2016);
		            ps.setInt(18, july_2016);
		            ps.setInt(19, august_2016);
		            ps.setInt(20, septemeber_2016);
		            ps.setInt(21, october_2016);
		            ps.setInt(22, november_2016);
		            ps.setInt(23, december_2016);
		            ps.setInt(24, january_2017);
		            ps.setInt(25, february_2017);
		            ps.setInt(26, march_2017);
		            ps.setInt(27, april_2017);
		            ps.setInt(28, may_2017);
		            ps.setInt(29, june_2017);
		            ps.setInt(30, july_2017);
		            ps.setInt(31, august_2017);
		            ps.setInt(32, septemeber_2017);
		            ps.setInt(33, october_2017);
		            ps.setInt(34, november_2017);
		            ps.setInt(35, december_2o17);
		            
		           

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
