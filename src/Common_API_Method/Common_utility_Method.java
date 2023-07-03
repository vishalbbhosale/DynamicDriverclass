package Common_API_Method;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Common_utility_Method {
	
	public static void EvidanceCreator(String Filename,String RequestBody,String ResponseBody,int Statuscode) throws IOException {
		File vishal= new File("C:\\Users\\SURESH SHELKE\\OneDrive\\Desktop\\evidance\\"+Filename+".txt");
		System.out.println("new blank text file:"+vishal.getName());
		
		FileWriter datawrite=new FileWriter(vishal);
		datawrite.write("Request body is;"+RequestBody+"\n\n");
		datawrite.write("Status code is;"+Statuscode+"\n\n");
		datawrite.write("Response body is;"+ResponseBody);
		
		datawrite.close();
		System.out.println("data is writen in text file:"+vishal.getName());
			
		
	}
	
	public static ArrayList<String> ReadDataExcel(String sheetname,String TestCaseName) throws IOException
	{
	
	ArrayList<String> ArrayData=new ArrayList<String>();
	
	//create the object of file input stream to locate the excel file
	FileInputStream Fis= new FileInputStream("C:\\Selenium\\Book2.xlsx");
	
	//open the excel file by creating the object XSSFSWorkbook
	XSSFWorkbook WorkBook= new XSSFWorkbook(Fis);
	
	//open the desired 
	int countofsheet=WorkBook.getNumberOfSheets();
	for(int i=0;i<countofsheet;i++) {
		String Sheetname= WorkBook.getSheetName(i);
		
		//access the desired sheet
		if(Sheetname.equalsIgnoreCase(Sheetname))
		 {
			            //use xssf sheet to save the sheet into the variable
			XSSFSheet Sheet=WorkBook.getSheetAt(i);
			
			//create iterator to iterate through row and find out in which column the test case name are found
			Iterator<Row> Rows=Sheet.iterator();
			Row FirstRow=Rows.next();
			
			//create the iterator to iterate through the cells of first row to find out with cell contnts test case name
			Iterator<Cell> CellsofFirstRow= FirstRow.cellIterator();
			int k=0;
			int TC_Column=0;
			while(CellsofFirstRow.hasNext())
			{
				Cell CellValue= CellsofFirstRow.next();
						if(CellValue.getStringCellValue().equalsIgnoreCase("TestCasename"))
						{
							TC_Column=k;
							//System.out.println("expected column for test case name:" +k);
							break;
						}
						    k++;
			}
			//verify the row where the desired test case is found and fetch the entire row
			while(Rows.hasNext())
			{           
				   Row DataRow =Rows.next();
				   String TCName=DataRow.getCell(TC_Column).getStringCellValue();
				   if(TCName.equalsIgnoreCase(TestCaseName))
				   {
					   Iterator<Cell> CellValues = DataRow.cellIterator();
					   while(CellValues.hasNext())
					   {
						   String Data=CellValues.next().getStringCellValue();
						   ArrayData.add(Data);
						   
					   }
					   break;
				   }
				   
				}
			}
		 }
	return ArrayData;

}

}
