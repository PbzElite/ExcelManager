import java.util.Iterator;  
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.Row;  
import org.apache.poi.xssf.usermodel.XSSFSheet;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  
import java.io.*;   
import org.apache.poi.hssf.usermodel.HSSFWorkbook;   
import org.apache.poi.ss.usermodel.Workbook; 
import java.util.ArrayList;
import  java.io.*;  
import  org.apache.poi.hssf.usermodel.HSSFSheet;  
import  org.apache.poi.hssf.usermodel.HSSFWorkbook;  
import  org.apache.poi.hssf.usermodel.HSSFRow;  

public class App {
    public static void main(String[] args)   
    {  
        try  
        {  
            File file = new File("C:\\JavaPrograms\\ExcelManager\\Book1.xlsx");   //creating a new file instance  
            FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file  
            //creating Workbook instance that refers to .xlsx file  
            XSSFWorkbook rwb = new XSSFWorkbook(fis);   

            XSSFWorkbook wwb = new XSSFWorkbook();
            XSSFSheet easyship = wwb.createSheet("EasyShip");
            

            XSSFSheet sheet = rwb.getSheetAt(0);     //creating a Sheet object to retrieve object  
            Iterator<Row> itr = sheet.iterator();    //iterating over excel file  
            int rownu = sheet.getLastRowNum();
            
            ArrayList<Row> rowarr = new ArrayList<Row>();

            Row r = easyship.createRow((short)0);
            rowarr.add(r);
            for(int i = 0;i<rownu;i++){
                rowarr.add(i+1,easyship.createRow((short)i));
            }   

            rowarr.get(0).createCell(0).setCellValue("Choose Courier by");  
            rowarr.get(0).createCell(1).setCellValue("Shipping Insurance");  
            rowarr.get(0).createCell(2).setCellValue("Taxes & Duties Paid by*");  
            rowarr.get(0).createCell(3).setCellValue("Platform");  
            rowarr.get(0).createCell(4).setCellValue("Platform Order Number");  
            rowarr.get(0).createCell(5).setCellValue("Order Tags");
            rowarr.get(0).createCell(6).setCellValue("Receiver's Full Name*");
            rowarr.get(0).createCell(7).setCellValue("Receiver's Phone Number*");
            rowarr.get(0).createCell(8).setCellValue("Receover's Email");
            rowarr.get(0).createCell(9).setCellValue("Receiver's Tax ID");
            rowarr.get(0).createCell(10).setCellValue("Receiver's Address Line 1*");
            rowarr.get(0).createCell(11).setCellValue("Receiver's Address Line 2");
            rowarr.get(0).createCell(12).setCellValue("Receiver's Postal Code*");
            rowarr.get(0).createCell(13).setCellValue("Receiver's City*");
            rowarr.get(0).createCell(14).setCellValue("Receiver's State/Province");
            rowarr.get(0).createCell(15).setCellValue("Receiver's Country*");
            rowarr.get(0).createCell(16).setCellValue("Item Length (in)*");
            rowarr.get(0).createCell(17).setCellValue("Item Width (in)*");
            rowarr.get(0).createCell(18).setCellValue("Item Height (in)*");
            rowarr.get(0).createCell(19).setCellValue("Item Weight (lb)*");
            rowarr.get(0).createCell(20).setCellValue("Item Category/HS codes*");
            rowarr.get(0).createCell(21).setCellValue("Item Contains Liquid");
            rowarr.get(0).createCell(22).setCellValue("Item Contains Battery");
            rowarr.get(0).createCell(23).setCellValue("Item Description*");
            rowarr.get(0).createCell(24).setCellValue("Item Country of Origin");
            rowarr.get(0).createCell(25).setCellValue("Item SKU");
            rowarr.get(0).createCell(26).setCellValue("Item Customs Value*");
            rowarr.get(0).createCell(27).setCellValue("Item Customs Value Currency*");
            rowarr.get(0).createCell(28).setCellValue("Item Quantity");
            rowarr.get(0).createCell(29).setCellValue("Buyer's Notes");
            rowarr.get(0).createCell(30).setCellValue("Seller's Notes");
            FileOutputStream fileOut = new FileOutputStream("C:\\JavaPrograms\\ExcelManager\\EasyShip.xlsx");  
            wwb.write(fileOut);  

            while(itr.hasNext()){
                Row r1 = itr.next();
                Iterator<Cell> cells = r1.cellIterator();
                while(cells.hasNext()){
                    Cell temp = cells.next();

                    switch(temp.getCellType()){
                        case Cell.CELL_TYPE_STRING:
                            if(temp.getStringCellValue() == "Order #"){
                                int r2 = temp.getRowIndex();
                                int r3 = temp.getColumnIndex();
                                
                                rowarr.get(r2).createCell(r3).setCellValue(r1.getCell(r3).getNumericCellValue());
                            }
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            break;
                    }
                }
            }
            //closing the Stream  
            fileOut.close();  
            //closing the workbook  
            wwb.close();  
            //prints the message on the console  
            System.out.println("Excel file has been generated successfully.");  
            
            /*
            while (itr.hasNext())                 
            {  
                Row row = itr.next();  
                Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column  
                while (cellIterator.hasNext())   
                {  
                    Cell cell = cellIterator.next();  
                    switch (cell.getCellType())               
                    {  
                        case Cell.CELL_TYPE_STRING:    //field that represents string cell type  
                        System.out.print(cell.getStringCellValue() + "|");  
                        if(cell.getStringCellValue() == "Order #"){
                            Iterator<Cell> inner = row.cellIterator();
                            int row1 = sheet.getRow(0).getLastCellNum();
                            int col1 = sheet.getRow(0).getLastCellNum();

                            for(int i = 0;i<col1;i++){
                                System.out.println("Cell");
                            }
                        }
                        break;  
                        case Cell.CELL_TYPE_NUMERIC:    //field that represents number cell type  
                        System.out.print(cell.getNumericCellValue() + "|");  
                        break;  
                        default:  
                    }
                    colnum++;  
                }  
                System.out.println(""); 
                rownum++; 
            } 
            
            for(int i = 0;i<rownum;i++){
               System.out.println("row");
               int cell = 0; 
            }

            */
        }  
        catch(Exception e)  
        {  
            e.printStackTrace();  
        }  
    }   

    public static ArrayList<Row> addHeadings(ArrayList<Row> rowarr){
        rowarr.get(0).createCell(0).setCellValue("Choose Courier by");  
        rowarr.get(0).createCell(1).setCellValue("Shipping Insurance");  
        rowarr.get(0).createCell(2).setCellValue("Taxes & Duties Paid by*");  
        rowarr.get(0).createCell(3).setCellValue("Platform");  
        rowarr.get(0).createCell(4).setCellValue("Platform Order Number");  
        rowarr.get(0).createCell(5).setCellValue("Order Tags");
        rowarr.get(0).createCell(6).setCellValue("Receiver's Full Name*");
        rowarr.get(0).createCell(7).setCellValue("Receiver's Phone Number*");
        rowarr.get(0).createCell(8).setCellValue("Receover's Email");
        rowarr.get(0).createCell(9).setCellValue("Receiver's Tax ID");
        rowarr.get(0).createCell(10).setCellValue("Receiver's Address Line 1*");
        rowarr.get(0).createCell(11).setCellValue("Receiver's Address Line 2");
        rowarr.get(0).createCell(12).setCellValue("Receiver's Postal Code*");
        rowarr.get(0).createCell(13).setCellValue("Receiver's City*");
        rowarr.get(0).createCell(14).setCellValue("Receiver's State/Province");
        rowarr.get(0).createCell(15).setCellValue("Receiver's Country*");
        rowarr.get(0).createCell(16).setCellValue("Item Length (in)*");
        rowarr.get(0).createCell(17).setCellValue("Item Width (in)*");
        rowarr.get(0).createCell(18).setCellValue("Item Height (in)*");
        rowarr.get(0).createCell(19).setCellValue("Item Weight (lb)*");
        rowarr.get(0).createCell(20).setCellValue("Item Category/HS codes*");
        rowarr.get(0).createCell(21).setCellValue("Item Contains Liquid");
        rowarr.get(0).createCell(22).setCellValue("Item Contains Battery");
        rowarr.get(0).createCell(23).setCellValue("Item Description*");
        rowarr.get(0).createCell(24).setCellValue("Item Country of Origin");
        rowarr.get(0).createCell(25).setCellValue("Item SKU");
        rowarr.get(0).createCell(26).setCellValue("Item Customs Value*");
        rowarr.get(0).createCell(27).setCellValue("Item Customs Value Currency*");
        rowarr.get(0).createCell(28).setCellValue("Item Quantity");
        rowarr.get(0).createCell(29).setCellValue("Buyer's Notes");
        rowarr.get(0).createCell(30).setCellValue("Seller's Notes");
        FileOutputStream fileOut = new FileOutputStream("C:\\JavaPrograms\\ExcelManager\\EasyShip.xlsx");  
        wwb.write(fileOut);  

        return rowarr;
    }
}
