import java.util.Iterator;   
import org.apache.poi.ss.usermodel.Row;  
import org.apache.poi.xssf.usermodel.XSSFSheet;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  
import java.io.*;   
import java.util.ArrayList;
import javax.swing.*;
import java.awt.event.*;
import javax.swing.filechooser.*;

// Main class
public class App extends JFrame implements ActionListener {
    public static String name;
    public static String directory;

    // Main driver method
    public static void main(String[] args)
    {
        // Creating instance of JFrame
        JFrame frame = new JFrame("BBP Enterprise EasyShip Formatter");
        JPanel panel = new JPanel();
       
        // Creating instances
        JButton button = new JButton("Choose File Path");
        JLabel l2 = new JLabel("Choose Name");
        JButton b2 = new JButton("Save");

        JLabel l = new JLabel("\nFile Directory");
        JTextField te = new JTextField(25);

        JLabel l3 = new JLabel("Item Length (in)");
        JTextField t = new JTextField(5);
        JLabel l4 = new JLabel("Item Width (in)");
        JTextField t1 = new JTextField(5);
        JLabel l5 = new JLabel("Item Heigth (in)");
        JTextField t2 = new JTextField(5);
        JLabel l6 = new JLabel("Item Weight (lb)");
        JTextField t3 = new JTextField(5);
        JLabel l7 = new JLabel("Item Category/HS codes");
        JTextField t4 = new JTextField(5);

        b2.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e){
            
                // if the user presses the save button show the save dialog
                String com = e.getActionCommand();
                
                if (com.equals("save")) {
                    // create an object of JFileChooser class
                    JFileChooser j = new JFileChooser("C:",FileSystemView.getFileSystemView());

                    // set the selection mode to directories only
                    j.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

                    // invoke the showsSaveDialog function to show the save dialog
                    int r = j.showSaveDialog(null);

                    if (r == JFileChooser.APPROVE_OPTION) {
                        // set the label to the path of the selected directory
                        l.setText(j.getSelectedFile().getAbsolutePath());
                    }
                    // if the user cancelled the operation
                    else
                        l.setText("the user cancelled the operation");
                }
                // if the user presses the open dialog show the open dialog
                else {
                    // create an object of JFileChooser class
                    JFileChooser j = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());

                    // set the selection mode to directories only
                    j.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

                    // invoke the showsOpenDialog function to show the save dialog
                    int r = j.showOpenDialog(null);

                    if (r == JFileChooser.APPROVE_OPTION) {
                        // set the label to the path of the selected directory
                        l.setText(j.getSelectedFile().getAbsolutePath());
                    }
                    // if the user cancelled the operation
                    else
                        l.setText("the user cancelled the operation");
                }
                name = te.getText();
                directory = l.getText();
            }
        });

        button.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e){
            try  
            { 
                File file = new File("C:\\JavaPrograms\\ExcelManager\\Book1.xlsx");   //creating a new file instance  
                FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file  
                //creating Workbook instance that refers to .xlsx file  
                XSSFWorkbook rwb = new XSSFWorkbook(fis);   
                XSSFWorkbook wwb = new XSSFWorkbook();
                
                //"EasyShip"                
                XSSFSheet easyship = wwb.createSheet(te.getText());

                XSSFSheet sheet = rwb.getSheetAt(0);     //creating a Sheet object to retrieve object  
                Iterator<Row> itr = sheet.iterator();    //iterating over excel file  
                int rownu = sheet.getLastRowNum();

                ArrayList<Row> rowarr = new ArrayList<Row>();
                ArrayList<Row> rrowarr = new ArrayList<Row>();

                Row r = easyship.createRow((short)0);
                rowarr.add(r);
                
                while(itr.hasNext()){
                    Row temp = itr.next();
                    rrowarr.add(temp);
                }

                for(int i = 1;i<=rownu;i++){
                    rowarr.add(i,easyship.createRow((short)i));
                }   

                for(Row x: rowarr){
                    System.out.println("row" + x);
                    for(int i = 0;i<rownu;i++){
                        System.out.println("x.getCell(i)" + x.getCell(i));
                    }
                }
                
                System.out.println("rownum " + easyship.getLastRowNum() + " " + rownu + " " + rowarr.size());
                rowarr.get(0).createCell(0).setCellValue((String)"Choose Courier by");  
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

                //int[] read = {0,1,5,18,26,28,29,30};
                //int[] write = {27,4,23,26,6,8,7,10};

                for(int j = 0;j<14;j++){
                    for(int i = 1;i<=rownu;i++){
                        String value;
                        int val;
                        switch(j){
                            case 0:
                                value = "USD";
                                System.out.println(value);
                                rowarr.get(i).createCell(27).setCellValue(value);
                                break;
                            case 1:
                                value = rrowarr.get(i).getCell(1).getStringCellValue();
                                System.out.println(value);
                                rowarr.get(i).createCell(4).setCellValue(value);
                                break;
                            case 2:
                                value = rrowarr.get(i).getCell(5).getStringCellValue();
                                System.out.println(value);
                                rowarr.get(i).createCell(23).setCellValue(value);  
                                break;
                            case 3:
                                val = (int)(rrowarr.get(i).getCell(18).getNumericCellValue());
                                System.out.println(val);
                                rowarr.get(i).createCell(26).setCellValue(val);
                                break;
                            case 4:
                                value = rrowarr.get(i).getCell(26).getStringCellValue() + " " + rrowarr.get(i).getCell(27).getStringCellValue();
                                value = value.substring(value.indexOf(":")+1);
                                value = value.substring(0,value.indexOf(" ")) + " " + value.substring(value.indexOf("Name:") + 5);
                                System.out.println(value);
                                rowarr.get(i).createCell(6).setCellValue(value);
                                break;
                            case 5:
                                value = rrowarr.get(i).getCell(28).getStringCellValue();
                                value = value.substring(value.indexOf(":") + 1);
                                System.out.println(value);
                                rowarr.get(i).createCell(8).setCellValue(value);
                                break;
                            case 6:
                                value = rrowarr.get(i).getCell(29).getStringCellValue();
                                value = "+" + value.substring(value.indexOf(":")+1);
                                System.out.println(value);
                                rowarr.get(i).createCell(7).setCellValue(value);
                                break;
                            case 7:
                                value = rrowarr.get(i).getCell(30).getStringCellValue();
                                String v1 = value.substring(value.indexOf(":")+1,value.indexOf(","));
                                String v2 = value.substring(value.indexOf(" ",value.indexOf(",")+2)+1);
                                String v3 = value.substring(value.indexOf(" ",value.indexOf(","))+1,value.indexOf(" ",value.indexOf(",")+2));
                                System.out.println(v1 + " " + v2 + " " + v3);
                                rowarr.get(i).createCell(10).setCellValue(v1);
                                rowarr.get(i).createCell(12).setCellValue(v2);
                                rowarr.get(i).createCell(13).setCellValue(v3);
                                break;
                            case 8:
                                value = "Receiver";
                                System.out.println(value);
                                rowarr.get(i).createCell(2).setCellValue(value);
                                break; 
                            case 9:
                                value = t.getText();
                                System.out.println(value);
                                rowarr.get(i).createCell(16).setCellValue(value);
                                break;
                            case 10:
                                value = t1.getText();
                                System.out.println(value);
                                rowarr.get(i).createCell(17).setCellValue(value);
                                break;
                            case 11:
                                value = t2.getText();
                                System.out.println(value);
                                rowarr.get(i).createCell(18).setCellValue(value);
                                break;
                            case 12:
                                value = t3.getText();
                                System.out.println(value);
                                rowarr.get(i).createCell(19).setCellValue(value);
                                break;
                            case 13:
                                value = t4.getText();
                                System.out.println(value);
                                rowarr.get(i).createCell(20).setCellValue(value);
                                break;
                    }
                }
            }
            FileOutputStream fileOut = new FileOutputStream(l.getText() + "\\" + te.getText() + ".xlsx");  
                wwb.write(fileOut);  
                //closing the Stream  
                fileOut.close();  
                //closing the workbook  
                wwb.close();  
                rwb.close();
                //prints the message on the console  
                System.out.println("Excel file has been generated successfully."); 
        }  
            catch(Exception t)  
            {  
                t.printStackTrace();  
            }  
                }
            });

            // x axis, y axis, width, height
            button.setBounds(150, 200, 150, 30);

            // adding button in JFrame
            panel.add(l2);
            panel.add(te);
            panel.add(l);
            panel.add(b2);

            panel.add(l3);
            panel.add(t);
            
            panel.add(l4);
            panel.add(t1);

            panel.add(l5);
            panel.add(t2);

            panel.add(l6);
            panel.add(t3);

            panel.add(l7);
            panel.add(t4);   
            
            panel.add(button);

            frame.add(panel);

            // 400 width and 500 height
            frame.setSize(500, 600);
    
            // using no layout managers
            //frame.setLayout(null);
    
            // making the frame visible
            frame.setVisible(true);
        }

        @Override
        public void actionPerformed(ActionEvent e) {
            String command = e.getActionCommand();

            if (command.equals("button1")) {
            }
        }
    }

    class GUI {
        public static void convert()   
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
                ArrayList<Row> rrowarr = new ArrayList<Row>();

                Row r = easyship.createRow((short)0);
                rowarr.add(r);
                
                while(itr.hasNext()){
                    Row temp = itr.next();
                    rrowarr.add(temp);
                }

                for(int i = 1;i<=rownu;i++){
                    rowarr.add(i,easyship.createRow((short)i));
                }   


                for(Row x: rowarr){
                    System.out.println("row" + x);
                    for(int i = 0;i<rownu;i++){
                        System.out.println("x.getCell(i)" + x.getCell(i));
                    }
                }
                
                System.out.println("rownum " + easyship.getLastRowNum() + " " + rownu + " " + rowarr.size());
                rowarr.get(0).createCell(0).setCellValue((String)"Choose Courier by");  
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

                //int[] read = {0,1,5,18,26,28,29,30};
                //int[] write = {27,4,23,26,6,8,7,10};

                for(int j = 0;j<8;j++){
                    for(int i = 1;i<=rownu;i++){
                        String value;
                        switch(j){
                            case 0:
                                value = "USD";
                                System.out.println(value);
                                rowarr.get(i).createCell(27).setCellValue(value);
                                break;
                            case 1:
                                value = rrowarr.get(i).getCell(1).getStringCellValue();
                                System.out.println(value);
                                rowarr.get(i).createCell(4).setCellValue(value);
                                break;
                            case 2:
                                value = rrowarr.get(i).getCell(5).getStringCellValue();
                                System.out.println(value);
                                rowarr.get(i).createCell(23).setCellValue(value);  
                                break;
                            case 3:
                                int val = (int)(rrowarr.get(i).getCell(18).getNumericCellValue());
                                System.out.println(val);
                                rowarr.get(i).createCell(26).setCellValue(val);
                                break;
                            case 4:
                                value = rrowarr.get(i).getCell(26).getStringCellValue() + " " + rrowarr.get(i).getCell(27).getStringCellValue();
                                value = value.substring(value.indexOf(":")+1);
                                value = value.substring(0,value.indexOf(" ")) + " " + value.substring(value.indexOf("Name:") + 5);
                                System.out.println(value);
                                rowarr.get(i).createCell(6).setCellValue(value);
                                break;
                            case 5:
                                value = rrowarr.get(i).getCell(28).getStringCellValue();
                                value = value.substring(value.indexOf(":") + 1);
                                System.out.println(value);
                                rowarr.get(i).createCell(8).setCellValue(value);
                                break;
                            case 6:
                                value = rrowarr.get(i).getCell(29).getStringCellValue();
                                value = "+" + value.substring(value.indexOf(":")+1);
                                System.out.println(value);
                                rowarr.get(i).createCell(7).setCellValue(value);
                                break;
                            case 7:
                                value = rrowarr.get(i).getCell(30).getStringCellValue();
                                String v1 = value.substring(value.indexOf(":")+1,value.indexOf(","));
                                String v2 = value.substring(value.indexOf(" ",value.indexOf(",")+2)+1);
                                String v3 = value.substring(value.indexOf(" ",value.indexOf(","))+1,value.indexOf(" ",value.indexOf(",")+2));
                                System.out.println(v1 + " " + v2 + " " + v3);
                                rowarr.get(i).createCell(10).setCellValue(v1);
                                rowarr.get(i).createCell(12).setCellValue(v2);
                                rowarr.get(i).createCell(13).setCellValue(v3);
                                break;
                        }
                    }
                }

                FileOutputStream fileOut = new FileOutputStream("C:\\JavaPrograms\\ExcelManager\\EasyShip.xlsx");  
                wwb.write(fileOut);  
                //closing the Stream  
                fileOut.close();  
                //closing the workbook  
                wwb.close();  
                rwb.close();
                //prints the message on the console  
                System.out.println("Excel file has been generated successfully."); 
            }  
            catch(Exception e)  
            {  
                e.printStackTrace();  
            }  
    }   
}

/*
            for(int i = 1;i<=rownu;i++){
                String value = rrowarr.get(i).getCell(5).getStringCellValue();
                System.out.println(value);
                rowarr.get(i).createCell(23).setCellValue(value);  
            }

            for(int i = 1;i<=rownu;i++){
                String value = rrowarr.get(i).getCell(1).getStringCellValue();
                System.out.println(value);
                rowarr.get(i).createCell(4).setCellValue(value);
            }

            for(int i = 1;i<=rownu;i++){
                String value = rrowarr.get(i).getCell(29).getStringCellValue();
                value = "+" + value.substring(value.indexOf(":")+1);
                System.out.println(value);
                rowarr.get(i).createCell(7).setCellValue(value);
            }

            for(int i = 1;i<=rownu;i++){
                String value = rrowarr.get(i).getCell(28).getStringCellValue();
                value = value.substring(value.indexOf(":") + 1);
                System.out.println(value);
                rowarr.get(i).createCell(8).setCellValue(value);
            }

            for(int i = 1;i<=rownu;i++){
                String value = rrowarr.get(i).getCell(26).getStringCellValue() + " " + rrowarr.get(i).getCell(27).getStringCellValue();
                value = value.substring(value.indexOf(":")+1);
                value = value.substring(0,value.indexOf(" ")) + " " + value.substring(value.indexOf("Name:") + 5);
                System.out.println(value);
                rowarr.get(i).createCell(6).setCellValue(value);
            }

            for(int i = 1;i<=rownu;i++){
                int value = (int)(rrowarr.get(i).getCell(18).getNumericCellValue());
                System.out.println(value);
                rowarr.get(i).createCell(26).setCellValue(value);
            }

            for(int i = 1;i<=rownu;i++){
                String value = "USD";
                System.out.println(value);
                rowarr.get(i).createCell(27).setCellValue(value);
            }

            for(int i = 1;i<=rownu;i++){
                String value = rrowarr.get(i).getCell(30).getStringCellValue();
                String v1 = value.substring(value.indexOf(":")+1,value.indexOf(","));
                String v2 = value.substring(value.indexOf(" ",value.indexOf(",")+2)+1);
                String v3 = value.substring(value.indexOf(" ",value.indexOf(","))+1,value.indexOf(" ",value.indexOf(",")+2));
                System.out.println(v1 + " " + v2 + " " + v3);
                rowarr.get(i).createCell(10).setCellValue(v1);
                rowarr.get(i).createCell(12).setCellValue(v2);
                rowarr.get(i).createCell(13).setCellValue(v3);
            }
            */
/*
            for(int i = 1;i<=rownu;i++){
                String value = rrowarr.get(i).getCell(30).getStringCellValue();
                value = value.substring(value.indexOf(" ",value.indexOf(",")+2)+1);
                System.out.println(value);
                rowarr.get(i).createCell(12).setCellValue(value);
            }

            for(int i = 1;i<=rownu;i++){
                String value = rrowarr.get(i).getCell(30).getStringCellValue();
                value = value.substring(value.indexOf(" ",value.indexOf(","))+1,value.indexOf(" ",value.indexOf(",")+2));
                System.out.println(value);
                rowarr.get(i).createCell(13).setCellValue(value);
            }
*/
            /*
            while(itr.hasNext()){
                Row r1 = itr.next();
                Iterator<Cell> cells = r1.cellIterator();
                while(cells.hasNext()){
                    Cell temp = cells.next();

                    

                    switch(temp.getCellType()){
                        case Cell.CELL_TYPE_STRING:
                            if(temp.getStringCellValue() == "Product"){
                                int r2 = temp.getRowIndex();
                                int r3 = temp.getColumnIndex();
                                System.out.println("r2" + r2 + " r3 " + r3);

                                rowarr.get(r2).createCell(r3).setCellValue(r1.getCell(r3).getNumericCellValue());
                            }
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            break;
                    }
                }
            }
*/

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