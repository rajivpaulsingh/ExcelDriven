import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

public class DataDriven {


    //Identify the 'Testcases' column by scanning the entire first wo
    //Once column is identified then scan the entire 'Testcase' column to identify 'Purchase' testcase row
    //After that, grab the 'Purchase' test case row data and feed into test

    public static void main (String args[]) throws IOException {

        /**
        - Strategy to access Excel
        - Create object to XSSFWorkbook class
        - Get access to Sheet
        - Get access to all rows of Sheet
        - Get access to specific row from all rows
        - Get access to all cells of row
        - Access the data from excel into arrays
         */


        ArrayList<String> a = new ArrayList<String>();
        FileInputStream fis = new FileInputStream("/Users/singh2_rajiv/Selenium/DemoData.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fis);

        int sheets = workbook.getNumberOfSheets();
        for(int i = 0; i < sheets; i++) {

            if(workbook.getSheetName(i).equalsIgnoreCase("testdata")) {

                XSSFSheet sheet = workbook.getSheetAt(i);

                Iterator<Row> rows = sheet.iterator();
                Row firstrow = rows.next();

                Iterator<Cell> cells = firstrow.cellIterator();
                cells.next();

                int k = 0;
                int column = 0;
                while(cells.hasNext()) {
                    Cell value = cells.next();

                    if(value.getStringCellValue().equalsIgnoreCase("TestCases")) {
                        //Desired column
                        column = k;
                    }

                    k++;
                }
                System.out.println(column);

                while(rows.hasNext()) {

                    Row r = rows.next();
                    if(r.getCell(column).getStringCellValue().equalsIgnoreCase("Purchase")) {
                        //Desired row
                        //Grab all the data from the row
                        Iterator<Cell> cv = r.cellIterator();

                        while(cv.hasNext()) {

                            Cell c = cv.next();
                            if(c.getCellType()== CellType.STRING) { //String value
                                a.add(c.getStringCellValue());
                            }
                            else { //Integer value
                                a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
//                                a.add(c.getNumericCellValue());
                            }
//                            System.out.println(cv.next().getStringCellValue());
//                            a.add(cv.next().getStringCellValue());
                        }
                        System.out.println(a);
                    }
                }
            }
        }

    }
}
