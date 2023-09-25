package excleOperactions;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ReadExcelFile {
    public static void main(String[] args) throws IOException {
        // first take location of file
        String excelFilePath = "DataFile/simple.xlsx";
        // open into read mode
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        // Then we take work book from the file
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook.getSheet("Sheet1");

        int rows = sheet.getLastRowNum();

        // Iterate through rows
        for (int rowData = 0; rowData <= rows; rowData++) {
            XSSFRow row = sheet.getRow(rowData);

            // Get the last cell number in the current row
            int cols = row.getLastCellNum();

            // Iterate through cells in the current row
            for (int colData = 0; colData < cols; colData++) {
                XSSFCell cell = row.getCell(colData);

                // Check the cell type
                switch (cell.getCellType()) {
                    case STRING:
                        System.out.println(cell.getStringCellValue());
                        break;
                    case NUMERIC:
                        System.out.println(cell.getNumericCellValue());
                        break;
                    case BOOLEAN:
                        System.out.println(cell.getBooleanCellValue());
                        break;
                }
            }
            System.out.println();
        }
    }
}

//Program for read the data from excel file and print it
//"C:\Program Files\Java\jdk-17\bin\java.exe" "-javaagent:C:\Program Files\JetBrains\IntelliJ IDEA Community Edition 2023.2\lib\idea_rt.jar=52006:C:\Program Files\JetBrains\IntelliJ IDEA Community Edition 2023.2\bin" -Dfile.encoding=UTF-8 -classpath C:\Users\Suyog.Waghmare\Desktop\Projects\SeleniumJavaProject\SeleniumJavaCode\target\test-classes;C:\Users\Suyog.Waghmare\Desktop\Projects\SeleniumJavaProject\SeleniumJavaCode\target\classes;C:\Users\Suyog.Waghmare\.m2\repository\org\apache\poi\poi\5.2.3\poi-5.2.3.jar;C:\Users\Suyog.Waghmare\.m2\repository\commons-codec\commons-codec\1.15\commons-codec-1.15.jar;C:\Users\Suyog.Waghmare\.m2\repository\org\apache\commons\commons-collections4\4.4\commons-collections4-4.4.jar;C:\Users\Suyog.Waghmare\.m2\repository\org\apache\commons\commons-math3\3.6.1\commons-math3-3.6.1.jar;C:\Users\Suyog.Waghmare\.m2\repository\commons-io\commons-io\2.11.0\commons-io-2.11.0.jar;C:\Users\Suyog.Waghmare\.m2\repository\com\zaxxer\SparseBitSet\1.2\SparseBitSet-1.2.jar;C:\Users\Suyog.Waghmare\.m2\repository\org\apache\logging\log4j\log4j-api\2.18.0\log4j-api-2.18.0.jar;C:\Users\Suyog.Waghmare\.m2\repository\org\apache\poi\poi-ooxml\5.2.3\poi-ooxml-5.2.3.jar;C:\Users\Suyog.Waghmare\.m2\repository\org\apache\poi\poi-ooxml-lite\5.2.3\poi-ooxml-lite-5.2.3.jar;C:\Users\Suyog.Waghmare\.m2\repository\org\apache\xmlbeans\xmlbeans\5.1.1\xmlbeans-5.1.1.jar;C:\Users\Suyog.Waghmare\.m2\repository\org\apache\commons\commons-compress\1.21\commons-compress-1.21.jar;C:\Users\Suyog.Waghmare\.m2\repository\com\github\virtuald\curvesapi\1.07\curvesapi-1.07.jar excleOperactions.ReadExcelFile
//ERROR StatusLogger Log4j2 could not find a logging implementation. Please add log4j-core to the classpath. Using SimpleLogger to log to the console...
//name
//address
//number
//
//suyog
//pune
//1.23456789E9
//
//malinath
//baner pune
//1.233214568E9
//
//anushree
//kothrud
//8.528520123E9
//
//deepak
//dhayri
//9.633699632E9
//
//
//Process finished with exit code 0