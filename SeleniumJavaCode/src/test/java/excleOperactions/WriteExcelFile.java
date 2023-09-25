package excleOperactions;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

//Workbook --> sheet--> Rows-> cells
public class WriteExcelFile {
    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Emp Info");

        Object empdata[][] = {{"EmpId", "Name", "Job"},
                {"101", "santosh", "Engg"},
                {"102", "suyog", "QA"},
                {"103", "Deepak", "Dev"}
        };
        int rowCount = 0;
        for (Object emp[] : empdata) {
            XSSFRow row = sheet.createRow(rowCount++);
            int columnCount = 0;
            for (Object value : emp) {
                XSSFCell cell = row.createCell(columnCount++);
                if (value instanceof String)
                    cell.setCellValue((String) value);
                if (value instanceof Integer)
                    cell.setCellValue((Integer) value);
                if (value instanceof Boolean)
                    cell.setCellValue((Boolean) value);
            }
        }
        String filepath = ".\\DataFile\\employee.xlsx";
        FileOutputStream outputStream = new FileOutputStream(filepath);
        workbook.write(outputStream);
        outputStream.close();
        System.out.println("employee.xlsx file created");


    }
}
