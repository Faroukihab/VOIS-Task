package pages;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.Period;
import java.time.ZoneId;
import java.util.Date;

public class UpdateVOISYears {

    public static void main(String[] args) throws IOException, ParseException {
        String excelFilePath = "C:\\Users\\fihab\\eclipse-workspace\\Voisautomationscript\\src\\main\\java\\data\\TaskData.xlsx"; 
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);

        SimpleDateFormat dateFormat = new SimpleDateFormat("EEEE, MMMM dd, yyyy");

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            Cell joinDateCell = row.getCell(2); // "Join Date" is the 3rd column (index 2)
            String joinDateString = joinDateCell.getStringCellValue();
            Date joinDate = dateFormat.parse(joinDateString);

            LocalDate joinLocalDate = joinDate.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
            LocalDate currentDate = LocalDate.now();

            int yearsInVOIS = Period.between(joinLocalDate, currentDate).getYears();

            Cell yearsCell = row.createCell(3); // "Today, How Many years in VOIS" is the 4th column (index 3)
            yearsCell.setCellValue(yearsInVOIS);
        }

        inputStream.close();

        FileOutputStream outputStream = new FileOutputStream(excelFilePath);
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();

        System.out.println("Excel file updated successfully!");
    }
}
