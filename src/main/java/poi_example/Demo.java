package poi_example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Demo {

    public static void writeOutput(String label1, String label2, String label3, double avg1, double avg2, double avg3)
            throws Exception {
        
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Averages");

        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue(label1);
        row.createCell(1).setCellValue(label2);
        row.createCell(2).setCellValue(label3);

        row = sheet.createRow(1);
        row.createCell(0).setCellValue(avg1);
        row.createCell(1).setCellValue(avg2);
        row.createCell(2).setCellValue(avg3);

        // Write the workbook in file system
        FileOutputStream out = new FileOutputStream(new File("out.xlsx"));
        workbook.write(out);
        out.close();
    }

    public static void main(String[] args) throws Exception {
        FileInputStream file = new FileInputStream(new File("input.xlsx"));
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheet("data");

        // Iterate through each rows one by one
        Iterator<Row> rowIterator = sheet.iterator();

        // header row
        Row row = rowIterator.next();
        String column1Label = row.getCell(0).getStringCellValue();
        String column2Label = row.getCell(1).getStringCellValue();
        String column3Label = row.getCell(2).getStringCellValue();

        double sum1 = 0;
        double sum2 = 0;
        double sum3 = 0;
        int rows = 0;

        // read until there are no more rows
        while (rowIterator.hasNext()) {
            row = rowIterator.next();
            rows++;

            sum1 += row.getCell(0).getNumericCellValue();
            sum2 += row.getCell(1).getNumericCellValue();
            sum3 += row.getCell(2).getNumericCellValue();

        }

        writeOutput(column1Label, column2Label, column3Label, sum1/rows, sum2/rows, sum3/rows);
    }
}
