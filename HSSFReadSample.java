import java.io.FileInputStream;
import java.text.DateFormat;
import java.util.Date;
import org.apache.poi.hssf.model.*;
import org.apache.poi.hssf.usermodel.*;
 
public class HSSFReadSample {
 
    public static void main(String[] arg) throws Exception {
 
        String filename = arg[0];
        FileInputStream is = new FileInputStream(filename);
 
        // ワークブック・ワークシートの取得
        HSSFWorkbook book = new HSSFWorkbook(is);  // ← (1)
        HSSFSheet sheet = book.getSheetAt(0);      // ← (2)
        System.out.println("number of sheets : " + book.getNumberOfSheets());
        System.out.println("sheet name : " + book.getSheetName(0));
 
        int first = sheet.getFirstRowNum();  // ← (3)
        int last  = sheet.getLastRowNum();   // ← (3)
        System.out.println("first row : " + first);
        System.out.println("last row  : " + last);
 
        HSSFRow row = null;
        HSSFCell cell = null;
        DateFormat dateform = DateFormat.getDateInstance();
        for (int i=4;i<=6;i++) {
 
            // 行・セル・セル値の取得
            row = sheet.getRow(i);          // ← (4)
            cell = row.getCell((short)1);   // ← (5)
            String product = cell.getStringCellValue();  // ← (6)
 
            // セル値を日付として取得
            cell = row.getCell((short)2);
            Date incoming = cell.getDateCellValue();     // ← (7)
            String date = dateform.format(incoming);
 
            cell = row.getCell((short)3);
            double price = cell.getNumericCellValue();   // ← (8)
 
            cell = row.getCell((short)4);
            int num = (int)cell.getNumericCellValue();
 
            cell = row.getCell((short)5);
            double sum = cell.getNumericCellValue();
 
            System.out.print(i + " : " );
            System.out.println(product + "," + date + "," + price + "," + sum);
        }
 
        // セルの型・表示形式の取得
        row = sheet.getRow(7);
        cell = row.getCell((short)5);
        System.out.println("cell_type    : " + cell.getCellType());
        System.out.println("cell_value   : " + cell.getNumericCellValue());
        System.out.println("cell_formula : " + cell.getCellFormula()); // ← (9)
 
        is.close();
    }
}