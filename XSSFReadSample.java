import java.io.FileInputStream;
import java.text.DateFormat;
import java.util.Date;
import org.apache.poi.xssf.model.*;
import org.apache.poi.xssf.usermodel.*;
 
public class XSSFReadSample {
 
    public static void main(String[] arg) throws Exception {
 
        String filename = arg[0];
        FileInputStream is = new FileInputStream(filename);
 
        // ���[�N�u�b�N�E���[�N�V�[�g�̎擾
        XSSFWorkbook book = new XSSFWorkbook(is);  // �� (1)
        XSSFSheet sheet = book.getSheetAt(0);      // �� (2)
        System.out.println("number of sheets : " + book.getNumberOfSheets());
        System.out.println("sheet name : " + book.getSheetName(0));
 
        int first = sheet.getFirstRowNum();  // �� (3)
        int last  = sheet.getLastRowNum();   // �� (3)
        System.out.println("first row : " + first);
        System.out.println("last row  : " + last);
 
        XSSFRow row = null;
        XSSFCell cell = null;
        DateFormat dateform = DateFormat.getDateInstance();
        for (int i=4;i<=6;i++) {
 
            // �s�E�Z���E�Z���l�̎擾
            row = sheet.getRow(i);          // �� (4)
            cell = row.getCell((short)1);   // �� (5)
            String product = cell.getStringCellValue();  // �� (6)
 
            // �Z���l����t�Ƃ��Ď擾
            cell = row.getCell((short)2);
            Date incoming = cell.getDateCellValue();     // �� (7)
            String date = dateform.format(incoming);
 
            cell = row.getCell((short)3);
            double price = cell.getNumericCellValue();   // �� (8)
 
            cell = row.getCell((short)4);
            int num = (int)cell.getNumericCellValue();
 
            cell = row.getCell((short)5);
            double sum = cell.getNumericCellValue();
 
            System.out.print(i + " : " );
            System.out.println(product + "," + date + "," + price + "," + sum);
        }
 
        // �Z���̌^�E�\���`���̎擾
        row = sheet.getRow(7);
        cell = row.getCell((short)5);
        System.out.println("cell_type    : " + cell.getCellType());
        System.out.println("cell_value   : " + cell.getNumericCellValue());
        System.out.println("cell_formula : " + cell.getCellFormula()); // �� (9)
 
        is.close();
    }
}