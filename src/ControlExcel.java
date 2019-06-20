import org.apache.poi.hssf.usermodel.HSSFShapeGroup;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class ControlExcel {
    public static void main(String[] args) {
       createXLS("D:/", "test");
       createXLSX("D:/", "test");
    }

    // 在指定的路径下创建指定名称的 .xls  的 Excel  文件
    public static void createXLS(String path, String name) {
        try {
            // 创建工作簿
            HSSFWorkbook workbook = new HSSFWorkbook();
            // 创建文件流
            FileOutputStream outputStream = new FileOutputStream(path + name + ".xls");
            // 在这里可以创建文件，但是文件打开的时候出错，因为此时仅仅创建了 Excel 的工作簿，但是没有创建 工作表
            // 也就是 Excel 文件结构不完整
            // 创建工作表
            HSSFSheet sheet = workbook.createSheet("工作表1");
            workbook.write(outputStream);
            outputStream.close();
        } catch (FileNotFoundException ex) {
            ex.printStackTrace();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }

    // 在指定路径下创建指定名称的 .xlsx 文件
    public static void createXLSX(String path, String name) {
        try {
            XSSFWorkbook workbook = new XSSFWorkbook();
            FileOutputStream outputStream = new FileOutputStream(path + name + ".xlsx");
            XSSFSheet sheet1 = workbook.createSheet("工作表1");
            workbook.write(outputStream);
            outputStream.close();
        } catch (FileNotFoundException ex) {
            ex.printStackTrace();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }
}
