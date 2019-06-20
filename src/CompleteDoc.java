import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

// 创建一个完整的 Excel 文件，并在其中输入信息

public class CompleteDoc {
    public static void main(String[] args) {
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss:SSS");
        try {
            // 创建工作簿
            HSSFWorkbook workbook = new HSSFWorkbook();
            // 创建文件流
            FileOutputStream outputStream = new FileOutputStream("D:/test.xls");
            // 创建工作表
            HSSFSheet sheet = workbook.createSheet("工作表1");
            // 参数 0 行表示第一行
            HSSFRow row = sheet.createRow(0);
            row.createCell(0).setCellValue("姓名");
            row.createCell(1).setCellValue("时间");
            // 参数 1 表示第二行
            row = sheet.createRow(1);
            row.createCell(0).setCellValue("碾作尘");
            Date date = new Date();
            row.createCell(1).setCellValue(dateFormat.format(date));

            //将信息写入输出流，将输出流中信息写入文件
            workbook.write(outputStream);
            outputStream.close();
        } catch (FileNotFoundException ex) {
            ex.printStackTrace();
        } catch (IOException ex) {
            ex.printStackTrace();
        }
    }
}
