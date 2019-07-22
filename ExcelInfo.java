import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;

public class ExcelInfo {
    public static File file;
    public static void setFile() {
        file = new File("excel\\Excel1.xlsx");
    }
    public static ArrayList<Sheet> getSheetList() throws Exception {
        Workbook wb;
        Iterator<Sheet> sheets;
        ArrayList<Sheet> sheetList = new ArrayList<>();
        try{
            wb = new XSSFWorkbook(new FileInputStream(file));
            sheets = wb.sheetIterator();
            if (sheets == null){
                throw new Exception("excel中不含有sheet工作表");
            }
            while (sheets.hasNext()) {
                sheetList.add(sheets.next());
            }
        } catch (IOException e) {
            JOptionPane.showMessageDialog(null, "未找到文件,程序即将退出", "Cannot find file", JOptionPane.ERROR_MESSAGE);
            Thread.sleep(1000);
            System.exit(0);
        }
        return sheetList;
    }
}
