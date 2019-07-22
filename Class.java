import org.apache.poi.ss.usermodel.Sheet;

import javax.swing.*;
import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;

public class Class {
    public static boolean function() throws Exception {
        HashMap<String, String> info = new HashMap<>();
        ArrayList<String> names = new ArrayList<>();
        createFile();
        ArrayList<Sheet> sheet;
        try {
            ExcelInfo.setFile();
            sheet = ExcelInfo.getSheetList();
            if (sheet == null){
                return false;
            }
            int sheetSize = sheet.size();
            Sheet tmp;
            for (int i = 0; i < sheetSize; i++) {
                if (!sheet.get(i).getSheetName().equals("联系人-对应邮箱")) {
                    Gui.textOut.append("==================\nStarting Sheet " + (i + 1) + "...\n");
                    tmp = ExcelUtil.rowClear(sheet.get(i));
                    tmp = ExcelUtil.cellClear(tmp);
                    try {
                        ExcelUtil.removeBlankCol(tmp, i);
                    } catch (NullPointerException e) {
                        e.printStackTrace();
                    }
                    names.add(tmp.getSheetName().split("-")[0]);
                }
                else{
                    info = ExcelUtil.getInfo(sheet.get(i));
                }
            }
            ExcelUtil.createFile();
            Convert.ConvertToImage(names, sheetSize);

        } catch (Exception e) {
            e.printStackTrace();
        }
        File dir = new File("images");
        File[] files = dir.listFiles();
        String fileName;
        String email;
        for (File f: files){
            fileName = f.getName().replaceAll(".jpg", "");
            email = info.get(fileName);
            if (email != null) {
                sendMail.send(f.getAbsolutePath(), email, fileName);
            }
            else{
                Gui.textOut.append("Cannot find Key!\n");
                Gui.textOut.append("Key: " + fileName + "\n");
            }
        }
        return true;
    }
    public static void createFile() throws InterruptedException {
        File dir1 = new File("excel");
        File dir2 = new File("images");
        if (!dir1.exists()) {
            dir1.mkdirs();
            JOptionPane.showMessageDialog(null, "未找到文件,程序即将退出", "Cannot find file", JOptionPane.ERROR_MESSAGE);
            Thread.sleep(1000);
            System.exit(0);
        }
        if (!dir2.exists()) dir2.mkdirs();
    }
}
