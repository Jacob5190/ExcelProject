import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import java.io.File;
import java.util.ArrayList;

public class Convert {
    public static void ConvertToImage (ArrayList<String> name, int num){
        String dataDir = getDataDir(Convert.class);
        Workbook book;
        try {
            book = new Workbook("excel\\ExcelTMP.xls");
            // Get the first worksheet
            for (int i = 0; i < num - 1; i++) {


                Worksheet sheet = book.getWorksheets().get(i);
                sheet.getPageSetup().setLeftMargin(-20);
                sheet.getPageSetup().setRightMargin(0);
                sheet.getPageSetup().setBottomMargin(0);
                sheet.getPageSetup().setTopMargin(0);

                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
                imgOptions.setImageFormat(ImageFormat.getJpeg());
                imgOptions.setCellAutoFit(true);
                imgOptions.setOnePagePerSheet(true);
                //imgOptions.setDesiredSize(1000,800);
                SheetRender render = new SheetRender(sheet, imgOptions);


                render.toImage(0, ("images\\"+name.get(i)+".jpg"));
            }
        } catch(Exception e){
            e.printStackTrace();
        }
    }

    public static String getDataDir(java.lang.Class c) {
        File dir = new File(System.getProperty("user.dir"));
        dir = new File(dir, "src");
        dir = new File(dir, "main");
        dir = new File(dir, "resources");

        for (String s : c.getName().split("\\.")) {
            dir = new File(dir, s);
        }

        if (!dir.exists()) {
            dir.mkdirs();
        }

        return dir.toString() + File.separator;
    }
}