import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;

public class ExcelUtil {

    static HSSFWorkbook wb = new HSSFWorkbook();
    public static void createFile() throws IOException {
        FileOutputStream out = new FileOutputStream("excel\\ExcelTMP.xls");
        wb.write(out);
        out.close();
    }
    public static Sheet rowClear(Sheet sheet){
        Iterator<Row> rows = sheet.rowIterator();
        Iterator<Cell> cells;
        Row currentRow;
        Cell currentCell;
        CellType cellType;
        boolean flag;
        while(rows.hasNext()){
            flag = true;
            currentRow = rows.next();
            cells = currentRow.cellIterator();
            while(cells.hasNext()){
                currentCell = cells.next();
                if (currentCell.getCellType() != CellType.BLANK){
                    cellType = currentCell.getCellType();
                    if (cellType == CellType.STRING && !currentCell.getStringCellValue().equals("")){
                        flag = false;
                    }
                    else if (cellType == CellType.NUMERIC && currentCell.getNumericCellValue() != 0){
                        flag = false;
                    }
                }
            }
            if (flag){
                rows.remove();
                sheet.removeRow(currentRow);
            }
        }
        return sheet;
    }
    public static Sheet cellClear(Sheet sheet){
        Cell currentCell;
        CellType type;
        Row currentRow;
        boolean flag = true;
        Iterator<Row> rows = sheet.rowIterator();
        Iterator<Cell> cells;
        while(rows.hasNext()) {
            currentRow = rows.next();
            cells = currentRow.cellIterator();
            while (cells.hasNext()) {
                currentCell = cells.next();
                type = currentCell.getCellType();
                if (type == CellType.STRING && !currentCell.getStringCellValue().equals("")) {
                    flag = false;
                } else if (type == CellType.NUMERIC && currentCell.getNumericCellValue() != 0) {
                    flag = false;
                }
                if (flag) {
                    cells.remove();
                    currentRow.removeCell(currentCell);
                }
            }
        }
        return sheet;
    }
    public static void removeBlankCol(Sheet sheet, int sheetNum) throws NullPointerException{
        int firstRow = sheet.getFirstRowNum();
        int lastRow = sheet.getLastRowNum();
        int rowNum = lastRow - firstRow + 1;
        int colNum = sheet.getRow(firstRow).getPhysicalNumberOfCells();
        int firstCol = sheet.getRow(firstRow).getFirstCellNum();
        int lastCol = firstCol + colNum - 1;
        int colIndex = firstCol;
        int rowIndex;
        HSSFSheet newSheet = wb.createSheet("Sheet " + (sheetNum + 1));
        newSheet.setDefaultRowHeight((short)450);
        for (int i = 0; i < rowNum; i++){
            newSheet.createRow(i);
        }
        int newSheetR;
        int newSheetC = 0;
        int maxWidth = 10;
        int width;
        Cell currentCell;
        boolean flag;
        while (colIndex <= lastCol) {
            flag = false;
            rowIndex = firstRow;
            newSheetR = 0;
            while (rowIndex <= lastRow) {
                try{
                    sheet.getRow(rowIndex).getCell(colIndex);
                }
                catch (NullPointerException e){
                    sheet.createRow(rowIndex).getCell(colIndex);
                }
                currentCell = sheet.getRow(rowIndex).getCell(colIndex);
                if (currentCell != null && rowIndex != firstRow) {
                    if (currentCell.getCellType() == CellType.STRING && !currentCell.getStringCellValue().equals("")) {
                        flag = true;
                    }
                    else if (currentCell.getCellType() == CellType.NUMERIC && currentCell.getNumericCellValue() != 0) {
                        flag = true;
                    }
                    else if (currentCell.getCellType() == CellType.FORMULA) {
                        flag = true;
                    }
                }
                rowIndex++;
            }
            if (flag) {
                for (int i = firstRow; i <= lastRow; i++) {
                    String val = "";
                    try {
                        Cell c = sheet.getRow(i).getCell(colIndex);
                        if (c.getCellType() == CellType.STRING) {
                            val = c.getStringCellValue();
                        } else if (c.getCellType() == CellType.NUMERIC) {
                            val = String.valueOf(c.getNumericCellValue());
                        } else if (c.getCellType() == CellType.FORMULA) {
                            try {
                                val = String.valueOf(c.getNumericCellValue());
                            } catch (IllegalStateException e) {
                                val = String.valueOf(c.getRichStringCellValue());
                            }
                        }
                        width = val.getBytes().length;
                        maxWidth = (maxWidth > width) ? maxWidth : width;
                        newSheet.getRow(newSheetR).
                                getCell(newSheetC, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).
                                setCellValue(val);
                    }catch (NullPointerException e)
                    {
                        continue;
                    }
                    newSheetR++;
                }
                newSheet.setColumnWidth(newSheetC, maxWidth * 190);
                newSheetC++;
            }
            colIndex++;
        }
        wb.setSheetName(sheetNum, newSheet.getRow(1).getCell(0).getStringCellValue());
    }
    public static HashMap<String, String> getInfo(Sheet sheet){
        HashMap<String, String> info = new HashMap<>();
        sheet = ExcelUtil.rowClear(sheet);
        sheet = ExcelUtil.cellClear(sheet);
        int firstRow = sheet.getFirstRowNum();
        int lastRow = sheet.getLastRowNum();
        int firstCell = sheet.getRow(firstRow).getFirstCellNum();
        int lastCell = sheet.getRow(firstRow).getLastCellNum();
        int nameInd = 0;
        int emailInd = 0;
        for (int i = firstCell; i < lastCell; i++){
            if (sheet.getRow(firstRow).getCell(i).getStringCellValue().equals("供应商")){
                nameInd = i;
            }
            else if (sheet.getRow(firstRow).getCell(i).getStringCellValue().equals("邮箱地址")){
                emailInd = i;
            }
        }
        for (int j = firstRow; j < lastRow; j++){
            info.put(sheet.getRow(j).getCell(nameInd).getStringCellValue(), sheet.getRow(j).getCell(emailInd).getStringCellValue());
        }
        return info;
    }
}
