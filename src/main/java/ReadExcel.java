import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;

import static java.lang.Integer.parseInt;

public class ReadExcel {

    public static ArrayList<HashMap<String, Object>> readExcelFromPath(String filePath){

        //TO DO -- how to check office version and use either XSSF and HSSF
        //hashmap to hold key, value pair
        HashMap<Integer, String> excelHeaderMap = new HashMap<>();
        //to hold list of rows as a list of hashmaps
        ArrayList<HashMap<String,Object>> excelHeaderValueList = new ArrayList<>();

        FileInputStream fileInputStream= null;
        XSSFWorkbook wb = null;

        FormulaEvaluator formulaEvaluator;

        try {
            fileInputStream = new FileInputStream(new File(filePath));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            wb=new XSSFWorkbook(fileInputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        XSSFSheet sheet=wb.getSheetAt(0);
        formulaEvaluator= wb.getCreationHelper().createFormulaEvaluator();

        //iterate each cells
        for(Row row: sheet){

            int rowNum = row.getRowNum();
            Integer cellIndex = 0;

            HashMap<String, Object> excelRowValueMap = new HashMap<>();

            for(Cell cell: row){

                Object cellValue = null;
                int cellType = cell.getCellType();

                if(cellType == Cell.CELL_TYPE_STRING){
                    cellValue = cell.getStringCellValue();
                }else if(cellType == Cell.CELL_TYPE_BOOLEAN){
                    cellValue = cell.getBooleanCellValue();
                }else if(cellType == Cell.CELL_TYPE_NUMERIC){
                    DataFormatter dataFormatter = new DataFormatter();
                    String formattedCellStr = dataFormatter.formatCellValue(cell);

                    try {
                        cellValue = Long.parseLong(formattedCellStr);
                    }catch (NumberFormatException e3){

                    try {
                        cellValue = parseInt(formattedCellStr);
                    }
                    catch(NumberFormatException e){
                        try {
                            cellValue =  Double.parseDouble(formattedCellStr.trim());
                        }catch (NumberFormatException e2){
                            cellValue = formattedCellStr;
                        }
                    }
                    }
                }else{
                    cellValue = cell.getStringCellValue();
                }
                if(rowNum == 0){
                    excelHeaderMap.put(cellIndex,(String)cellValue);
                }else{
                    excelRowValueMap.put(excelHeaderMap.get(cellIndex), cellValue);
                }
                cellIndex+=1;
            }
            if(rowNum != 0) excelHeaderValueList.add(excelRowValueMap);
        }
        return excelHeaderValueList;
    }
}
