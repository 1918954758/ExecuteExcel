package Excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @ClassName ParseExcel
 * @Discription 解析Excel
 * @Author 子辰
 * @Date 2020/9/16 20:12
 */
public class ParseExcel {
    public static void main(String[] args) {
        Workbook wb;
        Sheet sheet;
        Row row;
        List<Map<String,String>> list = null;
        String cellData;
        String filePath = "D:/workspace/testExcel/test.xls";
        String filePath1 = "D:/workspace/testExcel/test.xlsx";
        String columns[] = {"A","B","C","D","E","F","G","H","I","J","K","L","M","N","O"};
        wb = readExcel(filePath);
        if(wb != null){
            //用来存放表中数据// new ArrayList<Map<String,String>>
            list = new ArrayList<Map<String,String>>();
            //获取第一个sheet
            sheet = wb.getSheetAt(0);
            //获取最大行数
            int rownum = sheet.getPhysicalNumberOfRows();
            //获取第一行
            row = sheet.getRow(0);
            //获取最大列数
            int colnum = row.getPhysicalNumberOfCells();
            for (int i = 1; i<rownum; i++) {
                Map<String,String> map = new LinkedHashMap<>();
                row = sheet.getRow(i);
                if(row !=null){
                    for (int j=0;j<colnum;j++){
                        cellData = (String) getCellFormatValue(row.getCell(j));
                        map.put(columns[j], cellData);
                    }
                }else{
                    break;
                }
                list.add(map);
            }
        }
        //遍历解析出来的list
        /*for (Map<String,String> map : list) {
            for (Entry<String,String> entry : map.entrySet()) {
                System.out.print(entry.getKey()+":"+entry.getValue()+",");
            }
            System.out.println();
        }*/
        for (Map<String,String> map : list) {
//            for (Entry<String,String> entry : map.entrySet()) {
//                System.out.print(entry.getKey()+":"+entry.getValue()+",");
//            }
            System.out.println(map);
        }
    }

    /**
     * 读取excel
     * @param filePath
     * @return Workbook
     */
    public static Workbook readExcel(String filePath){
        Workbook wb = null;
        if(filePath==null){
            return null;
        }
        String fileName = filePath.substring(filePath.indexOf(".") + 1);
        InputStream is;
        try {
            is = new FileInputStream(filePath);
            if("xls".equals(fileName)){
                wb = new HSSFWorkbook(is);
                return wb;
            }else if("xlsx".equals(fileName)){
                wb = new XSSFWorkbook(is);
                return wb ;
            }else{
                return null;
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return wb;
    }

    /**
     * 解析 Cell
     * @param cell
     * @return Object
     */
    public static Object getCellFormatValue(Cell cell){
        Object cellValue = null;
        NumberFormat nf = NumberFormat.getInstance();
        if(cell!=null){
            //判断cell类型
            switch(cell.getCellType()){
                case Cell.CELL_TYPE_NUMERIC:{
                    String s = nf.format(cell.getNumericCellValue());
                    if(s.indexOf(",") >= 0){
                        cellValue = s.replace(",","");
                    }else{
                        cellValue = s;
                    }
                    //cellValue = String.valueOf(cell.getNumericCellValue());
                    break;
                }
                case Cell.CELL_TYPE_FORMULA:{
                    //判断cell是否为日期格式
                    if(DateUtil.isCellDateFormatted(cell)){
                        //转换为日期格式YYYY-mm-dd
                        cellValue = cell.getDateCellValue();
                    }else{
                        //数字
                        cellValue = nf.format(cell.getNumericCellValue());
                        //cellValue = String.valueOf(cell.getNumericCellValue());
                        System.out.println(cellValue);
                    }
                    break;
                }
                case Cell.CELL_TYPE_STRING:{
                    cellValue = cell.getRichStringCellValue().getString();
                    break;
                }
                default:
                    cellValue = "";
            }
        }else{
            cellValue = "";
        }
        return cellValue;
    }
}
