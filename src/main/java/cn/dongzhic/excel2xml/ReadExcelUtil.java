package cn.dongzhic.excel2xml;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

/**
 * @author dongzc
 * @desc
 * @date 2018/5/29 22:41
 **/
@Slf4j
public class ReadExcelUtil {

    /**
     * 固定excel第一行字段的顺序
     */
    private static String[] excelTitle = {"id", "name", "type", "length", "isNull", "isPK"};

    /**
     * 根据excel的版本获取操作对象
     * @param excelPath
     * @return
     */
    public static Workbook readExcelObject (String excelPath) {

        String ext = excelPath.substring(excelPath.lastIndexOf("."));
        try {
            InputStream is = new FileInputStream(excelPath);
            if (".xls".equals(ext)) {
                return new HSSFWorkbook(is);
            }else if (".xlsx".equals(ext)){
                return new XSSFWorkbook(is);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    public static boolean checkFile(String excelPath) {
        File file = new File(excelPath);
        if (!file.exists()) {
            log.error("文件不存在！");
            return false;
        }
        String fileName = file.getName();
        if (!fileName.endsWith("xls") && !fileName.endsWith("xlsx") ) {
            log.error(fileName+"不是excel文件");
            return false;
        }
        return true;
    }

    public static List<Map<String, Object>> readExcelContent(Workbook wb) throws Exception {
        List<Map<String, Object>> contentList = new ArrayList<Map<String, Object>>();

        if(wb==null){
            throw new Exception("Workbook对象为空！");
        }
        Sheet sheet = wb.getSheetAt(0);
        // 得到总行数
        int rowNum = sheet.getLastRowNum();
        Row row = sheet.getRow(0);
        int colNum = row.getPhysicalNumberOfCells();
        for (int i=1; i<= rowNum; i++) {
            row = sheet.getRow(i);
            int j = 0;
            Map<String,Object> cellValue = new HashMap<String, Object>();
            while (j < colNum) {
                Object obj = getCellFormatValue(row.getCell(j));
                cellValue.put(excelTitle[j], obj);
                j++;
            }
            contentList.add(cellValue);
        }
        return contentList;
    }

    /**
     *
     * 根据Cell类型设置数据
     *
     * @param cell
     * @return
     * @author zengwendong
     */
    private static Object getCellFormatValue(Cell cell) {
        Object cellvalue = "";
        if (cell != null) {
            // 判断当前Cell的Type
            switch (cell.getCellType()) {
                // 如果当前Cell的Type为NUMERIC
                case Cell.CELL_TYPE_NUMERIC:
                case Cell.CELL_TYPE_FORMULA: {
                    // 判断当前的cell是否为Date
                    if (DateUtil.isCellDateFormatted(cell)) {
                        // 如果是Date类型则，转化为Data格式
                        // data格式是带时分秒的：2013-7-10 0:00:00
                        // cellvalue = cell.getDateCellValue().toLocaleString();
                        // data格式是不带带时分秒的：2013-7-10
                        Date date = cell.getDateCellValue();
                        cellvalue = date;
                    } else {// 如果是纯数字

                        // 取得当前Cell的数值
                        cellvalue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                }
                // 如果当前Cell的Type为STRING
                case Cell.CELL_TYPE_STRING:
                    // 取得当前的Cell字符串
                    cellvalue = cell.getRichStringCellValue().getString();
                    break;
                // 默认的Cell值
                default:
                    cellvalue = "";
            }
        } else {
            cellvalue = "";
        }
        return cellvalue;
    }
}
