package cn.dongzhic.email;

import cn.dongzhic.excel2xml.ExcelFile;
import cn.dongzhic.excel2xml.ReadExcelUtil;
import cn.dongzhic.util.HanyuPinyinHelper;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * 提取姓名全拼、简拼
 * @Author dongzhic
 * @Date 2021/9/1 10:19
 */
public class App {

    private static  String openFilePath = "/Users/zhicheng/Desktop/email.xlsx";
    private static  String outputFilePath = "/Users/zhicheng/Desktop/邮箱_01.xlsx";

    public static void main(String[] args) throws IOException {

        //简拼全拼导出
        getEmail(outputFilePath, "邮箱_简拼全拼记录", openFilePath);

    }

    public static void getEmail(String filePath, String fileName, String openFilePath){
        getAllList(openFilePath);
    }


    private static void getAllList(String openFilePath){

        try {
            ExcelFile excelFile = new ExcelFile(openFilePath);

            if (ReadExcelUtil.checkFile(excelFile.getExcelPath())) {
                Workbook wb = ReadExcelUtil.readExcelObject(excelFile.getExcelPath());
                List<Map<String, Object>> unInitExcelList =  ReadExcelUtil.readExcelContent(wb);

                for (Map<String, Object> map : unInitExcelList) {
                    String username = map.get("id") + "";
                    // 全拼
                    map.put("B", HanyuPinyinHelper.toHanyuPinyin(username));
                    // 姓氏全拼
                    map.put("C", HanyuPinyinHelper.getFirstLetterPin(username));
                    // 简拼
                    map.put("D", HanyuPinyinHelper.getFirstLettersLo(username));
                }

                exportExcel(unInitExcelList);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    public static void exportExcel (List<Map<String, Object>> unInitExcelList) {

        // 创建工作薄
        XSSFWorkbook workbook = new XSSFWorkbook();
        // 创建工作表
        XSSFSheet sheet = workbook.createSheet("sheet1");

        //设置数据
        for (int row = 0; row < unInitExcelList.size(); row++) {
            XSSFRow sheetRow = sheet.createRow(row);
            Map<String, Object> map = unInitExcelList.get(row);
            sheetRow.createCell(0).setCellValue(map.get("id") + "");
            sheetRow.createCell(1).setCellValue(map.get("B") + "");
            sheetRow.createCell(2).setCellValue(map.get("C") + "");
            sheetRow.createCell(3).setCellValue(map.get("D") + "");
        }

        //写入文件
        try {
            workbook.write(new FileOutputStream(new File(outputFilePath)));
            workbook.close();
        } catch (Exception ex) {
            ex.printStackTrace();
            System.out.println(ex.getMessage());
        }

        System.out.println( "Hello World!" );
    }


}
