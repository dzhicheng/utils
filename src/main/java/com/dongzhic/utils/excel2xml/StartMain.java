package com.dongzhic.utils.excel2xml;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;
import java.util.Map;

/**
 * @desc
 * @author dongzc
 * @date 2018/5/29 22:32
 **/
@Slf4j
public class StartMain {

    public static void main(String[] args) throws Exception {
        String excelPath = "E:/用户表.xlsx";
        if (ReadExcelUtil.checkFile(excelPath)) {
            Workbook wb = ReadExcelUtil.readExcelObject(excelPath);
            List<Map<String, Object>> contentList =  ReadExcelUtil.readExcelContent(wb);
            System.out.println(contentList.size());
        }

    }
}
