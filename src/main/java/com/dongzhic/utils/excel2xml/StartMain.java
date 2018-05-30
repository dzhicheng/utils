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

        ExcelFile excelFile = new ExcelFile("user", "用户表", "E:/用户表.xlsx");

        if (ReadExcelUtil.checkFile(excelFile.getExcelPath())) {
            Workbook wb = ReadExcelUtil.readExcelObject(excelFile.getExcelPath());
            List<Map<String, Object>> contentList =  ReadExcelUtil.readExcelContent(wb);
            Excel2Sc.createSchemeFile(contentList, excelFile);
        }

    }
}
