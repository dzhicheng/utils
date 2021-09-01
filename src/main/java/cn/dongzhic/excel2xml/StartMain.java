package cn.dongzhic.excel2xml;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
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

//        ExcelFile excelFile1 = new ExcelFile("AIO_ANYCHECKTIZHI", "中医体质辨识数据临时表", "E:/bsoft/中医体质辨识1.0.xlsx");
        ExcelFile excelFile1 = new ExcelFile("AIO_MEASURE", "安测一体机传输数据临时表", "E:/bsoft/一体机接口字段1.0.xlsx");

        if (ReadExcelUtil.checkFile(excelFile1.getExcelPath())) {
            Workbook wb = ReadExcelUtil.readExcelObject(excelFile1.getExcelPath());
            List<Map<String, Object>> unInitExcelList =  ReadExcelUtil.readExcelContent(wb);
            List<Map<String, Object>> innitExcelList =  new ArrayList<Map<String, Object>>();

            innitExcelList = ExcelTransformUtils.initExcelList(unInitExcelList);

            //对ecxel中的数据进行初始化操作,省去接下来判空的步骤
            if (innitExcelList !=null && innitExcelList.size()>0) {

                ExcelTransformUtils.createSchemeFile(innitExcelList, excelFile1);
                ExcelTransformUtils.createSqlFile(innitExcelList, excelFile1);
                ExcelTransformUtils.createHbmFile(innitExcelList, excelFile1);

            }


        }

    }
}
