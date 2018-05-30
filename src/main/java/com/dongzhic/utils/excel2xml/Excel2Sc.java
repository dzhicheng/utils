package com.dongzhic.utils.excel2xml;

import org.dom4j.Document;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.XMLWriter;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Description:
 *
 * @author: dongzhic
 * @date: 2018/5/30 18:50
 */
public class Excel2Sc {

    public static void createSchemeFile (List<Map<String, Object>> excelContentList, ExcelFile excelFile) {

        Document document = DocumentHelper.createDocument();
        Element entry = document.addElement("entry");
        entry.addAttribute("entityName", excelFile.getTableName());
        entry.addAttribute("alias", excelFile.getTableMeaning());

        Map<String, Object> rowMap = new HashMap<String, Object>();
        for (int i=0; i<excelContentList.size(); i++){

            rowMap = excelContentList.get(i);
            Element item = entry.addElement("item");
            item.addAttribute("id", rowMap.get("id")+"");
            item.addAttribute("alias", rowMap.get("name")+"");
            item.addAttribute("type", rowMap.get("type")+"");
            item.addAttribute("length", rowMap.get("length")+"");

            if ("1".equals(rowMap.get("isPK")+"")) {
                item.addAttribute("not-null", "1");
                item.addAttribute("generator", "assigned");
                item.addAttribute("pkey", "true");
                Element key = item.addElement("key");
                Element rule = key.addElement("rule");
                rule.addAttribute("name", "increaseId");
                rule.addAttribute("type", "increase");
                rule.addAttribute("length", "16");
                rule.addAttribute("startPos", "1");

            } else {
                item.addAttribute("display", "0");
            }
        }

        //
        File file = new File("E:/bsoft"+excelFile.getTableName()+".sc");
        File fileParent = file.getParentFile();
        if (!fileParent.exists()) {
            fileParent.mkdirs();
        }
        try {
            file.createNewFile();

            OutputFormat format = OutputFormat.createPrettyPrint();
            XMLWriter writer = new XMLWriter(new FileOutputStream(file), format);
            writer.write(document);
        } catch (IOException e) {
            e.printStackTrace();
        }


    }
}
