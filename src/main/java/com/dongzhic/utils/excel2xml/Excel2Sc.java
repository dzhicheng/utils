package com.dongzhic.utils.excel2xml;

import lombok.extern.slf4j.Slf4j;
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
@Slf4j
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
            //length类型处理，为null时默认长度为50
            String length = rowMap.get("length")+"";
            item.addAttribute("length", "null".equals(length) ? "50" : length.substring(0, length.indexOf(".")));

            if ("1.0".equals(rowMap.get("isPK")+"")) {
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
        File file = new File("E:/bsoft/"+excelFile.getTableName()+".sc");
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

    public static void createSqlFile (List<Map<String, Object>> excelContentList, ExcelFile excelFile) {

        //写入sql
        StringBuffer sql = new StringBuffer();
        sql.append("-- Create table "+excelFile.getTableName()+"\n");
        sql.append("create table " + excelFile.getTableName()+"\n");
        sql.append("("+"\n");
        Map<String, Object> rowMap = new HashMap<String, Object>();
        for (int i=0; i<excelContentList.size(); i++) {
            rowMap = excelContentList.get(i);
            if ("null".equals(rowMap.get("id")+"")) {
                log.error("createSqlFile failed, 字段名不能为空");
                break;
            } else {
                sql.append(rowMap.get("id").toString());
                sql.append(" ");
                if ("null".equals(rowMap.get("type")+"")) {
                    sql.append("VARCHAR2(50)");
                } else {
                    if ("string".equals(rowMap.get("type").toString())) {
                        String length = "null".equals(rowMap.get("length")+"") ? "50" : rowMap.get("length").toString();
                        sql.append("VARCHAR2("+length+")");
                    } else if ("int".equals(rowMap.get("type").toString())) {
                        sql.append("NUMBER");
                    }
                }
                if ("0".equals(rowMap.get("isNull")+"")) {
                    sql.append(" not null");
                }

            }
            if (i != excelContentList.size()-1) {
                sql.append(",");
            }
            sql.append("\n");
        }

        sql.append(")");


    }



}
