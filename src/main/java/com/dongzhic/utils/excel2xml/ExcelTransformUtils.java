package com.dongzhic.utils.excel2xml;

import lombok.extern.slf4j.Slf4j;
import org.dom4j.Document;
import org.dom4j.DocumentHelper;
import org.dom4j.Element;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.XMLWriter;

import java.io.*;
import java.util.ArrayList;
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
public class ExcelTransformUtils {

    /**
     *  生成scheme文件
     * @param excelContentList
     * @param excelFile
     */
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

    /**
     *  生成Sql语句文件
     * @param excelContentList
     * @param excelFile
     */
    public static void createSqlFile (List<Map<String, Object>> excelContentList, ExcelFile excelFile) {

        String pkValue = "";

        //写入sql
        StringBuffer sql = new StringBuffer();
        sql.append("-- Create table "+excelFile.getTableName()+"\n");
        sql.append("create table " + excelFile.getTableName()+"\n");
        sql.append("("+"\n");
        Map<String, Object> rowMap = new HashMap<String, Object>();
        for (int i=0; i<excelContentList.size(); i++) {
            rowMap = excelContentList.get(i);

            sql.append(rowMap.get("id").toString());
            sql.append(" ");

            if ("string".equals(rowMap.get("type").toString())) {
                String length = "null".equals(rowMap.get("length")+"") ? "50" : rowMap.get("length").toString();
                length = length.substring(0, length.indexOf("."));
                sql.append("VARCHAR2("+length+")");
            } else if ("int".equals(rowMap.get("type").toString())) {
                sql.append("NUMBER");
            }
            if ("1.0".equals(rowMap.get("isNull")+"")) {
                sql.append(" not null");
            }
            if ("1.0".equals(rowMap.get("isPK")+"")) {
                pkValue = rowMap.get("id").toString().toUpperCase();
            }

            if (i != excelContentList.size()-1) {
                sql.append(",");
            }
            sql.append("\n");
        }
        sql.append("); \n");
        sql.append("comment on table "+excelFile.getTableName()+"\n");
        sql.append("  is '"+excelFile.getTableMeaning()+"'; \n");
        for (int i =0; i<excelContentList.size(); i++) {
            sql.append("comment on column "+excelFile.getTableName()+"."+excelContentList.get(i).get("id").toString()+"\n");
            sql.append("  is '"+excelContentList.get(i).get("name").toString()+"'; \n");
        }

        sql.append("-- Create/Recreate primary, unique and foreign key constraints \n");
        sql.append("alter table "+excelFile.getTableName()+"\n");
        sql.append("  add constraint PK_"+excelFile.getTableName().toUpperCase()+" primary key ("+pkValue+"); \n");


        File file = new File("E:/bsoft/"+excelFile.getTableName()+".sql");
        File fileParent = file.getParentFile();
        if (!fileParent.exists()) {
            fileParent.mkdirs();
        }

        BufferedWriter out = null;
        try {
            file.createNewFile();

            out = new BufferedWriter(new FileWriter(file));
            out.write(sql.toString());
            out.flush();
            out.close();

        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }

    }


    public static void createHbmFile (List<Map<String, Object>> excelContentList, ExcelFile excelFile) {

            Document document = DocumentHelper.createDocument();

            document.addDocType("hibernate-mapping", "-//Hibernate/Hibernate Mapping DTD 3.0//EN",
                    "http://www.hibernate.org/dtd/hibernate-mapping-3.0.dtd");
            Element hibernate = document.addElement("hibernate-mapping");
            Element clazz = hibernate.addElement("class");
            clazz.addAttribute("entity-name", excelFile.getTableName().toUpperCase());
            clazz.addAttribute("table", excelFile.getTableName().toUpperCase());

            Map<String, Object> rowMap = new HashMap<String, Object>();
            for (int i=0; i<excelContentList.size(); i++) {
                rowMap = excelContentList.get(i);

                if ("1.0".equals(rowMap.get("isPK")+"")) {
                    Element id = clazz.addElement("id");
                    id.addAttribute("name",rowMap.get("id").toString());
                    if ("string".equals(rowMap.get("type").toString().toLowerCase())){
                        id.addAttribute("type", "java.lang.String");
                    }
                    if ("int".equals(rowMap.get("type").toString().toLowerCase())) {
                        id.addAttribute("type", "java.lang.Integer");
                    }
                    Element column = id.addElement("column");
                    column.addAttribute("name", rowMap.get("id").toString());
                    Element generator = id.addElement("generator");
                    generator.addAttribute("class", "assigned");
                } else {
                    Element property = clazz.addElement("property");
                    property.addAttribute("name", rowMap.get("id").toString());
                    property.addAttribute("length", rowMap.get("length").toString().substring(0, rowMap.get("length").toString().indexOf(".")));
                    if ("string".equals(rowMap.get("type").toString().toLowerCase())){
                        property.addAttribute("type", "java.lang.String");
                    }
                    if ("int".equals(rowMap.get("type").toString().toLowerCase())) {
                        property.addAttribute("type", "java.lang.Integer");
                    }
                }
            }

            //
            File file = new File("E:/bsoft/"+excelFile.getTableName()+".hbm.xml");
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

    /**
     * 对excel中的空字段进行默认值操作
     * @param excelContentList
     * @return
     */
    public static List<Map<String, Object>> defaultExcelList (List<Map<String, Object>> excelContentList) {

        List<Map<String, Object>> resultList = new ArrayList<Map<String, Object>>();
        Map<String, Object> resultMap = new HashMap<String, Object>(16);
        Map<String, Object> tempMap = new HashMap<String, Object>();

        for (int i=0 ;i<excelContentList.size(); i++) {
            tempMap = excelContentList.get(i);

            //对id字段名处理,必填
            if ("null".equals(tempMap.get("id")+"")) {
                log.error("defaultExcelList failed, 字段名id不能为空");
                return null;
            } else {
                resultMap.put("id", tempMap.get("id").toString());
            }
            //对name字段名称处理
            resultMap.put("name", "null".equals(tempMap.get("name")+"") ? "" : tempMap.get("name").toString());
            resultMap.put("type", "null".equals(tempMap.get("type")+"") ? "string" : tempMap.get("type").toString());
            resultMap.put("length", "null".equals(tempMap.get("length")+"") ? "50" : tempMap.get("length").toString());
            resultMap.put("isNull", "null".equals(tempMap.get("isNull")+"") ? "0" : tempMap.get("isNull").toString());
            resultMap.put("idPK", "null".equals(tempMap.get("idPK")+"") ? "0" : tempMap.get("idPK").toString());

            resultList.add(resultMap);
        }

        return resultList;
    }


}
