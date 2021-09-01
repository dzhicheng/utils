package cn.dongzhic.excel2xml;

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
            item.addAttribute("id", rowMap.get("id").toString().toUpperCase());
            item.addAttribute("alias", rowMap.get("name").toString());

            String type = rowMap.get("type").toString();
            if ("int".equals(type)) {
                type = "long";
            }
            item.addAttribute("type", type);
            //length类型处理，为null时默认长度为50
            String length = rowMap.get("length").toString();
            if (!"".equals(length)) {
                item.addAttribute("length", length);
            }
            if ("1".equals(rowMap.get("isPK").toString())) {
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

            sql.append("  "+rowMap.get("id").toString());
            sql.append(" ");

            String type = rowMap.get("type").toString();
            if ("string".equals(type)) {
                String length = "null".equals(rowMap.get("length").toString()) ? "50" : rowMap.get("length").toString();
                sql.append("VARCHAR2("+length+")");
            } else if ("int".equals(type)) {
                sql.append("NUMBER");
            } else if ("date".equals(type)) {
                sql.append("DATE");
            }
            if ("1".equals(rowMap.get("isNull").toString())) {
                sql.append(" not null");
            }
            if ("1".equals(rowMap.get("isPK").toString())) {
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
        sql.append("  add constraint PK_"+excelFile.getTableName().toUpperCase()+"_"+pkValue
                +" primary key ("+pkValue+"); \n");


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

                if ("1".equals(rowMap.get("isPK").toString())) {
                    Element id = clazz.addElement("id");
                    id.addAttribute("name",rowMap.get("id").toString().toUpperCase());

                    String type = rowMap.get("type").toString().toLowerCase();
                    if ("string".equals(type)){
                        id.addAttribute("type", "java.lang.String");
                    }
                    if ("int".equals(type)) {
                        id.addAttribute("type", "java.lang.Long");
                    }
                    id.addAttribute("length", rowMap.get("length").toString());
                    Element column = id.addElement("column");
                    column.addAttribute("name", rowMap.get("id").toString().toUpperCase());
                    Element generator = id.addElement("generator");
                    generator.addAttribute("class", "assigned");
                } else {
                    Element property = clazz.addElement("property");
                    property.addAttribute("name", rowMap.get("id").toString().toUpperCase());
                    String type = rowMap.get("type").toString();
                    if (!"date".equals(type.toLowerCase())) {
                        property.addAttribute("length", rowMap.get("length").toString());
                    } else {
                        property.addAttribute("type", "date");
                    }
                    if ("string".equals(type.toLowerCase())){
                        property.addAttribute("type", "java.lang.String");
                    }
                    if ("int".equals(type.toLowerCase())) {
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
     * 对excel中字段进行操作:
     *  1、去掉字符串两边空格
     *  2、赋默认值
     * @param excelContentList
     * @return
     */
    public static List<Map<String, Object>> initExcelList (List<Map<String, Object>> excelContentList) {

        List<Map<String, Object>> resultList = new ArrayList<>();
        Map<String, Object> tempMap;

        for (int i=0 ;i<excelContentList.size(); i++) {

            Map<String, Object> resultMap = new HashMap<>(16);
            tempMap = excelContentList.get(i);

            //对id字段名处理,必填
            if ("".equals(tempMap.get("id").toString())) {
                log.error("defaultExcelList failed, 字段名id不能为空");
                return null;
            } else {
                resultMap.put("id", tempMap.get("id").toString().trim());
            }
            resultMap.put("name", tempMap.get("name").toString().trim());

            String type = tempMap.get("type").toString().trim();
            if ("".equals(type)) {
                type = "string";
            }
            resultMap.put("type", type);
            String length = tempMap.get("length").toString().trim();
            if ("".equals(length)) {
                length = "20";
            } else {
                length = length.substring(0, length.indexOf("."));
            }
            resultMap.put("length", length);
            String isNull = tempMap.get("isNull").toString();
            if ("".equals(isNull)) {
                isNull = "0";
            } else {
                isNull = length.substring(0, isNull.indexOf("."));
            }
            resultMap.put("isNull", isNull);
            String isPK = tempMap.get("isPK").toString();
            if ("".equals(isPK)) {
                isPK = "0";
            } else {
                isPK = length.substring(0, isPK.indexOf("."));
            }
            resultMap.put("isPK", isPK);

            resultList.add(resultMap);
        }

        return resultList;
    }


}
