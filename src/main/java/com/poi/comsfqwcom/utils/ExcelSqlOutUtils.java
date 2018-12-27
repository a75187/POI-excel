package com.poi.comsfqwcom.utils;

import org.dom4j.Document;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


/**
 * EXCEL 配置文件 解析类
 * <excels>
 * <excel id/fileName/sqlMapsId>
 * id 配置id
 * fileName导出文件名
 * sqlMapsId 查询语句对应的mybatis映射地址
 * <column key/width/dataType>head</column>
 * key必填
 * head必填
 * width表格宽度
 * dataType数据类型:目前仅支持number
 * </excel>
 * </excels>
 */
public class ExcelSqlOutUtils {
    /**
     * excel配置文件读取地址
     **/
    public static final String EXCEL_XML = "excel.xml";
    private static Map<String, Map<String, Object>> EXCEL_CONFIG_MAPS = null;

    /**
     * 根节点
     **/
    public static final String CONFIG_EXCEL_ROOT = "excel";
    /**
     * 根节点 id属性
     **/
    public static final String CONFIG_EXCEL_ID = "id";
    /**
     * 根节点 导出文件名属性
     **/
    public static final String CONFIG_EXCEL_FILE_NAME = "fileName";
    /**
     * 根节点 关联mybatis映射属性
     **/
    public static final String CONFIG_EXCEL_SQL_MAPS_ID = "sqlMapsId";
    /**
     * 根节点 配置列
     **/
    public static final String CONFIG_EXCEL_COLUMNS = "columns";

    /**
     * 列节点
     **/
    public static final String CONFIG_COLUMN_ROOT = "column";
    /**
     * 列节点 查询数据属性
     **/
    public static final String CONFIG_COLUMN_KEY = "key";
    /**
     * 列节点 表头显示值
     **/
    public static final String CONFIG_COLUMN_HEAD = "head";
    /**
     * 列节点 excel宽度值
     **/
    public static final String CONFIG_COLUMN_WIDTH = "width";
    /**
     * 列节点 数据类型
     **/
    public static final String CONFIG_COLUMN_DATA_TYPE = "dataType";
    /**
     * 列节点 数据类型number
     **/
    public static final String CONFIG_COLUMN_DATA_TYPE_NUMBER = "number";
    /**
     * 列节点 数据类型枚举
     **/
    public static final String CONFIG_COLUMN_DATA_TYPE_ENUM = "enum";

    private ExcelSqlOutUtils() {
    }

    private static ExcelSqlOutUtils instance = null;

    private static ExcelSqlOutUtils getInstance() {
        if (instance == null) {
            instance = new ExcelSqlOutUtils();
        }
        return instance;
    }

    /**
     * 获取配置值
     *
     * @param key excel id
     */
    public static Map<String, Object> getConfigById(String key) throws Exception {
        if (EXCEL_CONFIG_MAPS == null) {
            getInstance().getDataByXML();
        }
        return EXCEL_CONFIG_MAPS.get(key);
    }

    /**
     * 初始化(解析XML配置文件)
     */
    private void getDataByXML() throws Exception {
        EXCEL_CONFIG_MAPS = new HashMap<>();
        InputStream inputStream = this.getClass().getClassLoader().getResourceAsStream(EXCEL_XML);
        SAXReader saxReader = new SAXReader();
        Document document = saxReader.read(inputStream);
        Element root = document.getRootElement();
        List<Element> lst = root.elements(CONFIG_EXCEL_ROOT);
        for (int i = 0; lst != null && i < lst.size(); i++) {
            Map<String, Object> resultMap = new HashMap<>();
            Element element = lst.get(i);
            //找到满足条件的节点
            resultMap.put(CONFIG_EXCEL_SQL_MAPS_ID, element.attribute(CONFIG_EXCEL_SQL_MAPS_ID).getText());
            resultMap.put(CONFIG_EXCEL_FILE_NAME, element.attribute(CONFIG_EXCEL_FILE_NAME).getText());
            resultMap.put(CONFIG_EXCEL_COLUMNS, getColumnByElement(element));
            EXCEL_CONFIG_MAPS.put(element.attribute(CONFIG_EXCEL_ID).getText(), resultMap);
        }
    }


    /**
     * 读取需要导出的列信息
     */
    private List<Map<String, Object>> getColumnByElement(Element element) {
        //读取节点的内容
        List<Element> columnList = element.elements(CONFIG_COLUMN_ROOT);
        List<Map<String, Object>> resultList = new ArrayList<>();
        for (int i = 0; columnList != null && i < columnList.size(); i++) {
            Element column = columnList.get(i);
            Map<String, Object> cell = new HashMap<>();
            cell.put(CONFIG_COLUMN_KEY, column.attribute(CONFIG_COLUMN_KEY).getText());
            cell.put(CONFIG_COLUMN_HEAD, column.getText());
            if (column.attribute(CONFIG_COLUMN_WIDTH) != null) {
                cell.put(CONFIG_COLUMN_WIDTH, column.attribute(CONFIG_COLUMN_WIDTH).getText());
            }
            if (column.attribute(CONFIG_COLUMN_DATA_TYPE) != null) {
                cell.put(CONFIG_COLUMN_DATA_TYPE, column.attribute(CONFIG_COLUMN_DATA_TYPE).getText());
            }
            resultList.add(cell);
        }
        return resultList;
    }
}
