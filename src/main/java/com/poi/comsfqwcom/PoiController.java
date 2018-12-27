package com.poi.comsfqwcom;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.poi.comsfqwcom.utils.ExcelSqlOutUtils;
import org.apache.poi.hssf.usermodel.*;
import org.junit.jupiter.api.Test;
import org.springframework.util.StringUtils;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.util.List;
import java.util.Map;

public class PoiController {

    @Test
    public void excelExportX() throws Exception {
        String key="mytestpost";
        Map<String, Object> excelMap = ExcelSqlOutUtils.getConfigById(key);
        String fileName = "";
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("sheet");
        if (excelMap != null) {
            List<Map<String, Object>> columns = (List<Map<String, Object>>) excelMap.get(ExcelSqlOutUtils.CONFIG_EXCEL_COLUMNS);
            fileName = (String) excelMap.get(ExcelSqlOutUtils.CONFIG_EXCEL_FILE_NAME);
            JSONArray array = toList();
            createHeader(sheet, columns);
            createBook(workbook, sheet, array, columns);
        }
        exportExcelx(workbook, fileName);
    }

    private void exportExcelx(HSSFWorkbook workbook , String fileName) throws Exception {
        String exportFileName = StringUtils.isEmpty(fileName) ? "电子表格.xls" : fileName + ".xls";
        BufferedOutputStream bufferedOutputStream = new BufferedOutputStream(new FileOutputStream(exportFileName));
        workbook.write(bufferedOutputStream);
        bufferedOutputStream.close();
    }

    public JSONArray toList(){
        String josn="{\"result\":{\"currentPage\":1,\"currentPageStartIndex\":0,\"items\":[{\"bedRoomNum\":\"02卧室\",\"buildingName\":\"A单元\",\"createOperatorId\":\"\",\"endTime\":\"2019-01-31\",\"gardenName\":\"宝民花园\",\"id\":\"\",\"managerName\":\"安东尼\",\"refundLeaseNo\":\"TZ201812261152\",\"renterName\":\"申达股份\",\"roomNumber\":\"202\",\"startTime\":\"2018-12-25\",\"status\":\"WAIT_PAY\",\"statusDesc\":\"待支付\",\"sumMoney\":1.5,\"time\":\"2018-12-26\",\"updateOperatorId\":\"\"},{\"bedRoomNum\":\"02卧室\",\"buildingName\":\"A栋1单元\",\"createOperatorId\":\"\",\"endTime\":\"2019-02-22\",\"gardenName\":\"大冲商务中心\",\"id\":\"\",\"managerName\":\"饺子\",\"refundLeaseNo\":\"TZ201812269316\",\"renterName\":\"辣辣\",\"roomNumber\":\"307\",\"startTime\":\"2018-12-26\",\"status\":\"WAIT_CHECK\",\"statusDesc\":\"待处理\",\"sumMoney\":-3399,\"time\":\"2018-12-26\",\"updateOperatorId\":\"\"},{\"bedRoomNum\":\"02卧室\",\"buildingName\":\"A栋1单元\",\"createOperatorId\":\"\",\"endTime\":\"2019-02-22\",\"gardenName\":\"大冲商务中心\",\"id\":\"\",\"managerName\":\"LL\",\"refundLeaseNo\":\"TZ201812267865\",\"renterName\":\"房费\",\"roomNumber\":\"307\",\"startTime\":\"2018-12-01\",\"status\":\"WAIT_CHECK\",\"statusDesc\":\"待处理\",\"sumMoney\":0,\"time\":\"2018-12-26\",\"updateOperatorId\":\"\"},{\"bedRoomNum\":\"03卧室\",\"buildingName\":\"A栋1单元\",\"createOperatorId\":\"\",\"endTime\":\"2019-02-07\",\"gardenName\":\"大冲商务中心\",\"id\":\"\",\"managerName\":\"LL\",\"refundLeaseNo\":\"TZ201812259307\",\"renterName\":\"发个\",\"roomNumber\":\"307\",\"startTime\":\"2018-12-01\",\"status\":\"WAIT_CHECK\",\"statusDesc\":\"待处理\",\"sumMoney\":0,\"time\":\"2018-12-25\",\"updateOperatorId\":\"\"},{\"bedRoomNum\":\"03卧室\",\"buildingName\":\"B单元2栋\",\"createOperatorId\":\"\",\"endTime\":\"2019-03-07\",\"gardenName\":\"兴隆大厦\",\"id\":\"\",\"managerName\":\"LL\",\"refundLeaseNo\":\"TZ201812255641\",\"renterName\":\"哎哎哎\",\"roomNumber\":\"202\",\"startTime\":\"2018-12-25\",\"status\":\"WAIT_CHECK\",\"statusDesc\":\"待处理\",\"sumMoney\":0,\"time\":\"2018-12-25\",\"updateOperatorId\":\"\"}],\"nextPage\":2,\"pageCount\":108,\"pageSize\":5,\"previousPage\":1,\"queryRecordCount\":true,\"recordCount\":540,\"uri\":\"\"},\"message\":\"处理成功\",\"status\":\"C0000\"}";
        JSONObject jsonObject = JSON.parseObject(josn);
        Map  map= (Map) jsonObject.get("result");
        JSONArray items = (JSONArray) map.get("items");
        return items;

    }


    /**
     * 构建表头
     */
    private void createHeader(HSSFSheet sheet, List<Map<String, Object>> columns) {
        HSSFRow row = sheet.createRow(0);
        for (int index = 0; index < columns.size(); index++) {
            Map<String, Object> column = columns.get(index);
            HSSFCell cell = row.createCell(index);
            cell.setCellValue((String) column.get(ExcelSqlOutUtils.CONFIG_COLUMN_HEAD));
            sheet.setColumnWidth(index, column.get("width") == null
                    ? 200 : new BigDecimal((String) column.get("width")).intValue() * 36);
        }
    }

    /**
     * 构建数据表格
     */
    private void createBook(HSSFWorkbook workbook, HSSFSheet sheet,
                            JSONArray array, List<Map<String, Object>> columns) {
        if (array == null) {
            return;
        }
        //数据格式化
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));

        for (int i = 0; i < array.size(); i++) {
            JSONObject item = array.getJSONObject(i);
            HSSFRow row = sheet.createRow(i + 1);
            for (int index = 0; index < columns.size(); index++) {
                Map<String, Object> column = columns.get(index);
                HSSFCell cell = row.createCell(index);
                try {
                    String value = item.getString((String) column.get(ExcelSqlOutUtils.CONFIG_COLUMN_KEY));
                    if (value == null || value.equals("") || value.equals("{}")) {
                        value = "";
                    }
                    String dataType = (String) column.get(ExcelSqlOutUtils.CONFIG_COLUMN_DATA_TYPE);
                    if (dataType==null) {
                        if (ExcelSqlOutUtils.CONFIG_COLUMN_DATA_TYPE_NUMBER.equals(dataType)) {
                            cell.setCellStyle(cellStyle);
                            cell.setCellValue(Double.valueOf(value));
                        } else {
                            cell.setCellValue(value);
                        }
                    } else {
                        cell.setCellValue(value);
                    }
                } catch (Exception e) {
                    cell.setCellValue("");
                }
            }
        }
    }
}

