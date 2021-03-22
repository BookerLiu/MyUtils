package org.business.util;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @Author Demo-Liu
 * @Date 2021-03-17 10:25
 * @description 封装 apache POI
 */
/**
 <dependency>
     <groupId>org.apache.poi</groupId>
     <artifactId>poi</artifactId>
     <version>3.9</version>
 </dependency>
 其它版本poi依赖, 部分方法可能会有不同, 需自行修改
 **/
public class POIExcelUtil {

    /**
     * sheet允许最大行数
     */
    private final static int SHEET_MAX_ROW = 65536;
    /**
     * sheet默认最大行数
     */
    private final static int DEFAULT_SHEET_MAX_ROW = 60000;
    /**
     * 单元格默认最大宽度
     */
    private final static int DEFAULT_MAX_CELL_WIDTH = 10000;

    /**
     * 第一列 为序号列
     */
    private final static String SN_TITLE = "序号";

    /**
     * 序号列 title 宽度
     */
    private final static int SN_WIDTH = SN_TITLE.getBytes().length;

    /**
     * 默认时间字段 格式化格式
     */
    private final static String DEFAULT_DATE_FORMAT = "yyyy-MM-dd HH:mm:ss";


    /**
     * 根据数据创建HSSFWorkbook
     * @param dataList 数据, 可传入数据类型 List<T>,  List<Map<String, Object>
     *
     * @param fieldMap  字段=表头对应 有序map,   Map<字段, 表头>  例: Map<"test", "表头列">
     *
     * @param sheetMaxRow 每个sheet页行数
     *
     * @param dateFormatMap 日期字段=格式化格式   Map<日期字段, 格式化格式>  例: Map<"testDate", "yyyy-MM-dd">
     *
     * @param decimalFormatMap 数字字段=格式化格式 Map<数字字段, 格式化格式> 例: Map<"money", "#0.00">
     *
     * @param repeatFieldVal 如果字段值重复 那么不换sheet页
     *                       比如最大sheet页行数为 10, repeatFieldVal 包含字段 {"test1", "test2"}
     *                       如果下一条数据的对应值与第10行数据相同  那么不换sheet页
     * @return
     * @throws Exception
     */
    public static HSSFWorkbook getWorkBook(List<?> dataList, Map<String, String> fieldMap, Integer sheetMaxRow, Map<String, String> dateFormatMap, Map<String, String> decimalFormatMap, String...repeatFieldVal) throws Exception {

        if (sheetMaxRow == null || sheetMaxRow <= 0) {
            sheetMaxRow = DEFAULT_SHEET_MAX_ROW;
        } else if (sheetMaxRow > SHEET_MAX_ROW) {
            sheetMaxRow = SHEET_MAX_ROW;
        }
        //默认第一行为 表头  数据 +1
        sheetMaxRow += 1;

        // 检查字段map是否为 有序map
        if (fieldMap.getClass() != LinkedHashMap.class) throw new Exception("{fieldMap} Type Error:"+fieldMap.getClass()+",should be java.util.LinkedHashMap");


        //存储最大列宽  使支持中文
        Map<Integer,Integer> cellWidthMap = new HashMap<>();

        HSSFWorkbook wb = new HSSFWorkbook();
        // 创建一个居中格式
        HSSFCellStyle cstyle = getCenterStyle(wb);
        // 创建一个加粗 居中 边框格式  用于表头
        HSSFCellStyle cbbStyle = getCenterBoldBorderStyle(wb);


        Set<Map.Entry<String, String>> fieldEntry = fieldMap.entrySet();
        int dataIndex = 0;
        HSSFSheet sheet;
        HSSFRow row;

        int dataSize = 0;
        boolean isMap = false;
        //判断数据类型 实体类 或 map
        if (dataList != null && dataList.size() > 0) {
            Object obj = dataList.get(0);
            isMap = obj instanceof Map;
            dataSize = dataList.size();
        }

        //计算需要多少个 sheet页
        int sheetCount = (int)Math.ceil((double)dataSize / (double)sheetMaxRow);
        for (int i = 1; i <= sheetCount || dataSize==0; i++) {
            sheet = wb.createSheet("Sheet" + i);
            //序号列
            row = sheet.createRow(0);
            sheet.setDefaultColumnStyle(0, cstyle);
            row.createCell(0).setCellValue(SN_TITLE);
            cellWidthMap.put(0, SN_WIDTH);
            //创建表头
            int cellIndex = 1;
            for (Map.Entry<String, String> entry : fieldEntry) {
                sheet.setDefaultColumnStyle(cellIndex,cstyle);
                row.createCell(cellIndex).setCellValue(entry.getValue()); //设置title
                cellWidthMap.put(cellIndex, entry.getValue().getBytes().length); //存储title列宽
                cellIndex ++;
            }
            //如果数据为空 创建表头后即返回
            if (dataSize == 0) {
                autoCellAndTitleStyle(sheet, cbbStyle, cellWidthMap);
                break;
            }

            // 0 index 为表头
            int rowIndex = 1;
            String cellValue;
            String fieldKey;

            //写入数据
            for (; dataIndex < dataSize; dataIndex++, rowIndex ++) {
                //序号列
                row = sheet.createRow(rowIndex);
                row.createCell(0).setCellValue(dataIndex + 1);
                cellWidthMap.put(0, Math.max(cellWidthMap.get(0), String.valueOf(dataIndex + 1).getBytes().length));
                //写入 dataList
                cellIndex = 1;
                for (Map.Entry<String, String> entry : fieldEntry) {
                    fieldKey = entry.getKey();
                    cellValue = getStrByObj(getVal(dataList.get(dataIndex), fieldKey, isMap), fieldKey, dateFormatMap, decimalFormatMap);
                    row.createCell(cellIndex).setCellValue(cellValue);
                    cellWidthMap.put(cellIndex, Math.max(cellWidthMap.get(cellIndex), cellValue.getBytes().length));
                    cellIndex ++;
                }
                // 字段重复值不换sheet页
                // 满足条件:
                // 1.当前行达到最大sheet行
                // 2.当前行小于允许最大行数
                // 3.当前数据不为最后一条数据
                if (repeatFieldVal.length > 0 && rowIndex + 1 == sheetMaxRow && dataIndex < dataSize - 1) {
                    //判断 dataIndex 数据  是否 与 dataIndex +1 repeatFieldVal 数据相同
                    boolean repeat = true;
                    for (String field : repeatFieldVal) {
                        //比较当前数据 与 下一条数据是否相同 如果相同 则不换sheet页
                        if (!getStrByObj(getVal(dataList.get(dataIndex), field, isMap), field, dateFormatMap, decimalFormatMap).equals(
                                getStrByObj(getVal(dataList.get(dataIndex + 1), field, isMap), field, dateFormatMap, decimalFormatMap))) {
                            repeat = false;
                            break;
                        }
                    }
                    if (!repeat) {
                        //设置自适应列宽
                        autoCellAndTitleStyle(sheet, cbbStyle, cellWidthMap);
                        dataIndex ++;
                        break;
                    }
                } else if (rowIndex + 1 >= sheetMaxRow || dataIndex +1 == dataSize) {
                    //设置自适应列宽
                    autoCellAndTitleStyle(sheet, cbbStyle, cellWidthMap);
                    dataIndex ++;
                    break;
                }
            }
        }
        return wb;
    }

    public static HSSFWorkbook getWorkBook(List<?> entityList, Map<String, String> fieldMap) throws Exception {
        return getWorkBook(entityList, fieldMap, null, null, null);
    }
    public static HSSFWorkbook getWorkBook(List<?> entityList, Map<String, String> fieldMap, Integer sheetMaxRow) throws Exception {
        return getWorkBook(entityList, fieldMap, sheetMaxRow, null, null);
    }
    public static HSSFWorkbook getWorkBook(List<?> entityList, Map<String, String> fieldMap, Map<String, String> dateFormatMap) throws Exception {
        return getWorkBook(entityList, fieldMap, null, dateFormatMap, null);
    }

    /**
     * 设置第一行表头样式  和 自适应宽度
     * @param sheet sheet
     * @param titleStyle 表头样式
     * @param cellWidthMap  记录的cell列最大宽度
     */
    private static void autoCellAndTitleStyle(HSSFSheet sheet, HSSFCellStyle titleStyle, Map<Integer, Integer> cellWidthMap) {
        HSSFRow row = sheet.getRow(0);
        for (Integer cellIndex : cellWidthMap.keySet()) {
            row.getCell(cellIndex).setCellStyle(titleStyle);
            int cellWidth = cellWidthMap.get(cellIndex) * 256 + 100;
            sheet.setColumnWidth(cellIndex, Math.min(cellWidth, DEFAULT_MAX_CELL_WIDTH));
        }
    }

    /**
     * 获取一个居中格式
     * @param wb
     * @return
     */
    private static HSSFCellStyle getCenterStyle(HSSFWorkbook wb) {
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        return style;
    }

    /**
     * 获取加粗 居中 边框 样式
     * @param wb wb
     * @return
     */
    private static HSSFCellStyle getCenterBoldBorderStyle(HSSFWorkbook wb) {
        HSSFFont font = wb.createFont();
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
//        font.setBold(true);
        HSSFCellStyle styleBold = wb.createCellStyle();
        styleBold.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        styleBold.setFont(font);
        styleBold.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        styleBold.setBorderRight(HSSFCellStyle.BORDER_THIN);
        styleBold.setBorderTop(HSSFCellStyle.BORDER_THIN);
        styleBold.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        styleBold.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
        styleBold.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        return styleBold;
    }


    /**
     * 获取 字段值
     * @param obj 实体类 或 map
     * @param field 字段
     * @param isMap 是否为 map
     * @return
     * @throws NoSuchMethodException
     * @throws InvocationTargetException
     * @throws IllegalAccessException
     */
    private static Object getVal(Object obj, String field, boolean isMap) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        if (isMap) {
            return ((Map<?, ?>)obj).get(field);
        } else {
            Class<?> entityClass = obj.getClass();
            String getMethod = "get" + field.substring(0,1).toUpperCase() + field.substring(1);
            return entityClass.getMethod(getMethod).invoke(obj);
        }
    }



    /**
     * Object 获取String
     * @param obj val
     * @param field 字段 key 值
     * @param dateFormatMap 格式化map
     * @return
     */
    private static String getStrByObj(Object obj, String field, Map<String, String> dateFormatMap, Map<String, String> decimalFormatMap){
        if(null == obj) return "";
        Class<?> objClass = obj.getClass();
        if(objClass == Date.class || objClass == java.sql.Date.class || objClass == java.sql.Timestamp.class){
            String dateFormat = DEFAULT_DATE_FORMAT;
            if(dateFormatMap != null && dateFormatMap.containsKey(field)){
                dateFormat = dateFormatMap.get(field);
            }
            SimpleDateFormat sdf = new SimpleDateFormat(dateFormat);
            return sdf.format((Date)obj);
        }else if(decimalFormatMap != null && decimalFormatMap.containsKey(field)){
            String format = decimalFormatMap.get(field);
            DecimalFormat df = new DecimalFormat(format);
            return df.format(obj);
        }else{
            return String.valueOf(obj);
        }
    }

    /**
     * 将Excel文件传至客户端
     * @param response response
     * @param wb HSSFWorkbook
     * @param fileName 文件名.xls
     * @throws IOException
     */
    public static void writeHSSFWorkbook(HttpServletResponse response, HSSFWorkbook wb, String fileName) throws IOException {
        OutputStream output = null;
        try{
            output = response.getOutputStream();
            response.reset();
            response.setHeader("Content-Disposition", "attachment; filename="+ java.net.URLEncoder.encode(fileName, "UTF-8"));
            response.setContentType("application/ms-excel" + ";charset=GBK");
            wb.write(output);
            output.flush();
        }finally {
            if(output != null) output.close();
        }
    }

}
