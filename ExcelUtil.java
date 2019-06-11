package org.business.util;

import org.apache.poi.hssf.usermodel.*;

import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @Author Demo_Liu
 * @Date 2019/6/5 17:25
 * @description 封装 apache POI
 */
/**
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi</artifactId>
    <version>3.9</version>
</dependency>
**/
public class ExcelUtil {

    /**
     * @Author Demo_Liu
     * @Date 2019/6/6 14:33
     * @description 根据实体类 和 字段 创建 HSSFWorkbook
     * @Param [configMap map<实体字段,Excel表头>, entityList, rows sheet页行数, fileds 按照实体类字段 最后一行如果数据重复那么不分页, formatByFiled<字段,日期格式>]
     * @Reutrn org.apache.poi.hssf.usermodel.HSSFWorkbook
    */
    public static HSSFWorkbook getWorkBook(Map<String, String> configMap, List<?> entityList, int rows, String[] fileds, Map<String,String> formatByFiled) throws Exception {
        if(configMap.getClass() != LinkedHashMap.class)throw new Exception("Map Type Error:"+configMap.getClass()+",should be java.util.LinkedHashMap");
        if(rows<=0)throw new Exception("rows is too small,rows:"+rows);
        HSSFWorkbook wb = new HSSFWorkbook();
        int size = entityList.size();
        int sheets;
        if(size%rows==0){
            sheets = size/rows;
        }else{
            sheets = size/rows+1;
        }
        // 创建一个居中格式
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);

        HSSFSheet sheet;
        HSSFRow row;
        HSSFCell cell;

        int j=0;
        for(int i=1;i<=sheets;i++){
            sheet = wb.createSheet(i+"");
            row = sheet.createRow(0);
            cell = row.createCell(0);
            cell.setCellValue("序号");
            cell.setCellStyle(style);
            int line = 1;
            for (String s : configMap.keySet()) {
                cell = row.createCell(line);
                cell.setCellValue(configMap.get(s));
                cell.setCellStyle(style);
                line++;
            }
            for(int h=1,s=1;j<size && h<=rows;h++,j++,s++){
                Class<?> entityClass = entityList.get(j).getClass();
                if(j<size-1 && h==rows && fileds!=null){
                    boolean filedCheck = true;
                    Class<?> entityClass2 = entityList.get(j+1).getClass();
                    for (String filed : fileds) {
                        String str1 = filed.substring(0, 1);
                        String str2 = filed.substring(1);
                        String methodGet = "get"+str1.toUpperCase() + str2;
                        if(!getStrByObj(entityClass.getMethod(methodGet).invoke(entityList.get(j)),formatByFiled.get(filed)).equals(getStrByObj(entityClass2.getMethod(methodGet).invoke(entityList.get(j+1)),formatByFiled.get(filed)))){
                            filedCheck = false;
                        }
                    }
                    if(filedCheck){
                        h--;
                    }
                }
                row = sheet.createRow(s);
                cell = row.createCell(0);
                cell.setCellValue(j+1);
                cell.setCellStyle(style);
                int line2 = 1;
                for (String filed : configMap.keySet()) {
                    String str1 = filed.substring(0, 1);
                    String str2 = filed.substring(1);
                    String methodGet = "get"+str1.toUpperCase() + str2;
                    cell = row.createCell(line2);
                    cell.setCellValue(getStrByObj(entityClass.getMethod(methodGet).invoke(entityList.get(j)),formatByFiled.get(filed)));
                    cell.setCellStyle(style);
                    line2++;
                }
            }
        }
        return wb;
    }

    public static HSSFWorkbook getWorkBook(Map<String, String> configMap, List<?> entityList, int rows, String[] fileds) throws Exception{
        return getWorkBook(configMap,entityList,rows,fileds,new HashMap<String, String>());
    }
    public static HSSFWorkbook getWorkBook(Map<String, String> configMap, List<?> entityList, int rows) throws Exception{
        return getWorkBook(configMap,entityList,rows,null,new HashMap<String, String>());
    }
    /**
     * Object 获取String
     * @param obj
     * @param dateFormat
     * @return
     */
    public static String getStrByObj(Object obj, String dateFormat){
        if(null == obj){
            return "";
        }
        Class<?> objClass = obj.getClass();
        if(objClass == Date.class || objClass == java.sql.Date.class || objClass == java.sql.Timestamp.class){
            if(dateFormat == null || "".equals(dateFormat)){
                dateFormat = "yyyy-MM-dd HH:mm:ss";
            }
            SimpleDateFormat sdf = new SimpleDateFormat(dateFormat);
            return sdf.format((Date)obj);
        }else{
            return String.valueOf(obj);
        }

    }

}
