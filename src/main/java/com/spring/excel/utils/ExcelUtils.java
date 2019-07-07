package com.spring.excel.utils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.StringUtils;

import javax.servlet.ServletOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.Collection;
import java.util.List;
import java.util.Map;
/**
 * 导出excel工具类
 * @author 福小林
 */
public class ExcelUtils {
    /**
     * 日志框架对象
     */
    private static final Logger logger = LoggerFactory.getLogger(ExcelUtils.class);
    private ExcelUtils() {

    }
    /**
     *
     * @param titleMap excel标题map key-->标题对应的英文属性，value-->对应的中文标题名称
     * @param out  输出流对象
     * @param list 需要打印在excel上面的数据集合
     * @param <T> 泛型标识
     */
    public static <T> void exportExcel(Map<String,String> titleMap,ServletOutputStream out, List<T> list) {
        try{
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet hssfSheet = getRows(titleMap, workbook);
            setExcelData(titleMap, list, hssfSheet);
            closeServletOutputStream(out, workbook);
        }catch(Exception e){
            logger.error("导出excel信息失败！"+e.getMessage(),e);
        }
    }

    /**
     * 关闭excel导出的各种流
     * @param out  输出流对象
     * @param workbook   HSSFWorkbook对象
     * @throws IOException 可能抛出的IO异常
     */
    private static void closeServletOutputStream(ServletOutputStream out, HSSFWorkbook workbook) throws IOException {
        workbook.write(out);
        out.flush();
        out.close();
    }
    /**
     * 把数据放入excel表格中
     * @param titleNames  excel标题map key-->标题对应的英文属性，value-->对应的中文标题名称
     * @param list 需要打印在excel上面的数据集合
     * @param hssfSheet HSSFSheet对象
     * @param <T> 泛型标识
     * @throws Exception 异常
     */
    private static <T>  void setExcelData(Map<String,String> titleNames,List<T> list, HSSFSheet hssfSheet) throws  Exception {
        HSSFRow row;
        //循环数据，打印到excel
        for (int i = 0; i < list.size(); i++) {
            row = hssfSheet.createRow(i+1);
            //获取list里面的对象
            T t = list.get(i);
            //如果list集合的类型是Map<String,Object>
            if (t instanceof Map){
                Map map=(Map) t;
                int k=0;
                for(Map.Entry<String,String> entry :titleNames.entrySet()){
                    //标题英文字段
                    String key = entry.getKey();
                    Object obj = map.get(key);
                    //excel每个单元格插入的数据
                    String excelStr=getObjStr(obj);
                    row.createCell(k).setCellValue(excelStr);
                    k++;
                }
                //如果list集合的类型是实体类对象
            }else {
                int k=0;
                for(Map.Entry<String,String> entry :titleNames.entrySet()){
                    String key = entry.getKey();
                    //获取对应实体类Class对象
                    Class < ?>cls = t.getClass();
                    //把t的数据赋值给obj
                    Object obj = cls.cast(t);
                    //获取标题对应的哪一个字段对象
                    Field field = cls.getDeclaredField(key);
                    //可以访问私有字段数据
                    field.setAccessible(true);
                    //获取对应的字段值
                    Object fieldValue = field.get(obj);
                    //excel每个单元格插入的数据
                    String excelStr = getObjStr(fieldValue);
                    row.createCell(k).setCellValue(excelStr);
                    k++;
                }
            }
        }
    }

    /**
     * 获取对象的字符串
     * @param obj Object对象
     * @return  String
     */
    private static String getObjStr(Object obj) {
        if (StringUtils.isEmpty(obj)){
            return "";
        }else {
            return obj.toString();
        }
    }


    /**
     *
     * @param titleMap excel标题map key-->标题对应的英文属性，value-->对应的中文标题名称
     * @param workbook HSSFWorkbook对象
     * @return HSSFSheet已经把标题和样式设置和打印完成的对象
     */
    private static  HSSFSheet getRows(Map<String,String> titleMap, HSSFWorkbook workbook) {
        Collection<String> valueCollection = titleMap.values();
        final int size = valueCollection.size();
        String[] titles = new String[size];
        titleMap.values().toArray(titles);
        HSSFSheet hssfSheet = workbook.createSheet("sheet1");
        HSSFRow row = hssfSheet.createRow(0);
        HSSFCellStyle hssfCellStyle = workbook.createCellStyle();
        hssfCellStyle.setAlignment(HorizontalAlignment.CENTER);
        HSSFCell hssfCell ;
        for (int i = 0; i < titles.length; i++) {
            hssfCell = row.createCell(i);
            hssfCell.setCellValue(titles[i]);
            hssfCell.setCellStyle(hssfCellStyle);
        }
        return hssfSheet;
    }

}
