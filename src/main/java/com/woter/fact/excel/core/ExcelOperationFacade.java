package com.woter.fact.excel.core;

import com.woter.fact.excel.configuration.ExportExcelConfigurer;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

/**
 *
 * @author woter
 * @date 2017-9-29 下午8:39:52
 */
public final class ExcelOperationFacade {

    private static ExcelResolver hssfExcelResolver = new HssfExcelResolver();

    public static <T> void exportExcel(String title, List<ExcelField<?>> fieldList, List<?> data, OutputStream out) throws IOException {
        exportExcel(title, fieldList, data, out, null);
    }

    /**
     * 根据数据导出Excel文件
     * @param title Excel的title
     * @param fieldList 定义每个字段属性
     * @param data 列表数据
     * @param out 输出流
     * @param exportExcelConfigurer
     * @throws IOException
     */
    public static <T> void exportExcel(String title, List<ExcelField<?>> fieldList, List<?> data, OutputStream out, ExportExcelConfigurer exportExcelConfigurer) throws IOException {
        hssfExcelResolver.exportExcel(title, fieldList, data, out, exportExcelConfigurer);
    }

    public static <T> void exportExcel(String title, List<ExcelField<?>> fieldList, List<?> data, HttpServletResponse response, String fileName) throws IOException {
        exportExcel(title, fieldList, data, response, fileName, null);
    }


    /***
     * 根据数据导出Excel文件
     * @param title Excel的title
     * @param fieldList 定义每个字段属性
     * @param data 列表数据
     * @param response 响应流
     * @param fileName Excel文件名
     * @param exportExcelConfigurer Excel导出常用配置项
     * @throws IOException
     */
    public static <T> void exportExcel(String title, List<ExcelField<?>> fieldList, List<?> data, HttpServletResponse response, String fileName, ExportExcelConfigurer exportExcelConfigurer) throws IOException {
        hssfExcelResolver.exportExcel(title, fieldList, data, response, fileName, exportExcelConfigurer);
    }
}
