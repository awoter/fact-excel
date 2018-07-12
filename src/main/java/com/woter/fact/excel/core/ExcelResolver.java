package com.woter.fact.excel.core;

import com.woter.fact.excel.configuration.ExportExcelConfigurer;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

/**
 * @author woter
 * @date 2017-9-29 下午2:48:47
 */
public interface ExcelResolver {

    void exportExcel(String title, List<ExcelField<?>> fieldList, List<?> list, OutputStream out, ExportExcelConfigurer exportExcelConfigurer) throws IOException;

    void exportExcel(String title, List<ExcelField<?>> fieldList, List<?> list, HttpServletResponse response, String fileName, ExportExcelConfigurer exportExcelConfigurer) throws IOException;
}
