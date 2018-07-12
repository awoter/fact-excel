package com.woter.fact.excel.core;

import com.google.common.base.Preconditions;
import com.google.common.collect.Maps;
import com.woter.fact.excel.configuration.ExportExcelConfigurer;
import com.woter.fact.excel.configuration.FieldConfigurer;
import com.woter.fact.excel.core.support.HssfExcelField;
import org.apache.commons.codec.CharEncoding;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.util.LinkedHashMap;
import java.util.List;

/**
 * @author woter
 * @date 2017-9-29 下午6:30:18
 */
public abstract class AbstractExcelResolver implements ExcelResolver {

    public void exportExcel(String title, List<ExcelField<?>> fieldList, List<?> list, HttpServletResponse response, String fileName, ExportExcelConfigurer exportExcelConfigurer) throws IOException {
        response.reset();
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachment;filename=" + new String(fileName.getBytes(CharEncoding.UTF_8), CharEncoding.ISO_8859_1) + ".xls");
        OutputStream out = response.getOutputStream();
        try {
            exportExcel(title, fieldList, list, out, exportExcelConfigurer);
        } finally {
            if (out != null) {
                out.flush();
                out.close();
            }
        }
    }

    public void exportExcel(String title, List<ExcelField<?>> fieldList, List<?> list, OutputStream out, ExportExcelConfigurer exportExcelConfigurer) throws IOException {
        Preconditions.checkArgument(fieldList != null && fieldList.size() > 0, "export field cannot be empty");
        LinkedHashMap<String, HssfExcelField<?>> headerMap = Maps.newLinkedHashMapWithExpectedSize(list.size());
        for (ExcelField<?> field : fieldList) {
            if (field.getFieldConfigurer() == null) {
                field.setFieldConfigurer(new FieldConfigurer());
            }
            headerMap.put(field.getLabel(), HssfExcelField.convert(field));
        }
        if (exportExcelConfigurer == null) {
            exportExcelConfigurer = new ExportExcelConfigurer();
        }
        handleExportExcel(title, headerMap, list, out, exportExcelConfigurer);
    }

    public abstract void handleExportExcel(String title, LinkedHashMap<String, HssfExcelField<?>> headerMap, List<?> list, OutputStream out, ExportExcelConfigurer exportExcelConfigurer) throws IOException;

}
