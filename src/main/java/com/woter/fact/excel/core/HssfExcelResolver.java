package com.woter.fact.excel.core;

import com.google.common.collect.Lists;
import com.woter.fact.excel.configuration.ExportExcelConfigurer;
import com.woter.fact.excel.configuration.FieldConfigurer;
import com.woter.fact.excel.constant.FctExcelConstant;
import com.woter.fact.excel.core.support.HssfExcelField;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.io.OutputStream;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * @author woter
 * @date 2017-9-29 下午2:48:19
 */
@SuppressWarnings("all")
public class HssfExcelResolver extends AbstractExcelResolver {

    public HSSFCellStyle createTileStyle(HSSFWorkbook workbook, ExportExcelConfigurer exportExcelConfigurer) {

        HSSFCellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        titleStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        titleStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        if (exportExcelConfigurer.isTitleUseBgcolor()) {
            titleStyle.setFillForegroundColor(HSSFColor.LIGHT_ORANGE.index);
        }
        HSSFFont titleFont = workbook.createFont();
        titleFont.setFontHeightInPoints((short) 16);
        titleFont.setBoldweight((short) 700);
        titleStyle.setFont(titleFont);

        return titleStyle;
    }

    public HSSFCellStyle createHeaderStyle(HSSFWorkbook workbook, ExportExcelConfigurer exportExcelConfigurer) {
        HSSFCellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        headerStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        headerStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        headerStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        headerStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        headerStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        headerStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        if (exportExcelConfigurer.isHeaderUseBgcolor()) {
            headerStyle.setFillForegroundColor(HSSFColor.LIGHT_ORANGE.index);
        }
        HSSFFont headerFont = workbook.createFont();
        headerFont.setFontHeightInPoints((short) 10);
        headerFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        headerStyle.setFont(headerFont);
        headerStyle.setWrapText(true);

        return headerStyle;
    }

    public HSSFCellStyle createColumnCellStyle(HSSFWorkbook workbook) {
        // 单元格样式
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        HSSFFont cellFont = workbook.createFont();
        cellFont.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
        cellStyle.setFont(cellFont);
        return cellStyle;
    }

    public HSSFCellStyle createColumnCellStyle(HSSFWorkbook workbook, boolean background, boolean wrapped) {
        HSSFCellStyle cellStyle = createColumnCellStyle(workbook);
        if (background) {
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
        }
        if (wrapped) {
            cellStyle.setWrapText(wrapped);
        }
        return cellStyle;
    }

    public int calculateColumsWidth(String field, HssfExcelField<?> hssfExcelField) {
        int columsWidth = FctExcelConstant.DEFAULT_COLUMS_WIDTH;

        if (hssfExcelField.getFieldConfigurer().getColumnWidth() != FctExcelConstant.DEFAULT_COLOUMN_CHARSET_NUMBER) {
            return hssfExcelField.getFieldConfigurer().getColumnWidth() * 256;
        } else {
            if (StringUtils.isNotBlank(field) && field.getBytes().length > FctExcelConstant.DEFAULT_COLOUMN_CHARSET_NUMBER) {
                columsWidth = field.getBytes().length * 256;
                if (columsWidth > FctExcelConstant.DEFAULT_MAX_COLUMS_WIDTH) {
                    columsWidth = FctExcelConstant.DEFAULT_MAX_COLUMS_WIDTH;
                }
            }
        }
        return columsWidth;
    }

    public void setColumsWidth(HSSFSheet sheet, LinkedHashMap<String, HssfExcelField<?>> headerMap) {
        int i = 0;
        for (Map.Entry<String, HssfExcelField<?>> entry : headerMap.entrySet()) {
            sheet.setColumnWidth(i, calculateColumsWidth(entry.getKey(), entry.getValue()));
            i++;
        }
    }

    public void buildMuiltColumsStyle(HSSFWorkbook workbook, LinkedHashMap<String, HssfExcelField<?>> headerMap, ExportExcelConfigurer exportExcelConfigurer) {
        FieldConfigurer fieldConfigurer = null;
        for (HssfExcelField<?> hssfExcelField : headerMap.values()) {
            fieldConfigurer = hssfExcelField.getFieldConfigurer();
            hssfExcelField.setCellStyle(Pair.of(createColumnCellStyle(workbook, false, fieldConfigurer.isWrapped()), createColumnCellStyle(workbook, exportExcelConfigurer.isRowUseStripe(), fieldConfigurer.isWrapped())));
        }
    }

    public void handleExportExcel(String title, LinkedHashMap<String, HssfExcelField<?>> headerMap, List<?> list, OutputStream out, ExportExcelConfigurer exportExcelConfigurer) throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook();
        buildMuiltColumsStyle(workbook, headerMap, exportExcelConfigurer);
        try {
            doExportExcel(workbook, title, headerMap, list, out, exportExcelConfigurer);
        } finally {
            if (workbook != null) {
                workbook.write(out);
                out.flush();
                workbook.close();
            }
        }
    }

    public void doExportExcel(HSSFWorkbook workbook, String title, LinkedHashMap<String, HssfExcelField<?>> headerMap, List<?> list, OutputStream out, ExportExcelConfigurer exportExcelConfigurer) {

        HSSFCellStyle titleStyle = createTileStyle(workbook, exportExcelConfigurer);
        HSSFCellStyle headerStyle = createHeaderStyle(workbook, exportExcelConfigurer);
        List<String> headerFields = Lists.newArrayList(headerMap.keySet());

        HSSFSheet hssfSheet = null;
        int rowIndex = 0;
        Object cellFieldValue = null;
        HSSFRow hssfRow = null;
        HSSFCell dataFieldCell = null;
        int headerTotalRow = 0;
        Object obj = null;
        HssfExcelField<?> field = null;
        for (int i = 0; i < list.size(); i++) {
            obj = list.get(i);
            if (rowIndex % FctExcelConstant.SHEET_SIZE_LIMIT == 0) {
                rowIndex = 0;
                headerTotalRow = 0;
                hssfSheet = workbook.createSheet(); // 如果数据超过最大行,则在第二页显示
                setColumsWidth(hssfSheet, headerMap);
                if (StringUtils.isNotBlank(title)) {
                    HSSFRow titleRow = hssfSheet.createRow(rowIndex);
                    titleRow.createCell(0).setCellValue(title);
                    titleRow.getCell(0).setCellStyle(titleStyle);
                    titleRow.setHeight((short) (FctExcelConstant.DEFAULT_TITLE_ROW_HIGHT));
                    hssfSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, headerMap.size() - 1));
                    rowIndex++;
                    headerTotalRow++;
                }
                HSSFRow headerRow = hssfSheet.createRow(rowIndex);
                headerRow.setHeight((short) (rowIndex == 0 ? FctExcelConstant.DEFAULT_TITLE_ROW_HIGHT : FctExcelConstant.DEFAULT_HEADER_ROW_HIGHT));
                for (int j = 0; j < headerFields.size(); j++) {
                    headerRow.createCell(j).setCellValue(headerFields.get(j));
                    headerRow.getCell(j).setCellStyle(headerStyle);
                }
                rowIndex++;
                headerTotalRow++;
            }
            int colIndex = 0;
            hssfRow = hssfSheet.createRow(rowIndex);
            for (Map.Entry<String, HssfExcelField<?>> header : headerMap.entrySet()) {
                field = header.getValue();
                cellFieldValue = field.translate(obj);
                dataFieldCell = hssfRow.createCell(colIndex);
                dataFieldCell.setCellStyle((rowIndex - headerTotalRow) % 2 == 0 ? field.getCellStyle().getLeft() : field.getCellStyle().getRight());
                dataFieldCell.setCellValue(cellFieldValue == null ? StringUtils.EMPTY : cellFieldValue.toString());
                colIndex++;
            }
            rowIndex++;
        }
    }
}
