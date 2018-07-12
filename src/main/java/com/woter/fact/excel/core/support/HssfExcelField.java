package com.woter.fact.excel.core.support;

import com.woter.fact.excel.configuration.FieldConfigurer;
import com.woter.fact.excel.core.ExcelField;
import com.woter.fact.excel.core.FieldTranslation;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;

/**
 * @author woter
 * @date 2017-10-1 上午10:29:07
 */
public class HssfExcelField<T> extends ExcelField<T> {

    private Pair<HSSFCellStyle, HSSFCellStyle> cellStyle;  //left:sigle  right:double

    public HssfExcelField(String label) {
        super(label);
    }

    public HssfExcelField(String label, FieldTranslation<?> translation, FieldConfigurer fieldConfigurer) {
        super(label, translation, fieldConfigurer);
    }

    public Pair<HSSFCellStyle, HSSFCellStyle> getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(Pair<HSSFCellStyle, HSSFCellStyle> cellStyle) {
        this.cellStyle = cellStyle;
    }

    public static <T> HssfExcelField<T> convert(ExcelField<T> excelField) {
        if (excelField == null)
            return null;
        return new HssfExcelField<T>(excelField.getLabel(), excelField.getTranslation(), excelField.getFieldConfigurer());
    }
}
