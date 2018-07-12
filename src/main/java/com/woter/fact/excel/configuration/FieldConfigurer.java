package com.woter.fact.excel.configuration;

import com.woter.fact.excel.constant.FctExcelConstant;

/**
 * @author woter
 * @date 2017-10-1 上午8:33:45
 */
public class FieldConfigurer {

    private int characterNumber; // 当前列展示字符个数

    private boolean wrapped; // 是否自动换行

    public FieldConfigurer() {
        this.wrapped = true;
        this.characterNumber = FctExcelConstant.DEFAULT_COLOUMN_CHARSET_NUMBER;
    }

    public FieldConfigurer(int characterNumber) {
        this.characterNumber = characterNumber;
        this.wrapped = true;
    }

    public FieldConfigurer(boolean wrapped) {
        this.wrapped = wrapped;
        this.characterNumber = FctExcelConstant.DEFAULT_COLOUMN_CHARSET_NUMBER;
    }

    public FieldConfigurer(int columnWidth, boolean wrapped) {
        this.characterNumber = columnWidth;
        this.wrapped = wrapped;
    }

    public int getColumnWidth() {
        return characterNumber;
    }

    public void setColumnWidth(int columnWidth) {
        this.characterNumber = columnWidth;
    }

    public boolean isWrapped() {
        return wrapped;
    }

    public void setWrapped(boolean wrapped) {
        this.wrapped = wrapped;
    }

}
