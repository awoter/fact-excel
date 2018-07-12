package com.woter.fact.excel.configuration;

/**
 * @author woter
 * @date 2017-9-29 下午8:09:50
 */
public class ExportExcelConfigurer {

    private boolean rowUseStripe; // 每隔行之间是否使用条纹背景

    private boolean titleUseBgcolor; // 标题是否使用背景颜色

    private boolean headerUseBgcolor; // 表头是否使用背景颜色

    public ExportExcelConfigurer() {
        this.rowUseStripe = true;
        this.titleUseBgcolor = true;
        this.headerUseBgcolor = true;

    }

    public ExportExcelConfigurer(boolean rowUseStripe) {
        this.rowUseStripe = rowUseStripe;
        this.titleUseBgcolor = true;
        this.headerUseBgcolor = true;
    }

    public ExportExcelConfigurer(boolean rowUseStripe, boolean titleUseBgcolor, boolean headerUseBgcolor) {
        this.rowUseStripe = rowUseStripe;
        this.titleUseBgcolor = titleUseBgcolor;
        this.headerUseBgcolor = headerUseBgcolor;
    }

    public boolean isRowUseStripe() {
        return rowUseStripe;
    }

    public void setRowUseStripe(boolean rowUseStripe) {
        this.rowUseStripe = rowUseStripe;
    }

    public boolean isTitleUseBgcolor() {
        return titleUseBgcolor;
    }

    public void setTitleUseBgcolor(boolean titleUseBgcolor) {
        this.titleUseBgcolor = titleUseBgcolor;
    }

    public boolean isHeaderUseBgcolor() {
        return headerUseBgcolor;
    }

    public void setHeaderUseBgcolor(boolean headerUseBgcolor) {
        this.headerUseBgcolor = headerUseBgcolor;
    }

}
