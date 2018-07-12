package com.woter.fact.excel.core;

import com.woter.fact.excel.configuration.FieldConfigurer;

/**
 * @author woter
 * @date 2017-9-29 上午11:25:44
 */
public class ExcelField<T> {

    private String label;

    private FieldTranslation<T> translation;

    private FieldConfigurer fieldConfigurer;

    public ExcelField(String label) {
        this.label = label;
    }

    public ExcelField(String label, String field) {
        this(label, FieldTranslationFacade.defaultFieldTranslation(field));
    }

    public ExcelField(String label, FieldTranslation<?> translation) {
        this.label = label;
        this.translation = (FieldTranslation<T>) translation;
        this.fieldConfigurer = new FieldConfigurer();
    }

    public ExcelField(String label, FieldTranslation<?> translation, FieldConfigurer fieldConfigurer) {
        this.label = label;
        this.translation = (FieldTranslation<T>) translation;
        this.fieldConfigurer = fieldConfigurer;
    }


    public static <T> ExcelField<T> of(final String label) {
        return new ExcelField<T>(label);
    }

    public static <T> ExcelField<T> of(final String label, final String field) {
        return new ExcelField<T>(label, FieldTranslationFacade.defaultFieldTranslation(field));
    }

    public static <T> ExcelField<T> of(String label, FieldTranslation<T> translation) {
        return new ExcelField<T>(label, translation);
    }

    public static <T> ExcelField<T> of(String label, FieldTranslation<T> translation, FieldConfigurer fieldConfigurer) {
        return new ExcelField<T>(label, translation, fieldConfigurer);
    }

    public static <T> ExcelField<T> ofStdDate(String label, String field) {
        return new ExcelField<T>(label, FieldTranslationFacade.dateFieldTranslation(field));
    }

    public static <T> ExcelField<T> ofStdTime(String label, String field) {
        return new ExcelField<T>(label, FieldTranslationFacade.timeFieldTranslation(field));
    }

    public Object translate(Object t) {
        if (translation != null) {
            return translation.translate((T) t);
        }
        return t;
    }

    public String getLabel() {
        return label;
    }


    public FieldTranslation<T> getTranslation() {
        return translation;
    }

    public FieldConfigurer getFieldConfigurer() {
        return fieldConfigurer;
    }

    public ExcelField<T> setFieldConfigurer(FieldConfigurer fieldConfigurer) {
        this.fieldConfigurer = fieldConfigurer;
        return this;
    }


}
