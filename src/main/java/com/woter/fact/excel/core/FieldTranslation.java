package com.woter.fact.excel.core;

/**
 * @author woter
 * @date 2017-9-29 上午11:25:57
 */
public interface FieldTranslation<T> {

    Object translate(T val);

}
 