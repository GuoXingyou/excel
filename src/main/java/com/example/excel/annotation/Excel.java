package com.example.excel.annotation;

import com.example.excel.enums.AlignType;
import com.example.excel.enums.OperationType;

import java.lang.annotation.*;

/**
 * @Author: Jax
 * @Email: guoxingyou@xjiye.com
 * @Date: 2017/12/14/17:39
 * @Desc:
 **/
@Target({ElementType.FIELD,ElementType.TYPE,ElementType.METHOD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Excel {

    /**
     * 导出字段名（默认调用当前字段的“get”方法，如指定导出字段为对象，请填写“对象名.对象属性”，例：“area.name”、“office.name”）
     */
    String value() default "";

    /**
     * 导出字段标题（需要添加批注请用“**”分隔，标题**批注，仅对导出模板有效）
     */
    String title();

    /**
     * 字段排序 从0开始
     */
    int index();

    /**
     * 字段类型
     */
    OperationType type() default OperationType.BOTH;

    /**
     * 导出字段对齐方式
     */
    AlignType align() default AlignType.AUTO;

    /**
     * 时间字段格式
     */
    String dateFmt() default "yyyy-MM-dd";

    /**
     * 如果是字典类型，请设置字典的type值
     */
    String dictType() default "";

    /**
     * 反射类型
     */
    Class<?> fieldType() default Class.class;

    /**
     * 字段归属组（根据分组导出导入）
     */
    int[] groups() default {};

}
