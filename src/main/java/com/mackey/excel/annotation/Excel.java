package com.mackey.excel.annotation;

import com.mackey.excel.NotNeedFormat;

import java.lang.annotation.*;


@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Excel {
	/**
	 * 表头
	 */
	String header();

	/**
	 * 字段
	 */
	String field();

	/**
	 * 是否需要格式化，默认否
	 */
	boolean isFormat() default false;

	/**
	 * 格式化工具类
	 */
	Class<?> format() default NotNeedFormat.class;

	/**
	 * 格式化方法
	 */
	String method() default "";

	/**
	 * 格式化方法参数类型
	 */
	Class<?>[] paramType() default {};

	/**
	 * 格式化参数
	 */
	String[] params() default {};

	/**
	 * 格式化方法默认参数类型
	 */
	Class<?>[] defaultParamType() default {};

	/**
	 * 格式化方法默认参数值
	 */
	String[] defaultValue() default {};
}
