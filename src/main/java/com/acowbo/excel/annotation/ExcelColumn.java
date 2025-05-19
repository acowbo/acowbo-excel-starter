package com.acowbo.excel.annotation;

import java.lang.annotation.*;

/**
 * description: 用于标记实体类中的字段与Excel列的映射关系
 *
 * @author <a href="https://acowbo.fun">acowbo</a>
 * @version 1.0
 * @since 2025/5/19
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelColumn {
    /**
     * 列名
     */
    String value();

    /**
     * 列索引
     */
    int index() default -1;

    /**
     * 是否必填
     */
    boolean required() default false;

    /**
     * 数据格式校验类型
     */
    ValidationType validationType() default ValidationType.NONE;

    /**
     * 最小长度（用于字符串类型）
     */
    int minLength() default -1;

    /**
     * 最大长度（用于字符串类型）
     */
    int maxLength() default -1;

    /**
     * 下拉选项的key（用于从字典中获取下拉选项）
     */
    String dropdownKey() default "";

    /**
     * 数据格式校验类型枚举
     */
    enum ValidationType {
        /**
         * 无校验
         */
        NONE("none"),
        /**
         * 手机号校验
         */
        PHONE("phone"),
        /**
         * 身份证号校验
         */
        ID_CARD("idCard"),
        /**
         * 银行卡号校验
         */
        BANK_CARD("bankCard"),
        /**
         * 邮箱校验
         */
        EMAIL("email"),
        /**
         * 数字校验
         */
        NUMBER("number");

        private final String value;

        ValidationType(String value) {
            this.value = value;
        }

        public String getValue() {
            return value;
        }
    }
} 