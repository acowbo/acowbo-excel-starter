package com.acowbo.excel.config;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;

/**
 * description: Excel模板配置类
 *
 * @author <a href="https://acowbo.fun">acowbo</a>
 * @version 1.0
 * @since 2025/5/19
 */
@Data
@ConfigurationProperties(prefix = "excel.template")
public class ExcelTemplateConfig {
    /**
     * 默认行高
     */
    private int defaultRowHeight = 20;

    /**
     * 默认列宽
     */
    private int defaultColumnWidth = 20;

    /**
     * 下拉选项最大行数
     */
    private int maxDropdownRows = 1000;

    /**
     * 是否启用数据验证
     */
    private boolean enableValidation = true;

    /**
     * 是否启用下拉选项
     */
    private boolean enableDropdown = true;

    /**
     * 是否启用数字格式
     */
    private boolean enableNumberFormat = true;

    /**
     * 是否启用日期格式
     */
    private boolean enableDateFormat = true;

    /**
     * 默认日期格式
     */
    private String defaultDateFormat = "yyyy-MM-dd";

    /**
     * 是否启用错误提示
     */
    private boolean enableErrorPrompt = true;

    /**
     * 错误提示样式
     */
    private String errorPromptStyle = "STOP";

    /**
     * 是否启用警告提示
     */
    private boolean enableWarningPrompt = true;

    /**
     * 警告提示样式
     */
    private String warningPromptStyle = "WARNING";
} 