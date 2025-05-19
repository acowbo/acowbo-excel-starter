package com.acowbo.excel.config;

import com.acowbo.excel.service.ExcelService;

import com.acowbo.excel.service.impl.ExcelServiceImpl;
import org.springframework.boot.autoconfigure.condition.ConditionalOnMissingBean;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

/**
 * description: Excel自动配置类
 *
 * @author <a href="https://acowbo.fun">acowbo</a>
 * @version 1.0
 * @since 2025/5/19
 */
@Configuration
@EnableConfigurationProperties(ExcelTemplateConfig.class)
public class ExcelAutoConfiguration {

    @Bean
    @ConditionalOnMissingBean
    public ExcelService excelService(ExcelTemplateConfig config) {
        return new ExcelServiceImpl(config);
    }
} 