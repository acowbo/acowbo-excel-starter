package com.acowbo.excel.service.impl;

import cn.hutool.core.map.MapUtil;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import com.acowbo.excel.config.ExcelTemplateConfig;
import com.acowbo.excel.model.ExcelTemplate;
import com.acowbo.excel.service.ExcelService;
import com.acowbo.excel.util.ExcelValidationUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.net.URLEncoder;
import java.util.*;
import java.util.stream.Collectors;

/**
 * description: Excel服务实现类
 *
 * @author <a href="https://acowbo.fun">acowbo</a>
 * @version 1.0
 * @since 2025/5/19
 */
@Slf4j
@Service
public class ExcelServiceImpl implements ExcelService {

    private final ExcelTemplateConfig config;

    public ExcelServiceImpl(ExcelTemplateConfig config) {
        this.config = config;
    }

    @Override
    public void downloadTemplate(HttpServletResponse response, ExcelTemplate template, String fileName) {
        ExcelWriter writer = new ExcelWriter();
        writer.setDefaultRowHeight(config.getDefaultRowHeight());

        try {
            Map<String, Object> templateData = template.getTemplateData();
            // 获取表头和列索引
            LinkedHashMap<String, String> headers = (LinkedHashMap<String, String>) templateData.get("headers");
            Map<String, Integer> columnIndices = MapUtil.get(templateData, "columnIndices", Map.class);

            // 设置表头别名
            headers.forEach(writer::addHeaderAlias);

            // 设置列宽
            for (int i = 0; i < headers.size(); i++) {
                writer.setColumnWidth(i, config.getDefaultColumnWidth());
            }

            // 添加下拉选项
            if (config.isEnableDropdown()) {
                addDropdownOptions(writer, templateData, columnIndices);
            }

            // 添加数据验证
            if (config.isEnableValidation()) {
                addDataValidation(writer, templateData, columnIndices);
            }

            // 写入表头
            writer.setOnlyAlias(true);
            writer.writeHeadRow(new ArrayList<>(headers.values()));

            // 设置响应头
            response.setContentType("application/octet-stream;charset=utf-8");
            response.setHeader("Content-Disposition", "attachment;fileName=" + URLEncoder.encode(fileName, "UTF-8"));
            writer.flush(response.getOutputStream(), true);
        } catch (IOException e) {
            log.error("模板导出异常：", e);
        } finally {
            writer.close();
        }
    }

    @Override
    public String importData(MultipartFile file, ExcelTemplate template, ImportHandler importHandler) {
        if (file.isEmpty()) {
            return "请选择导入文件!";
        }

        try {
            // 读取Excel文件
            ExcelReader reader = ExcelUtil.getReader(file.getInputStream());

            // 设置别名，将Excel列名映射到字段名
            LinkedHashMap<String, String> headers = (LinkedHashMap<String, String>) template.getTemplateData().get("headers");
            headers.forEach((key, value) -> reader.addHeaderAlias(value, key));

            // 读取数据，跳过标题行
            List<Map<String, Object>> rawDataList = reader.readAll();

            // 处理数据
            List<Map<String, Object>> processedDataList = rawDataList.stream()
                    .map(this::processRowData)
                    .collect(Collectors.toList());

            // 调用导入处理函数
            return importHandler.handle(processedDataList);
        } catch (Exception e) {
            log.error("导入数据异常", e);
            return "导入失败：" + e.getMessage();
        }
    }

    /**
     * 处理行数据
     */
    private Map<String, Object> processRowData(Map<String, Object> row) {
        Map<String, Object> newRow = new HashMap<>();
        row.forEach((key, value) -> {
            if (value instanceof Date) {
                newRow.put(key, new java.text.SimpleDateFormat(config.getDefaultDateFormat()).format((Date) value));
            } else if (value instanceof String) {
                newRow.put(key, cleanString((String) value));
            } else {
                newRow.put(key, value);
            }
        });
        return newRow;
    }

    /**
     * 清理字符串
     */
    private String cleanString(String input) {
        return input.replace("\n", "").replace("\r", "").trim();
    }

    /**
     * 添加下拉选项
     */
    private void addDropdownOptions(ExcelWriter writer, Map<String, Object> templateData, Map<String, Integer> columnIndices) {
        Sheet sheet = writer.getSheet();
        templateData.forEach((key, value) -> {
            if (key.endsWith("List") && value instanceof List) {
                List<Map<String, Object>> dataList = (List<Map<String, Object>>) value;
                String fieldName = key.substring(0, key.length() - 4);
                Integer colIndex = columnIndices.get(fieldName);
                if (colIndex != null && !dataList.isEmpty()) {
                    String[] options = dataList.stream()
                            .map(item -> MapUtil.getStr(item, "name"))
                            .filter(Objects::nonNull)
                            .toArray(String[]::new);

                    CellRangeAddressList rangeList = new CellRangeAddressList(1, config.getMaxDropdownRows(), colIndex, colIndex);
                    writer.addSelect(rangeList, options);
                }
            }
        });
    }

    /**
     * 添加数据验证
     */
    private void addDataValidation(ExcelWriter writer, Map<String, Object> templateData, Map<String, Integer> columnIndices) {
        Sheet sheet = writer.getSheet();
        Workbook workbook = writer.getWorkbook();
        DataValidationHelper helper = sheet.getDataValidationHelper();

        templateData.forEach((key, value) -> {
            if (key.endsWith("Validation") && value instanceof Map) {
                String fieldName = key.replace("Validation", ""); // 去掉 "Validation" 后缀
                Integer colIndex = columnIndices.get(fieldName);
                if (colIndex != null) {
                    ExcelValidationUtil.addValidation(helper, sheet, colIndex, (Map<String, Object>) value, config);
                }
            }
        });
    }
} 