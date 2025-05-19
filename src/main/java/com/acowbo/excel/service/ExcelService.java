package com.acowbo.excel.service;


import com.acowbo.excel.model.ExcelTemplate;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.util.List;
import java.util.Map;

/**
 * description: Excel服务接口
 *
 * @author <a href="https://acowbo.fun">acowbo</a>
 * @version 1.0
 * @since 2025/5/19
 */
public interface ExcelService {
    /**
     * 下载Excel模板
     *
     * @param response HTTP响应对象
     * @param template Excel模板
     * @param fileName 文件名
     */
    void downloadTemplate(HttpServletResponse response, ExcelTemplate template, String fileName);

    /**
     * 导入Excel数据
     *
     * @param file Excel文件
     * @param template Excel模板
     * @param importHandler 导入处理函数
     * @return 导入结果
     */
    String importData(MultipartFile file, ExcelTemplate template, ImportHandler importHandler);

    /**
     * 导入处理函数接口
     */
    @FunctionalInterface
    interface ImportHandler {
        /**
         * 处理导入数据
         *
         * @param dataList 数据列表
         * @return 处理结果
         */
        String handle(List<Map<String, Object>> dataList);
    }
} 