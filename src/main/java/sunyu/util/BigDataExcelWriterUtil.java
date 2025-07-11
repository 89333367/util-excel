package sunyu.util;

import cn.hutool.core.exceptions.ExceptionUtil;
import cn.hutool.core.io.FileUtil;
import cn.hutool.log.Log;
import cn.hutool.log.LogFactory;
import cn.hutool.poi.excel.BigExcelWriter;
import cn.hutool.poi.excel.ExcelUtil;

import java.io.*;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 大数据Excel写出工具类
 *
 * @author 孙宇
 */
public class BigDataExcelWriterUtil implements AutoCloseable {
    private final Log log = LogFactory.get();
    private final Config config;

    public static Builder builder() {
        return new Builder();
    }

    private BigDataExcelWriterUtil(Config config) {
        log.info("[构建BigDataExcelWriterUtil] 开始");
        if (config.destFile == null) {
            config.destFile = FileUtil.file("temp.xlsx");
        }
        log.info("目标路径 {}", config.destFile.getAbsolutePath());
        if (config.sheetName == null) {
            config.sheetName = "Sheet";
        }
        log.info("Sheet名称 {}", config.sheetName);
        log.info("pageSize {}", config.pageSize);
        log.info("cacheSize {}", config.cacheSize);
        log.info("临时文件路径 {}", System.getProperty("java.io.tmpdir"));
        log.info("[构建BigDataExcelWriterUtil] 结束");

        this.config = config;
    }

    private static class Config {
        //表头，key是用来找数据对应的值，value用来展示表头信息
        private final Map<String, String> headers = new LinkedHashMap<>();
        //每Sheet最大数据行数
        private int pageSize = 1000000;
        //数据缓存行数，配置越大越费内存
        private int cacheSize = 5000;
        //默认Sheet名称
        private String sheetName;
        //写出文件
        private File destFile;
        //每一行数据
        private final List<List<?>> rows = new ArrayList<>();
        //临时记录序列化文件路径
        private final List<String> tmpSerializeFilePath = new ArrayList<>();
        //写出数据行数计数器
        private int counter = 0;
    }

    public static class Builder {
        private final Config config = new Config();

        public BigDataExcelWriterUtil build() {
            return new BigDataExcelWriterUtil(config);
        }

        /**
         * 设置每Sheet最大数据行数，超过行数会自动新建Sheet，默认1000000
         *
         * @param size
         */
        public Builder pageSize(int size) {
            if (size < config.pageSize) {
                config.pageSize = size;
            }
            return this;
        }

        /**
         * 设置数据缓存行数，配置越大越费内存，默认5000
         *
         * @param size
         */
        public Builder cacheSize(int size) {
            config.cacheSize = size;
            return this;
        }

        /**
         * 设置目标文件
         *
         * @param file
         */
        public Builder destFile(File file) {
            config.destFile = file;
            return this;
        }

        /**
         * 设置目标文件
         *
         * @param file 全路径+文件名称+后缀名(/tmp/temp.xlsx)
         */
        public Builder destFile(String file) {
            config.destFile = FileUtil.file(file);
            return this;
        }

        /**
         * 设置Sheet名称
         *
         * @param name
         */
        public Builder sheetName(String name) {
            config.sheetName = name;
            return this;
        }
    }

    /**
     * 回收资源
     */
    @Override
    public void close() {
        log.info("[销毁BigDataExcelWriterUtil] 开始");
        log.info("清理临时序列化文件开始");
        config.tmpSerializeFilePath.parallelStream().forEach(filePath -> {
            try {
                //log.debug("清理 {}", filePath);
                FileUtil.del(filePath);
            } catch (Exception e) {
                log.warn("清理临时序列化文件异常 {}", ExceptionUtil.stacktraceToString(e));
            }
        });
        log.info("清理临时序列化文件完毕");
        log.info("[销毁BigDataExcelWriterUtil] 结束");
    }

    /**
     * 添加多行数据
     *
     * @param rows
     */
    synchronized public void append(List<Map<String, ?>> rows) {
        for (Map<String, ?> row : rows) {
            append(row);
        }
    }

    /**
     * 添加一行数据
     *
     * @param row
     */
    synchronized public void append(Map<String, ?> row) {
        boolean headersChanged = false;
        for (String k : row.keySet()) {
            if (!config.headers.containsKey(k)) {
                config.headers.put(k, k);

                headersChanged = true;
            }
        }
        if (headersChanged) {
            log.debug("表头有变动 {}", config.headers);
        }
        List<Object> rowData = new ArrayList<>(config.headers.size());
        for (String header : config.headers.keySet()) {
            rowData.add(row.get(header));
        }
        config.rows.add(rowData);
        if (config.rows.size() == config.cacheSize) {//到达缓存上限，序列化到磁盘
            File tempFile = FileUtil.createTempFile();
            config.tmpSerializeFilePath.add(tempFile.getAbsolutePath());
            serialize(config.rows, tempFile);
        }
    }

    /**
     * 更改表头别名，通过回调方法，更改headers里面的value即可
     *
     * @param handler
     */
    public void setHeadersAlias(java.util.function.Consumer<Map<String, String>> handler) {
        handler.accept(config.headers);
        log.debug("表头更改别名 {}", config.headers);
    }

    /**
     * 写出excel
     */
    public void write() {
        try {
            BigExcelWriter bigWriter = ExcelUtil.getBigWriter();
            bigWriter.disableDefaultStyle();//禁用样式，导出速度快
            bigWriter.setDestFile(config.destFile);
            bigWriter.renameSheet(config.sheetName);//重命名Sheet
            bigWriter.writeRow(config.headers.values());//写入表头//写入第一个Sheet的表头
            for (String filePath : config.tmpSerializeFilePath) {//从磁盘反序列化数据
                List<List<?>> dsRows = deserializer(FileUtil.file(filePath));
                if (dsRows != null) {//写出数据
                    writeRows(dsRows, bigWriter);
                }
            }
            if (!config.rows.isEmpty()) {//写出剩余数据
                writeRows(config.rows, bigWriter);
            }
            bigWriter.close();
        } catch (Exception e) {
            log.error("写出excel异常 {}", ExceptionUtil.stacktraceToString(e));
        } finally {
            log.debug("清理临时序列化文件开始");
            config.tmpSerializeFilePath.parallelStream().forEach(filePath -> {
                try {
                    //log.debug("清理 {}", filePath);
                    FileUtil.del(filePath);
                } catch (Exception e) {
                    log.warn("清理临时序列化文件异常 {}", ExceptionUtil.stacktraceToString(e));
                }
            });
            config.tmpSerializeFilePath.clear();
            log.debug("清理临时序列化文件结束");
        }
        log.debug("写出文件完毕 {}", config.destFile.getAbsolutePath());
    }

    /**
     * 写出多行数据
     *
     * @param rows
     * @param bigWriter
     * @return
     */
    private void writeRows(List<List<?>> rows, BigExcelWriter bigWriter) {
        for (List<?> row : rows) {
            if (config.counter == config.pageSize) {//如果超出限制，新建Sheet
                bigWriter.setSheet(config.sheetName + (bigWriter.getSheetCount() + 1));
                config.counter = 0;

                bigWriter.writeRow(config.headers.values());//新建一个Sheet后，写入表头
            }

            bigWriter.writeRow(row);//写出一行数据
            config.counter++;
        }
    }


    /**
     * 将数据序列化到磁盘
     *
     * @param rows
     * @param file
     */
    private void serialize(List<List<?>> rows, File file) {
        //log.debug("序列化 {}", file.getAbsolutePath());
        try (ObjectOutputStream oos = new ObjectOutputStream(new BufferedOutputStream(Files.newOutputStream(file.toPath())))) {
            oos.writeObject(rows);
            rows.clear();
        } catch (Exception e) {
            log.error("序列化文件异常 {}", ExceptionUtil.stacktraceToString(e));
        }
        //log.debug("序列化完毕 {}", file.getAbsolutePath());
    }

    /**
     * 从磁盘反序列化出数据
     *
     * @param file
     * @return
     */
    private List<List<?>> deserializer(File file) {
        //log.debug("反序列化 {}", file.getAbsolutePath());
        List<List<?>> rows = null;
        try (ObjectInputStream ois = new ObjectInputStream(new BufferedInputStream(Files.newInputStream(file.toPath())))) {
            rows = (List<List<?>>) ois.readObject();
            //log.debug("反序列化完毕 {}", file.getAbsolutePath());
        } catch (Exception e) {
            log.error("反序列化文件异常 {}", ExceptionUtil.stacktraceToString(e));
        }
        return rows;
    }

}