package sunyu.util;

import cn.hutool.log.Log;
import cn.hutool.log.LogFactory;
import cn.hutool.poi.excel.ExcelUtil;
import sunyu.util.pojo.ExcelRow;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

/**
 * 大数据Excel读取工具类
 *
 * @author SunYu
 */
public class BigDataExcelReaderUtil implements AutoCloseable {
    private final Log log = LogFactory.get();
    private final Config config;

    public static Builder builder() {
        return new Builder();
    }

    private BigDataExcelReaderUtil(Config config) {
        log.info("[构建BigDataExcelReaderUtil] 开始");
        log.info("[构建BigDataExcelReaderUtil] 结束");
        this.config = config;
    }

    private static class Config {
        private final List<String> sheetNames = new ArrayList<>();//存储每个sheet的名称
        private final Map<Integer, List<String>> sheetHeaders = new HashMap<>();//存储每个sheet的里的标题列表
        private int rid = 0;//设置读取sheet rid，-1表示读取全部Sheet, 0表示只读取第一个Sheet
        private String filePath;//读取文件路径
        private File file;//读取文件
    }

    public static class Builder {
        private final Config config = new Config();

        public BigDataExcelReaderUtil build() {
            return new BigDataExcelReaderUtil(config);
        }

        /**
         * 设置读取sheet rid，-1表示读取全部Sheet, 0表示只读取第一个Sheet
         *
         * @param rid
         * @return
         */
        public Builder setRid(int rid) {
            config.rid = rid;
            return this;
        }

        /**
         * 设置读取文件路径
         *
         * @param filePath d:/tmp/xxx.xlsx
         * @return
         */
        public Builder setFilePath(String filePath) {
            config.filePath = filePath;
            return this;
        }

        /**
         * 设置读取文件
         *
         * @param file
         * @return
         */
        public Builder setFile(File file) {
            config.file = file;
            return this;
        }

    }

    /**
     * 回收资源
     */
    @Override
    public void close() {
        log.info("[销毁BigDataExcelReaderUtil] 开始");
        log.info("[销毁BigDataExcelReaderUtil] 结束");
    }

    /**
     * 读取Excel数据
     *
     * @param consumer 数据行处理器
     */
    public void read(Consumer<ExcelRow> consumer) {
        if (config.filePath != null) {
            ExcelUtil.readBySax(config.filePath, config.rid, (sheetIndex, rowIndex, rowCells) -> {
                extracted(consumer, sheetIndex, rowIndex, rowCells);
            });
        } else if (config.file != null) {
            ExcelUtil.readBySax(config.file, config.rid, (sheetIndex, rowIndex, rowCells) -> {
                extracted(consumer, sheetIndex, rowIndex, rowCells);
            });
        }
    }

    private void extracted(Consumer<ExcelRow> consumer, int sheetIndex, long rowIndex, List<Object> rowCells) {
        if (rowIndex == 0) {
            // 将标题行转换为String类型并去除空格
            List<String> headers = new ArrayList<>();
            for (Object cell : rowCells) {
                headers.add(cell.toString().trim());  // 标题去除左右空格
            }
            config.sheetHeaders.put(sheetIndex, headers);
        } else {
            // 获取当前sheet的标题
            List<String> headers = config.sheetHeaders.get(sheetIndex);
            if (headers != null) {
                // 将行数据转换为Map，值保持Object类型
                Map<String, Object> rowMap = new HashMap<>();
                for (int i = 0; i < headers.size(); i++) {
                    if (i < rowCells.size()) {
                        Object value = rowCells.get(i);
                        // 如果是String类型，去除左右空格
                        if (value instanceof String) {
                            value = ((String) value).trim();
                        }
                        rowMap.put(headers.get(i), value);
                    }
                }
                consumer.accept(new ExcelRow(sheetIndex, rowIndex, rowMap, rowCells));
            }
        }
    }


}
