package sunyu.util;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.function.Consumer;

import org.ttzero.excel.entity.Sheet;
import org.ttzero.excel.entity.Workbook;
import org.ttzero.excel.reader.ExcelReader;
import org.ttzero.excel.reader.Row;

import cn.hutool.log.Log;
import cn.hutool.log.LogFactory;
import sunyu.util.config.DynamicColumnListMapSheet;
import sunyu.util.config.DynamicColumnWorksheetWriter;

/**
 * Excel文件工具类
 *
 * @author SunYu
 */
public class ExcelUtil implements AutoCloseable {
    private final Log log = LogFactory.get();
    private final Config config;

    public static Builder builder() {
        return new Builder();
    }

    private ExcelUtil(Config config) {
        log.info("[构建 {}] 开始", this.getClass().getSimpleName());
        // 其他初始化语句
        log.info("[构建 {}] 结束", this.getClass().getSimpleName());
        this.config = config;
    }

    private static class Config {
    }

    public static class Builder {
        private final Config config = new Config();

        public ExcelUtil build() {
            return new ExcelUtil(config);
        }
    }

    /**
     * 回收资源
     */
    @Override
    public void close() {
        log.info("[销毁 {}] 开始", this.getClass().getSimpleName());
        // 回收各种资源
        log.info("[销毁 {}] 结束", this.getClass().getSimpleName());
    }

    /**
     * 读取Excel文件
     *
     * @param path     Excel文件路径
     * @param consumer 读取器消费者
     */
    public void read(Path path, Consumer<ExcelReader> consumer) {
        try (ExcelReader reader = ExcelReader.read(path)) {
            consumer.accept(reader);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 读取Excel文件所有数据
     *
     * @param path       Excel文件路径
     * @param sheetIndex 工作表索引，从0开始
     * @param fromRowNum 表头起始行号，从1开始（包含）
     * @param toRowNum   表头结束行号，从1开始（包含）
     * @return 所有数据
     */
    public List<Map<String, Object>> read(Path path, int sheetIndex, int fromRowNum, int toRowNum) {
        List<Map<String, Object>> list = new ArrayList<>();
        read(path, reader -> reader
                .sheet(sheetIndex)
                .asFullSheet()
                .copyOnMerged() // <- 转为FullSheet并复制合并单元格
                .header(fromRowNum, toRowNum)
                .rows()
                .map(Row::toMap)
                .forEach(list::add));
        return list;
    }

    /**
     * 写入Excel文件
     *
     * @param path     Excel文件路径
     * @param fileName Excel文件名
     * @param consumer 内部workbook回调对象
     */
    public void write(Path path, String fileName, Consumer<Workbook> consumer) {
        try {
            Workbook wb = new Workbook(fileName); // 新增一个Workbook并指定名称，也就是Excel文件名
            wb.bestSpeed(); // 启用性能模式
            consumer.accept(wb);
            wb.writeTo(path);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 写入Excel文件
     *
     * @param path     Excel文件路径
     * @param fileName Excel文件名
     * @param sheet    工作表对象
     */
    public void write(Path path, String fileName, Sheet sheet) {
        write(path, fileName, wb -> {
            if (sheet instanceof DynamicColumnListMapSheet) {
                sheet.setSheetWriter(new DynamicColumnWorksheetWriter());
            }
            wb.addSheet(sheet);
        });
    }

}