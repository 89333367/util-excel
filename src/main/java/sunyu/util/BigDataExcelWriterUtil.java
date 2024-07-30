package sunyu.util;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.io.IORuntimeException;
import cn.hutool.log.Log;
import cn.hutool.log.LogFactory;
import cn.hutool.poi.excel.BigExcelWriter;
import cn.hutool.poi.excel.ExcelUtil;

import java.io.*;
import java.util.*;
import java.util.stream.Collectors;

/**
 * 大数据Excel写出工具类
 *
 * @author 孙宇
 */
public class BigDataExcelWriterUtil implements Serializable, Closeable {
    private Log log = LogFactory.get();

    /**
     * 获得工具类工厂
     *
     * @return
     */
    public static BigDataExcelWriterUtil builder() {
        return new BigDataExcelWriterUtil();
    }

    /**
     * 构建工具类
     *
     * @return
     */
    public BigDataExcelWriterUtil build() {
        if (destFile == null) {
            destFile = FileUtil.file("temp.xlsx");
        }
        if (sheetName == null) {
            sheetName = "Sheet";
        }
        return this;
    }

    /**
     * 回收资源，等待sql缓存和所有线程队列执行完毕
     */
    @Override
    public void close() {
        counter = 0;
        rows.clear();
        headers.clear();
        //log.debug("清理临时序列化文件");
        tmpSerializeFilePath.parallelStream().forEach(filePath -> {
            //log.debug("清理 {}", filePath);
            try {
                FileUtil.del(filePath);
            } catch (IORuntimeException e) {
            }
        });
        tmpSerializeFilePath.clear();
    }


    //表头，key是用来找数据对应的值，value用来展示表头信息
    private LinkedHashMap<String, String> headers = new LinkedHashMap<>();
    //每Sheet最大数据行数
    private int pageSize = 1000000;
    //数据缓存行数，配置越大越费内存
    private int cacheSize = 5000;
    //默认Sheet名称
    private String sheetName;
    //写出文件
    private File destFile;
    //每一行数据
    private List<List<Object>> rows = new ArrayList<>();
    //临时记录序列化文件路径
    private List<String> tmpSerializeFilePath = new ArrayList<>();
    //写出数据行数计数器
    private int counter = 0;


    /**
     * 表头别名回调，保留key，修改value即可
     */
    public interface HeadersAliasCallback {
        void execute(LinkedHashMap<String, String> headers);
    }

    /**
     * 设置每Sheet最大数据行数，超过行数会自动新建Sheet，默认1000000
     *
     * @param pageSize
     */
    public BigDataExcelWriterUtil setPageSize(int pageSize) {
        if (pageSize < this.pageSize) {
            this.pageSize = pageSize;
        }
        return this;
    }

    /**
     * 设置数据缓存行数，配置越大越费内存，默认5000
     *
     * @param cacheSize
     */
    public BigDataExcelWriterUtil setCacheSize(int cacheSize) {
        this.cacheSize = cacheSize;
        return this;
    }

    /**
     * 设置目标文件
     *
     * @param destFile
     */
    public BigDataExcelWriterUtil setDestFile(File destFile) {
        this.destFile = destFile;
        return this;
    }

    /**
     * 设置目标文件
     *
     * @param destFile 全路径+文件名称+后缀名(/tmp/temp.xlsx)
     */
    public BigDataExcelWriterUtil setDestFile(String destFile) {
        this.destFile = FileUtil.file(destFile);
        return this;
    }

    /**
     * 设置Sheet名称
     *
     * @param sheetName
     */
    public BigDataExcelWriterUtil setSheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    /**
     * 添加多行数据
     *
     * @param rows
     */
    public void append(List<Map<String, Object>> rows) {
        for (Map<String, Object> row : rows) {
            append(row);
        }
    }

    /**
     * 添加一行数据
     *
     * @param row
     */
    public void append(Map<String, Object> row) {
        boolean headersIsChange = false;
        for (String k : row.keySet()) {
            if (!headers.containsKey(k)) {
                headers.put(k, k);

                headersIsChange = true;
            }
        }
        if (headersIsChange) {
            log.debug("表头有变动 {}", headers);
        }
        List<Object> rowData = new ArrayList<>(headers.size());
        for (String header : headers.keySet()) {
            rowData.add(row.get(header));
        }
        rows.add(rowData);
        if (rows.size() == cacheSize) {//到达缓存上限，序列化到磁盘
            File tempFile = FileUtil.createTempFile();
            tmpSerializeFilePath.add(tempFile.getAbsolutePath());
            serialize(rows, tempFile);
        }
    }

    /**
     * 添加一行数据
     *
     * @param row
     */
    public void append(TreeMap<String, String> row) {
        Map<String, Object> m = convertToMap(row);
        append(m);
    }

    /**
     * 将TreeMap<String,String>转换成Map<String,String>
     *
     * @param treeMap
     * @return
     */
    public Map<String, Object> convertToMap(TreeMap<String, String> treeMap) {
        return treeMap.entrySet().stream().collect(Collectors.toMap(Map.Entry::getKey, e -> e.getValue()));
    }

    /**
     * 更改表头别名，通过回调方法，更改headers里面的value即可
     *
     * @param callback
     */
    public void setHeadersAlias(HeadersAliasCallback callback) {
        callback.execute(headers);
        log.debug("表头更改别名 {}", headers);
    }

    /**
     * 写出excel
     */
    public void write() {
        try {
            BigExcelWriter bigWriter = ExcelUtil.getBigWriter();
            bigWriter.disableDefaultStyle();//禁用样式，导出速度快
            bigWriter.setDestFile(destFile);
            bigWriter.renameSheet(sheetName);//重命名Sheet
            bigWriter.writeRow(headers.values());//写入表头//写入第一个Sheet的表头
            for (String filePath : tmpSerializeFilePath) {//从磁盘反序列化数据
                List<List<Object>> dsRows = deserializer(FileUtil.file(filePath));
                if (dsRows != null) {//写出数据
                    writeRows(dsRows, bigWriter);
                }
            }
            if (rows.size() > 0) {//写出剩余数据
                writeRows(rows, bigWriter);
            }
            bigWriter.close();
        } catch (Exception e) {
            log.error(e);
        } finally {
            log.debug("清理临时序列化文件");
            tmpSerializeFilePath.parallelStream().forEach(filePath -> {
                //log.debug("清理 {}", filePath);
                FileUtil.del(filePath);
            });
        }
        log.debug("写出文件完毕 {}", destFile.getAbsolutePath());
    }

    /**
     * 写出多行数据
     *
     * @param rows
     * @param bigWriter
     * @return
     */
    private void writeRows(List<List<Object>> rows, BigExcelWriter bigWriter) {
        for (List<Object> row : rows) {
            if (counter == pageSize) {//如果超出限制，新建Sheet
                bigWriter.setSheet(sheetName + (bigWriter.getSheetCount() + 1));
                counter = 0;

                bigWriter.writeRow(headers.values());//新建一个Sheet后，写入表头
            }

            bigWriter.writeRow(row);//写出一行数据
            counter++;
        }
    }


    /**
     * 将数据序列化到磁盘
     *
     * @param rows
     * @param file
     */
    private void serialize(List<List<Object>> rows, File file) {
        //log.debug("序列化 {}", file.getAbsolutePath());
        try (ObjectOutputStream oos = new ObjectOutputStream(new BufferedOutputStream(new FileOutputStream(file)))) {
            oos.writeObject(rows);
            rows.clear();
        } catch (Exception e) {
            log.error(e);
        }
        //log.debug("序列化完毕 {}", file.getAbsolutePath());
    }

    /**
     * 从磁盘反序列化出数据
     *
     * @param file
     * @return
     */
    private List<List<Object>> deserializer(File file) {
        //log.debug("反序列化 {}", file.getAbsolutePath());
        List<List<Object>> rows = null;
        try (ObjectInputStream ois = new ObjectInputStream(new BufferedInputStream(new FileInputStream(file)))) {
            rows = (List<List<Object>>) ois.readObject();
            //log.debug("反序列化完毕 {}", file.getAbsolutePath());
        } catch (Exception e) {
            log.error(e);
        }
        return rows;
    }

}