package sunyu.util;

import cn.hutool.log.Log;
import cn.hutool.log.LogFactory;
import cn.hutool.poi.excel.ExcelUtil;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStrings;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;
import sunyu.util.pojo.ExcelRow;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
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
        this.config = config;
        initSheetNamesAndHeaders();
        log.info("[构建BigDataExcelReaderUtil] 结束");
    }

    private static class Config {
        private final List<String> sheetNames = new ArrayList<>();
        private final Map<Integer, List<String>> sheetHeaders = new HashMap<>();
        private int rid = 0;
        private String filePath;
        private File file;
    }

    public static class Builder {
        private final Config config = new Config();

        public BigDataExcelReaderUtil build() {
            return new BigDataExcelReaderUtil(config);
        }

        public Builder setRid(int rid) {
            config.rid = rid;
            return this;
        }

        public Builder setFilePath(String filePath) {
            config.filePath = filePath;
            return this;
        }

        public Builder setFile(File file) {
            config.file = file;
            return this;
        }
    }

    @Override
    public void close() {
        log.info("[销毁BigDataExcelReaderUtil] 开始");
        log.info("[销毁BigDataExcelReaderUtil] 结束");
    }

    private void initSheetNamesAndHeaders() {
        try (InputStream inputStream = getSourceStream(); OPCPackage pkg = OPCPackage.open(inputStream)) {
            XSSFReader reader = new XSSFReader(pkg);
            XSSFReader.SheetIterator sheetIterator = (XSSFReader.SheetIterator) reader.getSheetsData();
            SharedStrings sst = reader.getSharedStringsTable();
            int currentSheetIndex = 0;
            while (sheetIterator.hasNext()) {
                InputStream sheetStream = sheetIterator.next();
                config.sheetNames.add(sheetIterator.getSheetName());
                parseSheetHeader(sheetStream, sst, currentSheetIndex);
                currentSheetIndex++;
            }
        } catch (Exception e) {
            throw new RuntimeException("初始化Sheet名称和表头信息失败", e);
        }
    }

    private InputStream getSourceStream() throws IOException {
        if (config.filePath != null) {
            return Files.newInputStream(Paths.get(config.filePath));
        } else if (config.file != null) {
            return Files.newInputStream(config.file.toPath());
        } else {
            throw new RuntimeException("请设置读取文件路径或文件");
        }
    }

    public Map<Integer, List<String>> getSheetHeaders() {
        return config.sheetHeaders;
    }

    public List<String> getSheetNames() {
        return config.sheetNames;
    }

    private void parseSheetHeader(InputStream sheetStream, SharedStrings sst, int sheetIndex) throws SAXException, IOException {
        XMLReader xmlReader = XMLReaderFactory.createXMLReader();
        HeaderSAXHandler handler = new HeaderSAXHandler(sst);
        handler.setSheetIndex(sheetIndex);
        xmlReader.setContentHandler(handler);
        xmlReader.parse(new InputSource(sheetStream));

        if (!handler.getHeaders().isEmpty()) {
            config.sheetHeaders.put(sheetIndex, handler.getHeaders());
        }
    }

    private static class HeaderSAXHandler extends DefaultHandler {
        private final SharedStrings sst;
        private int sheetIndex;
        private final List<String> headers = new ArrayList<>();
        private boolean inRow = false;
        private boolean inCell = false;
        private int rowIndex = -1;
        private final StringBuilder cellValue = new StringBuilder();
        private String sharedStringIndex = null;
        private String currentQName; // 新增变量，用于跟踪当前标签名

        public HeaderSAXHandler(SharedStrings sst) {
            this.sst = sst;
        }

        public void setSheetIndex(int sheetIndex) {
            this.sheetIndex = sheetIndex;
        }

        @Override
        public void startElement(String uri, String localName, String qName, Attributes attributes) {
            currentQName = qName; // 记录当前开始的标签名
            if ("row".equals(qName)) {
                inRow = true;
                rowIndex++;
                if (rowIndex > 0) {
                    inRow = false;
                }
            } else if (inRow && rowIndex == 0 && "c".equals(qName)) {
                inCell = true;
                cellValue.setLength(0);
                sharedStringIndex = attributes.getValue("t");
            } else if (inCell && "v".equals(qName)) {
                cellValue.setLength(0);
            }
        }

        @Override
        public void characters(char[] ch, int start, int length) {
            if (inCell && "v".equals(currentQName)) {
                cellValue.append(ch, start, length);
            }
        }

        @Override
        public void endElement(String uri, String localName, String qName) {
            currentQName = qName; // 记录当前结束的标签名
            if ("row".equals(qName)) {
                inRow = false;
            } else if (inRow && rowIndex == 0) {
                if ("c".equals(qName)) {
                    inCell = false;
                    String value = cellValue.toString().trim();
                    if ("s".equals(sharedStringIndex)) {
                        int idx = Integer.parseInt(value);
                        value = sst.getItemAt(idx).getString().trim();
                    }
                    headers.add(value);
                    sharedStringIndex = null;
                }
            }
        }

        public List<String> getHeaders() {
            return headers;
        }
    }

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
        if (rowIndex > 0) {
            List<String> headers = config.sheetHeaders.get(sheetIndex);
            if (headers != null) {
                Map<String, Object> rowMap = new HashMap<>();
                for (int i = 0; i < headers.size(); i++) {
                    if (i < rowCells.size()) {
                        Object value = rowCells.get(i);
                        if (value instanceof String) {
                            value = ((String) value).trim();
                        }
                        rowMap.put(headers.get(i), value);
                    }
                }
                consumer.accept(new ExcelRow(sheetIndex, config.sheetNames.get(sheetIndex), rowIndex, rowMap, rowCells));
            }
        }
    }
}