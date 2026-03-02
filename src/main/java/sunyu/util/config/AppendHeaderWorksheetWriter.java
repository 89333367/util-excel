package sunyu.util.config;

import org.ttzero.excel.entity.e7.XMLWorksheetWriter;
import org.ttzero.excel.util.ExtBufferedWriter;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.channels.SeekableByteChannel;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;

public class AppendHeaderWorksheetWriter extends XMLWorksheetWriter {
    // 记录body的位置
    long position = 0L;

    @Override
    protected void beforeSheetData(boolean nonHeader) throws IOException {
        super.beforeSheetData(nonHeader);
        bw.flush(); // 刷新流

        // 获取当前position，从position以后开始写实际的body
        position = Files.size(workSheetPath.resolve(sheet.getFileName()));
    }

    @Override
    public void close() throws IOException {
        super.close();

        XMLWorksheetWriter _writer = new XMLWorksheetWriter(sheet) {
            private boolean hasMedia() {
                return false;
            }
        };
        Class<XMLWorksheetWriter> clazz = XMLWorksheetWriter.class;
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        try {
            Field totalRowsField = clazz.getDeclaredField("totalRows");
            totalRowsField.setAccessible(true);
            totalRowsField.set(_writer, totalRows);
            Field startRowField = clazz.getDeclaredField("startRow");
            startRowField.setAccessible(true);
            startRowField.set(_writer, startRow - columns[0].subColumnSize());
            Field startHeaderRowField = clazz.getDeclaredField("startHeaderRow");
            startHeaderRowField.setAccessible(true);
            startHeaderRowField.set(_writer, startHeaderRow);
            Field includeAutoWidthField = clazz.getDeclaredField("includeAutoWidth");
            includeAutoWidthField.setAccessible(true);
            includeAutoWidthField.set(_writer, includeAutoWidth);
            Field stylesField = clazz.getDeclaredField("styles");
            stylesField.setAccessible(true);
            stylesField.set(_writer, styles);
            Field bwField = clazz.getDeclaredField("bw");
            bwField.setAccessible(true);
            BufferedWriter bw = new ExtBufferedWriter(new OutputStreamWriter(baos, StandardCharsets.UTF_8));
            bwField.set(_writer, bw);
            // 重写col
            Method writeBeforeMethod = clazz.getDeclaredMethod("writeBefore");
            writeBeforeMethod.setAccessible(true);
            writeBeforeMethod.invoke(_writer);

            // 重写表头
            Method beforeSheetDataMethod = clazz.getDeclaredMethod("beforeSheetData", boolean.class);
            beforeSheetDataMethod.setAccessible(true);
            beforeSheetDataMethod.invoke(_writer, sheet.getNonHeader() == 1);

            bw.close();
        } catch (NoSuchFieldException | IllegalAccessException | InvocationTargetException | NoSuchMethodException |
                 IOException e) {
            e.printStackTrace();
        }

        try {
            Path currentPath = workSheetPath.resolve(sheet.getFileName());
            String fileName = currentPath.getFileName().toString();
            // 创建临时文件
            Path tmpPath = Files.createFile(workSheetPath.resolve(fileName + "_cp"));
            // 将新表头复制到临时文件中
            try (SeekableByteChannel channel = Files.newByteChannel(tmpPath, StandardOpenOption.WRITE, StandardOpenOption.READ)) {
                ByteBuffer buffer = ByteBuffer.wrap(baos.toByteArray());
                buffer.order(ByteOrder.LITTLE_ENDIAN);
                channel.write(buffer);
            }

            // 将Body复制到临时文件中
            try (InputStream is = Files.newInputStream(currentPath);
                 OutputStream os = Files.newOutputStream(tmpPath, StandardOpenOption.APPEND)) {
                is.skip(position); // <- 跳到body处

                byte[] bytes = new byte[4096];
                int n;
                while ((n = is.read(bytes)) > 0) {
                    os.write(bytes, 0, n);
                }
            }

            // 替换现有文件
            Files.delete(currentPath);
            tmpPath.toFile().renameTo(currentPath.toFile());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}