package cn.xuyang520.util.excel;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

/**
 * excel工具
 *
 * @author xy
 */
public class ExcelUtil {

    private static final char CHAR_MIN = 'a';
    private static final char CHAR_MAX = 'z';

    /**
     * 输出到流
     * @param sheetName sheet页名
     * @param os 输出流
     * @param sources 源数据，需要导出的字段加{@link ExcelColumn}注解
     * @param <T> 类型
     */
    public static <T> void outputToStream(String sheetName, OutputStream os, List<T> sources) {

        if (sources == null || sources.size() == 0 || sources.get(0) == null) {
            throw new RuntimeException("请确保sources有值且第一个元素不为null");
        }
        if (os == null) {
            throw new RuntimeException("请确保os不为null");
        }
        if (sheetName == null) {
            sheetName = "sheet1";
        }

        // 获取表格列定义
        ExcelColumnDefine[] excelColumnDefines = getExcelColumnDefine(sources.get(0).getClass());

        // 创建对象
        XSSFWorkbook sheets = new XSSFWorkbook();
        XSSFSheet sheet = sheets.createSheet(sheetName);

        // 标题行
        List<String> titles = Arrays.stream(excelColumnDefines).map(ExcelColumnDefine::getTitle).collect(Collectors.toList());
        setCellValues(sheet.createRow(0), titles);

        // 内容
        String[][] content = getContext(sources, excelColumnDefines);
        for (int i = 0; i < content.length; i++) {
            setCellValues(sheet.createRow(i + 1), content[i]);
        }

        // 自适应列宽
        for (int i = 0; i < excelColumnDefines.length; i++) {
            sheet.autoSizeColumn(i);
        }

        try {
            sheets.write(os);
        } catch (IOException e) {
            throw new RuntimeException("导出失败", e);
        }
    }

    private static void setCellValues(XSSFRow row, List<String> values) {
        for (int i = 0; i < values.size(); i++) {
            row.createCell(i).setCellValue(values.get(i));
        }
    }
    private static void setCellValues(XSSFRow row, String[] values) {
        for (int i = 0; i < values.length; i++) {
            row.createCell(i).setCellValue(values[i]);
        }
    }

    private static String initialCapital(String s) {
        char[] chars = s.toCharArray();
        if (chars[0] >= CHAR_MIN && chars[0] <= CHAR_MAX) {
            chars[0] = (char)(chars[0] - 32);
        }
        return new String(chars);
    }

    @Getter
    @Setter
    @ToString
    @AllArgsConstructor
    private static class ExcelColumnDefine {
        private ExcelColumn annotation;
        private Field field;
        private Method m;
        private String title;
        private int index;
    }

    private static ExcelColumnDefine[] getExcelColumnDefine(Class<?> clazz) {
        List<ExcelColumnDefine> excelColumnDefinesList = new ArrayList<>();
        for (Field field : clazz.getDeclaredFields()) {
            ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
            if (annotation == null) {
                continue;
            }
            String name = field.getName();
            try {
                Method m = clazz.getMethod("get" + initialCapital(name));
                if (m != null) {
                    excelColumnDefinesList.add(new ExcelColumnDefine(annotation, field, m, annotation.title(), annotation.index()));
                }
            } catch (NoSuchMethodException e) {
                e.printStackTrace();
            }
        }

        ExcelColumnDefine[] excelColumnDefines = new ExcelColumnDefine[excelColumnDefinesList.size()];
        for (ExcelColumnDefine define : excelColumnDefinesList) {
            ExcelColumn annotation = define.getAnnotation();
            if (annotation.index() < 0 || annotation.index() >= excelColumnDefines.length) {
                throw new RuntimeException("字段[" + define.getField().getName() + "]索引[" + annotation.index() + "]无效，有效范围：0-" + (excelColumnDefines.length - 1));
            }
            if (excelColumnDefines[annotation.index()] != null) {
                throw new RuntimeException("字段["+ define.getField().getName() + "]的索引与字段[" + excelColumnDefines[annotation.index()].getField().getName() + "]重复，重复索引[" + annotation.index() + "]");
            }
            excelColumnDefines[annotation.index()] = define;
        }

        return excelColumnDefines;
    }

    private static String[][] getContext(List<?> sources, ExcelColumnDefine[] excelColumnDefines) {
        String[][] content = new String[sources.size()][excelColumnDefines.length];

        for (int x = 0; x < sources.size(); x++) {
            Object t = sources.get(x);
            if (t == null) {
                continue;
            }
            for (int y = 0; y < excelColumnDefines.length; y++) {
                ExcelColumnDefine define = excelColumnDefines[y];
                if (define == null) {
                    continue;
                }
                try {
                    Object value = define.getM().invoke(t);
                    content[x][y] = value == null ? "" : value.toString();
                } catch (IllegalAccessException | InvocationTargetException e) {
                    e.printStackTrace();
                }
            }
        }
        return content;
    }

}
