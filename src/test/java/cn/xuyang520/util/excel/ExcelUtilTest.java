package cn.xuyang520.util.excel;

import lombok.Getter;
import lombok.Setter;
import org.junit.Test;

import java.io.ByteArrayOutputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.Arrays;

import static org.junit.Assert.*;

public class ExcelUtilTest {
    @Getter
    @Setter
    public class ExcelColumnTest {

        @ExcelColumn(index = 0, title = "test1")
        private String aa;
        @ExcelColumn(index = 1, title = "test2")
        private String bb;
        @ExcelColumn(index = 2, title = "test3adsadsadsadas")
        private String cc;
        @ExcelColumn(index = 3, title = "asd")
        private String dd;
        @ExcelColumn(index = 4, title = "dwdwd")
        private String ee;
    }

    @Test
    public void outputToStream() throws FileNotFoundException {
        ExcelColumnTest excelColumnTest1 = new ExcelColumnTest();
        excelColumnTest1.setAa("asdasadsadadsadsadasasdsadasdasad");
        excelColumnTest1.setCc("dwadawdwadw");
        ExcelColumnTest excelColumnTest2 = new ExcelColumnTest();
        excelColumnTest2.setAa("qweqwe");
        excelColumnTest2.setBb("asdasdsadasdd");


        ByteArrayOutputStream os = new ByteArrayOutputStream();

        ExcelUtil.outputToStream("订单列表", new FileOutputStream("/Users/xuyang/aaa.xlsx"), Arrays.asList(excelColumnTest1, excelColumnTest2));
    }
}