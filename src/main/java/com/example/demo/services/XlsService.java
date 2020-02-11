package com.example.demo.services;

import com.github.jferard.fastods.TableRowImpl;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Random;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.IntStream;

import static com.example.demo.enums.Extension.*;

public class XlsService {

    private static Random random = new Random();
    private static String underline = "_";
    private static String point = ".";

    protected static void exportExcel(List<String> stringHeaderList, List<List<String>> stringContentList, String dirName, String fileName, String titleTag) throws IOException {
        builderSheet(stringHeaderList, stringContentList, dirName, fileName, titleTag);
    }

    private static void builderSheet(List<String> stringHeaderList, List<List<String>> stringContentList, String dirName, String fileName, String titleTag) throws IOException {
        AtomicInteger rownum = new AtomicInteger();

        final Workbook workbook = new HSSFWorkbook();
        final Sheet sheet = workbook.createSheet(titleTag);

        AtomicInteger cellnum = new AtomicInteger();
        final Row rw = sheet.createRow(0);

        stringHeaderList.forEach(header -> {
            rw.createCell(cellnum.getAndIncrement()).setCellValue(header);
        });

        rownum.set(1);
        stringContentList.forEach(strings -> {
            Row row = sheet.createRow(rownum.getAndIncrement());
            strings.forEach(content -> {
                Row[] rows = {row};
                AtomicInteger cellNumber = new AtomicInteger();
                IntStream.range(0, strings.size())
                        .forEach(index -> {
                            AtomicInteger idx = new AtomicInteger();
                            rows[idx.getAndIncrement()].createCell(cellNumber.getAndIncrement()).setCellValue(strings.get(index));
                        });
            });
        });

            FileOutputStream out =
                    new FileOutputStream(new File(dirName, fileName.concat(underline.concat(String.valueOf(random).concat(point)).concat(String.valueOf(xls)))));
            workbook.write(out);
            out.close();
    }

}
