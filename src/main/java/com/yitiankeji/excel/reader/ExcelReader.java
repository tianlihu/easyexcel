package com.yitiankeji.excel.reader;

import com.yitiankeji.excel.constants.Constants;
import com.yitiankeji.excel.utils.ExcelType;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelReader<T> {

    private final Workbook workbook;

    public ExcelReader(BufferedInputStream input) throws IOException {
        String type = ExcelType.type(input);
        if (StringUtils.equals(Constants.XLS, type)) {
            workbook = new HSSFWorkbook(input);
        } else {
            workbook = new XSSFWorkbook(input);
        }
    }

    public List<T> doReadAll(Class<T> type) {
        return doReadAll(type, null);
    }

    public List<T> doReadAll(Class<T> type, ExcelReadListener<T> listener) {
        List<T> records = new ArrayList<>(1000);
        int sheetCount = workbook.getNumberOfSheets();
        for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
            ReadSheet<T> readSheet = sheet(sheetIndex, type, listener);
            records.addAll(readSheet.doRead());
        }
        close();
        return records;
    }

    public ReadSheet<T> sheet(int sheetNo, Class<T> type) {
        return sheet(sheetNo, type, null);
    }

    public ReadSheet<T> sheet(int sheetNo, Class<T> type, ExcelReadListener<T> listener) {
        if (sheetNo >= workbook.getNumberOfSheets()) {
            return sheet(type, null, listener);
        }
        Sheet sheet = workbook.getSheetAt(sheetNo);
        return sheet(type, sheet, listener);
    }

    public ReadSheet<T> sheet(String sheetName, Class<T> type) {
        return sheet(sheetName, type, null);
    }

    public ReadSheet<T> sheet(String sheetName, Class<T> type, ExcelReadListener<T> listener) {
        Sheet sheet = workbook.getSheet(sheetName);
        return sheet(type, sheet, listener);
    }

    private ReadSheet<T> sheet(Class<T> type, Sheet sheet, ExcelReadListener<T> listener) {
        ReadSheet<T> readSheet = new ReadSheet<>();
        readSheet.type(type);
        readSheet.sheet(sheet);
        readSheet.listener(listener);
        return readSheet;
    }

    public void close() {
        IOUtils.closeQuietly(workbook);
    }
}
