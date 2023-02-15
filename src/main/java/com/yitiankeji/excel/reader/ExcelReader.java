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

public class ExcelReader {

    private final Workbook workbook;

    public ExcelReader(BufferedInputStream input) throws IOException {
        String type = ExcelType.type(input);
        if (StringUtils.equals(Constants.XLS, type)) {
            workbook = new HSSFWorkbook(input);
        } else {
            workbook = new XSSFWorkbook(input);
        }
    }

    public <T> List<T> doReadAll(Class<T> type) {
        return doReadAll(type, null);
    }

    public <T> List<T> doReadAll(Class<T> type, ExcelReadListener<T> listener) {
        List<T> records = new ArrayList<>(1000);
        int sheetCount = workbook.getNumberOfSheets();
        for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
            ReadSheet<T> readSheet = sheet(sheetIndex, type, listener);
            if (readSheet == null) {
                continue;
            }
            records.addAll(readSheet.doRead());
        }
        close();
        return records;
    }

    public <T> ReadSheet<T> sheet(int sheetIndex, Class<T> type) {
        return sheet(sheetIndex, type, null);
    }

    public <T> ReadSheet<T> sheet(int sheetIndex, Class<T> type, ExcelReadListener<T> listener) {
        if (sheetIndex >= workbook.getNumberOfSheets()) {
            return sheet(type, null, listener);
        }
        if (workbook.isSheetHidden(sheetIndex)) {
            return null;
        }
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        return sheet(type, sheet, listener);
    }

    public <T> ReadSheet<T> sheet(String sheetName, Class<T> type) {
        return sheet(sheetName, type, null);
    }

    public <T> ReadSheet<T> sheet(String sheetName, Class<T> type, ExcelReadListener<T> listener) {
        Sheet sheet = workbook.getSheet(sheetName);
        int sheetIndex = workbook.getSheetIndex(sheet);
        if (workbook.isSheetHidden(sheetIndex)) {
            return null;
        }
        return sheet(type, sheet, listener);
    }

    private <T> ReadSheet<T> sheet(Class<T> type, Sheet sheet, ExcelReadListener<T> listener) {
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
