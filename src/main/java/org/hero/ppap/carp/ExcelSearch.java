package org.hero.ppap.carp;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.hero.ppap.carp.excel.CARPCell;

import java.io.File;
import java.io.IOException;
import java.util.stream.IntStream;
import java.util.stream.Stream;

public class ExcelSearch {
    public static Stream<CARPCell> execute(File file) {
        try (Workbook workbook = WorkbookFactory.create(file, null, true)) {
            return IntStream.range(0, workbook.getNumberOfSheets())
                    .mapToObj(workbook::getSheetAt)
                    .flatMap(ExcelSearch::expandRow)
                    .flatMap(ExcelSearch::expandColumn)
                    .filter(cell -> cell.getCellType().equals(CellType.STRING))
                    .filter(cell -> cell.getStringCellValue().length() > 0)
                    .map(it -> new CARPCell(it, file));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static Stream<Row> expandRow(Sheet sheet) {
        return IntStream.range(0, sheet.getLastRowNum() + 1)
                .mapToObj(sheet::getRow);
    }

    private static Stream<Cell> expandColumn(Row row) {
        return IntStream.range(0, row.getLastCellNum())
                .mapToObj(row::getCell);
    }
}
