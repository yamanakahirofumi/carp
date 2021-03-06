package org.hero.ppap.carp.excel.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.hero.ppap.carp.ExcelSearch;
import org.hero.ppap.carp.excel.CARPCell;

import java.io.File;
import java.io.IOException;
import java.util.Objects;
import java.util.stream.IntStream;
import java.util.stream.Stream;

public class PoiSearch implements ExcelSearch {
    @Override
    public Stream<CARPCell> execute(File file) {
        try (Workbook workbook = WorkbookFactory.create(file, null, true)) {
            return IntStream.range(0, workbook.getNumberOfSheets())
                    .mapToObj(workbook::getSheetAt)
                    .flatMap(PoiSearch::expandRow)
                    .flatMap(PoiSearch::expandColumn)
                    .filter(cell -> cell.getCellType().equals(CellType.STRING))
                    .filter(cell -> cell.getStringCellValue().length() > 0)
                    .map(it -> new CARPCell(it, file));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static Stream<Row> expandRow(Sheet sheet) {
        return IntStream.range(0, sheet.getLastRowNum() + 1)
                .mapToObj(sheet::getRow)
                .filter(it -> !Objects.isNull(it));
    }

    private static Stream<Cell> expandColumn(Row row) {
        return IntStream.range(0, row.getLastCellNum())
                .mapToObj(row::getCell)
                .filter(it -> !Objects.isNull(it));
    }
}
