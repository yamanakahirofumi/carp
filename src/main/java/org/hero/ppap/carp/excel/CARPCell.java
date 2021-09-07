package org.hero.ppap.carp.excel;

import cc.redpen.validator.ValidationError;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellReference;
import org.hero.ppap.carp.excel.stax.CellPosition;

import java.io.File;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class CARPCell {
    private final CellPosition cellPosition;
    private List<ValidationError> message;

    public CARPCell(Cell cell, File file) {
        this.cellPosition = new CellPosition(file, cell.getSheet().getSheetName(),
                cell.getSheet().getWorkbook().getSheetIndex(cell.getSheet()), cell.getRowIndex(), cell.getColumnIndex(),
                new CellReference(cell).formatAsString(false), cell.getStringCellValue());
    }

    public CARPCell(CellPosition cellPosition) {
        this.cellPosition = cellPosition;
    }

    public void setMessage(Stream<ValidationError> message) {
        this.message = message.collect(Collectors.toList());
    }

    public String getValue() {
        return this.cellPosition.sentence();
    }

    private String getPositionFromLineOffset(ValidationError validationError) {
        return validationError.getStartPosition().map(offset -> String.valueOf(offset.lineNum)).orElse("-");
    }

    public Stream<String> getDebug() {
        return this.message.stream().map(it ->
                this.cellPosition.file().getName() + "," +
                        this.getSheetName() + "," +
                        this.cellPosition.cellPosition() + "," +
                        this.getPositionFromLineOffset(it) + "," +
                        it.getLevel().toString() + "," +
                        it.getMessage() + "," +
                        this.cellPosition.file().toPath().normalize().toAbsolutePath());
    }

    public boolean isMessage() {
        return this.message.size() > 0;
    }

    public File getFile() {
        return this.cellPosition.file();
    }

    public String getSheetName() {
        return this.cellPosition.sheetName();
    }

    public int getRowIndex() {
        return this.cellPosition.rowNum();
    }

    public int getColumnIndex() {
        return this.cellPosition.columnNum();
    }

    public String getPosition(int i) {
        return this.getPositionFromLineOffset(this.message.get(i));
    }

    public String getLevel(int i) {
        return this.message.get(i).getLevel().toString();
    }

    public String getMessage(int i) {
        return this.message.get(i).getMessage();
    }
}
