package org.hero.ppap.carp.excel;

import cc.redpen.validator.ValidationError;
import org.apache.poi.ss.usermodel.Cell;

import java.io.File;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class CARPCell {
    private final Cell cell;
    private final File file;
    private List<ValidationError> message;

    public CARPCell(Cell cell, File file) {
        this.cell = cell;
        this.file = file;
    }

    public void setMessage(Stream<ValidationError> message) {
        this.message = message.collect(Collectors.toList());
    }

    public String getValue() {
        return cell.getStringCellValue();
    }

    public Stream<String> getDebug() {
        return this.message.stream().map(it ->
                this.file.getName() + "," + this.cell.getSheet().getSheetName() + "," + this.cell.getRowIndex() + ","
                        + this.cell.getColumnIndex() + "," + it.getLevel().toString() + "," + it.getMessage());
    }
}
