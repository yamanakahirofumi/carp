package org.hero.ppap.carp.excel;

import cc.redpen.parser.LineOffset;
import cc.redpen.validator.ValidationError;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellReference;

import java.io.File;
import java.util.List;
import java.util.Optional;
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

    private String getPositionFromLineOffset(Optional<LineOffset> lineOffset){
        return lineOffset.map(offset -> String.valueOf(offset.lineNum)).orElse("-");
    }

    public Stream<String> getDebug() {
        return this.message.stream().map(it ->
                this.file.getName() + "," +
                        this.cell.getSheet().getSheetName() + "," +
                        new CellReference(this.cell).formatAsString(false) + "," +
                        this.getPositionFromLineOffset(it.getStartPosition()) + "," +
                        it.getLevel().toString() + "," +
                        it.getMessage() + "," +
                        this.file.toPath().normalize().toAbsolutePath());
    }

    public File getFile(){
        return this.file;
    }

    public Cell getCell(){
        return this.cell;
    }

    public String getPosition(int i){
        return this.getPositionFromLineOffset(this.message.get(i).getStartPosition());
    }

    public String getLevel(int i){
        return this.message.get(i).getLevel().toString();
    }

    public String getMessage(int i){
        return this.message.get(i).getMessage();
    }
}
