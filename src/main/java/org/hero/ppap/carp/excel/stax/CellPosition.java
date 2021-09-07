package org.hero.ppap.carp.excel.stax;

import java.io.File;

public class CellPosition {
    private final File file;
    private final String sheetName;
    private final int sheetId;
    private final int rowNum;
    private final int columnNum;
    private final String cellPosition;
    private final String sentence;

    public CellPosition(File file, String sheetName, int sheetId, int rowNum, int columnNum, String cellPosition, String sentence) {
        this.file = file;
        this.sheetName = sheetName;
        this.sheetId = sheetId;
        this.rowNum = rowNum;
        this.columnNum = columnNum;
        this.cellPosition = cellPosition;
        this.sentence = sentence;
    }

    public File file() {
        return this.file;
    }

    public String sheetName() {
        return this.sheetName;
    }

    public int rowNum() {
        return this.rowNum;
    }

    public int columnNum() {
        return this.columnNum;
    }

    public String cellPosition() {
        return this.cellPosition;
    }

    public String sentence() {
        return this.sentence;
    }
}
