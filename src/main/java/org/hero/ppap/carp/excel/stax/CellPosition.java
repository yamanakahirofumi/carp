package org.hero.ppap.carp.excel.stax;

import java.io.File;

public class CellPosition {
    private final File file;
    private final String sheetName;
    private final int sheetId;
    private final String cellPosition;
    private final String sentence;

    public CellPosition(File file, String sheetName, int sheetId, String cellPosition, String sentence) {
        this.file = file;
        this.sheetName = sheetName;
        this.sheetId = sheetId;
        this.cellPosition = cellPosition;
        this.sentence = sentence;
    }
}
