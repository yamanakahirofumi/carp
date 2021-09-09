package org.hero.ppap.carp.outputs;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.hero.ppap.carp.excel.CARPCell;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

public class ExcelReport implements Report {
    private final Path result;
    private long size;
    private Path target;
    private Map<String, Map<String, List<CARPCell>>> cellMap;

    ExcelReport(File file) {
        this.result = file.toPath().toAbsolutePath();
    }

    @Override
    public void prepare(File target) {
        if (!this.result.toFile().exists()) {
            try {
                Files.createDirectory(this.result);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        } else {
            throw new RuntimeException("結果格納先が既に存在しています。");
        }
        this.target = target.toPath().toAbsolutePath();
        this.size = 0;
        this.cellMap = new HashMap<>();
    }

    @Override
    public void write(CARPCell cell) {
        final int batchSize = 250;
        String absolutePath = cell.getFile().getAbsolutePath();
        Map<String, List<CARPCell>> bookMap = this.cellMap.getOrDefault(absolutePath, new HashMap<>());
        List<CARPCell> cellList = bookMap.getOrDefault(cell.getSheetName(), new ArrayList<>());
        cellList.add(cell);
        bookMap.put(cell.getSheetName(), cellList);
        this.cellMap.put(absolutePath, bookMap);
        this.size++;
        if (this.size % batchSize == 0) {
            writeExec(cellMap);
            this.cellMap = new HashMap<>();
        }
    }

    private void writeExec(Map<String, Map<String, List<CARPCell>>> map) {
        try {
            for (var bookSet : map.entrySet()) {
                File sourceFile = new File(bookSet.getKey());
                File resultFile = this.resultPath(sourceFile).toFile();
                if (!resultFile.exists()) {
                    Files.copy(sourceFile.toPath(), resultFile.toPath());
                }
                try (Workbook workbook = WorkbookFactory.create(new FileInputStream(resultFile))) {
                    CreationHelper helper = workbook.getCreationHelper();

                    Font redFont;
                    XSSFFont xssfFont = new XSSFFont();
                    short red = IndexedColors.RED.getIndex();
                    redFont = workbook.findFont(true, red, xssfFont.getFontHeight(), xssfFont.getFontName(),
                            false, false, xssfFont.getTypeOffset(), xssfFont.getUnderline());
                    if (redFont == null) {
                        redFont = workbook.createFont();
                        redFont.setColor(red);
                        redFont.setBold(true);
                    }
                    for (var sheetSet : bookSet.getValue().entrySet()) {
                        Sheet sheet = workbook.getSheet(sheetSet.getKey());
                        for (var cell : sheetSet.getValue()) {
                            Row row = sheet.getRow(cell.getRowIndex());
                            Cell cell1 = row.getCell(cell.getColumnIndex());
                            Comment cellComment = cell1.getCellComment();
                            int defaultFontIndex = cell1.getCellStyle().getFontIndex();
                            RichTextString richTextString = helper.createRichTextString(cell1.getStringCellValue());
                            richTextString.applyFont(0, cell1.getStringCellValue().length(), (short) defaultFontIndex);

                            StringBuilder sb = new StringBuilder();
                            int lineNumber = (int) cell.getDebug().count();
                            for (int i = 0; i < lineNumber; i++) {
                                String startPosition = cell.getStartPosition(i);
                                sb.append(startPosition).append(",")
                                        .append(cell.getLevel(i)).append(",")
                                        .append(cell.getMessage(i)).append("\n");
                                if (startPosition.chars().allMatch(Character::isDigit)) {
                                    richTextString.applyFont(Integer.parseInt(startPosition),
                                            Integer.parseInt(cell.getEndPosition(i)), redFont);
                                }
                            }
                            cell1.setCellValue(richTextString);

                            if (cellComment == null) {
                                var drawing = sheet.createDrawingPatriarch();
                                ClientAnchor anchor = drawing.createAnchor(
                                        Units.EMU_PER_PIXEL * 10,
                                        Units.EMU_PER_PIXEL * 5,
                                        Units.EMU_PER_PIXEL * 10,
                                        Units.EMU_PER_PIXEL * 5,
                                        cell1.getColumnIndex() + 1,
                                        cell1.getRowIndex(),
                                        cell1.getColumnIndex() + 6,
                                        cell1.getRowIndex() + lineNumber);
                                cellComment = drawing.createCellComment(anchor);
                            }
                            cellComment.setString(helper.createRichTextString(sb.toString()));
                            cellComment.setAuthor("carp");
                            cellComment.setVisible(true);
                            cell1.setCellComment(cellComment);
                        }
                    }
                    FileOutputStream outputStream = new FileOutputStream(resultFile);
                    workbook.write(outputStream);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Override
    public void postWrite() {
        writeExec(this.cellMap);
    }

    @Override
    public Optional<Path> resultFile() {
        return Optional.of(this.result);
    }

    private Path relativePath(File f) {
        return target.relativize(f.toPath().toAbsolutePath());
    }

    private Path resultPath(File f) {
        return this.result.resolve(this.relativePath(f));
    }
}
