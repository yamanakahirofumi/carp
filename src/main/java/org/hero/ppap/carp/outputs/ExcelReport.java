package org.hero.ppap.carp.outputs;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.Units;
import org.hero.ppap.carp.excel.CARPCell;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Optional;

public class ExcelReport implements Report {
    private final Path result;
    private Path target;

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
    }

    @Override
    public void write(CARPCell cell) {
        File sourceFile = cell.getFile();
        File resultFile = this.resultPath(sourceFile).toFile();
        try {
            if (!resultFile.exists()) {
                Files.copy(sourceFile.toPath(), resultFile.toPath());
            }
            try (Workbook workbook = WorkbookFactory.create(new FileInputStream(resultFile))) {
                CreationHelper helper = workbook.getCreationHelper();

                Sheet sheet = workbook.getSheet(cell.getSheetName());
                Row row = sheet.getRow(cell.getRowIndex());
                Cell cell1 = row.getCell(cell.getColumnIndex());
                Comment cellComment = cell1.getCellComment();

                StringBuilder sb = new StringBuilder();
                int lineNumber = (int) cell.getDebug().count();
                for (int i = 0; i < lineNumber; i++) {
                    sb.append(cell.getPosition(i)).append(",")
                            .append(cell.getLevel(i)).append(",")
                            .append(cell.getMessage(i)).append("\n");
                }

                if (cellComment == null) {
                    var drawing = sheet.createDrawingPatriarch();
                    ClientAnchor anchor = drawing.createAnchor(
                            Units.EMU_PER_PIXEL * 10,
                            Units.EMU_PER_PIXEL * 5,
                            Units.EMU_PER_PIXEL * 10,
                            Units.EMU_PER_PIXEL * 5,
                            cell1.getColumnIndex() + 1,
                            cell1.getRowIndex(),
                            cell1.getColumnIndex() + 5,
                            cell1.getRowIndex() + lineNumber);
                    cellComment = drawing.createCellComment(anchor);
                }
                cellComment.setString(helper.createRichTextString(sb.toString()));
                cellComment.setAuthor("carp");
                cellComment.setVisible(true);
                cell1.setCellComment(cellComment);

                FileOutputStream outputStream = new FileOutputStream(resultFile);
                workbook.write(outputStream);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

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
