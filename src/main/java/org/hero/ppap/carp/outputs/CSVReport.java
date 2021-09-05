package org.hero.ppap.carp.outputs;

import org.hero.ppap.carp.excel.CARPCell;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Optional;
import java.util.stream.Collectors;

public class CSVReport implements Report {
    private final File file;

    CSVReport(File file) {
        this.file = file;
    }

    @Override
    public void prepare(File target) {
        if (this.file.exists()) {
            this.file.delete();
        }
    }

    @Override
    public void write(CARPCell cell) {
        try (FileWriter writer = new FileWriter(file, true)) {
            for (var line : cell.getDebug().collect(Collectors.toList())) {
                writer.write(line + System.getProperty("line.separator"));
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public Optional<Path> resultFile() {
        return Optional.of(this.file.toPath());
    }
}
