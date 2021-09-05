package org.hero.ppap.carp.outputs;

import org.hero.ppap.carp.excel.CARPCell;

import java.io.File;
import java.nio.file.Path;
import java.util.Optional;

public class DisplayReport implements Report {
    @Override
    public void prepare(File target) {
    }

    @Override
    public void write(CARPCell cell) {
        cell.getDebug().forEach(System.out::println);
    }

    @Override
    public Optional<Path> resultFile() {
        return Optional.empty();
    }
}
