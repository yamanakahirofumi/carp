package org.hero.ppap.carp.outputs;

import org.hero.ppap.carp.excel.CARPCell;

import java.io.File;
import java.nio.file.Path;
import java.util.Optional;

public class NoReport implements Report {
    @Override
    public void prepare(File target) {
    }

    @Override
    public void write(CARPCell cell) {
    }

    @Override
    public void postWrite(){
    }

    @Override
    public Optional<Path> resultFile() {
        return Optional.empty();
    }
}
