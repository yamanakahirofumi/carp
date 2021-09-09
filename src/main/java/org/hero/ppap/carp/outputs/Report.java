package org.hero.ppap.carp.outputs;

import org.hero.ppap.carp.excel.CARPCell;

import java.io.File;
import java.nio.file.Path;
import java.util.Optional;

public interface Report {
    void prepare(File target);

    void write(CARPCell cell);

    void postWrite();

    Optional<Path> resultFile();
}
