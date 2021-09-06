package org.hero.ppap.carp;

import org.hero.ppap.carp.excel.CARPCell;

import java.io.File;
import java.util.stream.Stream;

public interface ExcelSearch {
    Stream<CARPCell> execute(File file);
}
