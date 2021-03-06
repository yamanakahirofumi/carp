package org.hero.ppap.carp;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.stream.Stream;

public class DirectorySearch {
    static Stream<File> excelSearch(File file, String exclude) {
        try {
            if (file.isDirectory()) {
                return Files.list(file.toPath())
                        .filter(it -> !it.normalize().toAbsolutePath().startsWith(exclude))
                        .map(Path::toFile)
                        .flatMap(it -> DirectorySearch.excelSearch(it, exclude));
            } else {
                return Stream.of(file)
                        .filter(it -> it.getName().endsWith(".xlsx"));
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
