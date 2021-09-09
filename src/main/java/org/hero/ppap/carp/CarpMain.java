package org.hero.ppap.carp;

import org.hero.ppap.carp.excel.CARPCell;
import org.hero.ppap.carp.excel.poi.PoiSearch;
import org.hero.ppap.carp.excel.stax.StaxSearch;
import org.hero.ppap.carp.outputs.Report;
import org.hero.ppap.carp.outputs.ReportFactory;
import org.hero.ppap.carp.outputs.ResultType;
import picocli.CommandLine;
import picocli.CommandLine.Command;
import picocli.CommandLine.Option;
import picocli.CommandLine.Parameters;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.concurrent.Callable;


@Command(name = "carp", mixinStandardHelpOptions = true, version = "CARP 0.1.6",
        description = "Personal Poi And redPen check tool.(Check excel files of A directory by RedPen.)")
class CarpMain implements Callable<Integer> {

    @Parameters(index = "0", description = "A directory")
    private File file;

    @Option(names = {"-c", "--conf"}, description = "RedPen Config File Path")
    private File configFile;

    @Option(names = {"-r", "--result_format"}, description = "Result Format. Choose ${COMPLETION-CANDIDATES}")
    private final ResultType type = ResultType.DISPLAY;

    @Option(names = {"-o", "--output_file"}, description = "Output File Path")
    private File resultFile;

    @Option(names = {"--no-poi"}, description = "It is Experimental feature.")
    private boolean noPoi;

    @Override
    public Integer call() throws Exception {
        if (!file.exists()) {
            System.out.println("Directory is Not Found.");
            return 0;
        }
        if (file.isFile()) {
            System.out.println("Please choice a Directory.");
            return 0;
        }

        RedPenManager redPenManager;
        if (configFile != null) {
            redPenManager = new RedPenManager(configFile);
        } else {
            redPenManager = new RedPenManager("/redpen-conf-ja.xml");
        }
        Report report = ReportFactory.create(type, resultFile);
        report.prepare(file);
        ExcelSearch search;
        if (noPoi) {
            search = new StaxSearch();
        } else {
            search = new PoiSearch();
        }
        Files.list(file.toPath())
                .map(Path::toFile)
                .flatMap(it -> DirectorySearch.excelSearch(it, report.resultFile().map(Path::toString).orElse("-")))
                .flatMap(search::execute)
                .peek(it -> it.setMessage(redPenManager.validate(it.getValue())))
                .filter(CARPCell::isMessage)
                .forEach(report::write);
        report.postWrite();
        return 0;
    }

    public static void main(String... args) {
        int exitCode = new CommandLine(new CarpMain()).setCaseInsensitiveEnumValuesAllowed(true).execute(args);
        System.exit(exitCode);
    }
}
