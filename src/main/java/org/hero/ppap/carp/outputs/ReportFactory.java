package org.hero.ppap.carp.outputs;

import java.io.File;
import java.util.Objects;

public class ReportFactory {
    public static Report create(ResultType resultType, File resultFile){
        File file;
        file = Objects.requireNonNullElseGet(resultFile, () -> new File("result" + resultType.getExt()));
        // TODO: Switch-Expressionに変更したい
        switch (resultType){
            case CSV:
                return new CSVReport(file);
            case None:
                return new NoReport();
            case EXCEL:
                return new ExcelReport(file);
            case DISPLAY:
            default:
                return new DisplayReport();
        }
    }
}
