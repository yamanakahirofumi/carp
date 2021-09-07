package org.hero.ppap.carp.outputs;

public enum ResultType {
    CSV(".csv"),
    DISPLAY(""),
    EXCEL("_xlsx"),
    None("");

    private final String ext;

    ResultType(String ext) {
        this.ext = ext;
    }

    public String getExt() {
        return this.ext;
    }
}
