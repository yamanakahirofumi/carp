package org.hero.ppap.carp.outputs;

import org.hero.ppap.carp.excel.CARPCell;

public class DisplayReport implements Report {
    @Override
    public void write(CARPCell cell) {
        cell.getDebug().forEach(System.out::println);
    }
}
