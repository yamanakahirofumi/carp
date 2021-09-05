package org.hero.ppap.carp.outputs;

import org.hero.ppap.carp.excel.CARPCell;

public interface Report {
    void write(CARPCell cell);
}
