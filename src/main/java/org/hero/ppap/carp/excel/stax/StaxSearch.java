package org.hero.ppap.carp.excel.stax;

import org.hero.ppap.carp.ExcelSearch;
import org.hero.ppap.carp.excel.CARPCell;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;
import java.io.File;
import java.io.IOException;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;
import java.util.stream.Stream;

public class StaxSearch implements ExcelSearch {
    @Override
    public Stream<CARPCell> execute(File file) {
        XMLStreamReader reader = null;
        XMLInputFactory factory = XMLInputFactory.newInstance();
        try {
            FileSystem fs = FileSystems.newFileSystem(file.toPath(), null);
            reader = factory.createXMLStreamReader(Files.newInputStream(fs.getPath("xl/sharedStrings.xml")));
            List<String> stringList = prepare(reader);
            reader.close();
            reader = factory.createXMLStreamReader(Files.newInputStream(fs.getPath("xl/workbook.xml")));
            Map<Integer, String> sheetNameMap = getSheetInfo(reader);
            return Files.list(fs.getPath("xl/worksheets"))
                    .filter(it -> !Files.isDirectory(it))
                    .flatMap(it -> getCellInfo(fs, it, stringList, file, sheetNameMap))
                    .map(CARPCell::new);
        } catch (XMLStreamException | IOException e) {
            throw new RuntimeException(e);
        } finally {
            try {
                if (reader != null) {
                    reader.close();
                }
            } catch (XMLStreamException ignored) {
            }
        }
    }

    private List<String> prepare(XMLStreamReader reader) throws XMLStreamException {
        boolean si = false;
        List<String> result = new ArrayList<>();
        StringBuilder sb = new StringBuilder();
        while (reader.hasNext()) {
            int eventType = reader.next();
            switch (eventType) {
                case XMLStreamConstants.START_ELEMENT:
                    if (reader.getName().getLocalPart().equals("si")) {
                        si = true;
                        sb = new StringBuilder();
                    }
                    break;
                case XMLStreamConstants.END_ELEMENT:
                    if (reader.getName().getLocalPart().equals("si")) {
                        si = false;
                        result.add(sb.toString());
                    }
                    break;
                case XMLStreamConstants.CHARACTERS:
                    if (si) {
                        sb.append(reader.getText());
                    }
                    break;
                default:
            }
        }
        return result;
    }

    private Map<Integer, String> getSheetInfo(XMLStreamReader reader) throws XMLStreamException {
        Map<Integer, String> result = new HashMap<>();
        while (reader.hasNext()) {
            int eventType = reader.next();
            if (eventType == XMLStreamConstants.START_ELEMENT) {
                if (reader.getName().getLocalPart().equals("sheet")) {
                    Integer num = Integer.valueOf(reader.getAttributeValue("", "sheetId"));
                    String name = reader.getAttributeValue("", "name");
                    result.put(num, name);
                }
            }
        }
        return result;
    }

    private Stream<CellPosition> getCellInfo(FileSystem fs, Path path, List<String> stringList, File file, Map<Integer, String> sheetMap) {
        XMLStreamReader reader = null;
        XMLInputFactory factory = XMLInputFactory.newInstance();
        boolean cell = false;
        boolean share = false;
        boolean value = false;
        String cellPosition = "";
        List<CellPosition> result = new LinkedList<>();
        Pattern p = Pattern.compile("([A-Z]+)([0-9]+)");
        int sheetId = Integer.parseInt(path.getFileName().toString().split("\\.")[0].split("t")[1]);
        try {
            reader = factory.createXMLStreamReader(Files.newInputStream(fs.getPath(path.toString())));
            while (reader.hasNext()) {
                int eventType = reader.next();
                switch (eventType) {
                    case XMLStreamConstants.START_ELEMENT:
                        String startLocalPart = reader.getName().getLocalPart();
                        if (startLocalPart.equals("c")) {
                            cellPosition = reader.getAttributeValue("", "r");
                            cell = true;
                            if (reader.getAttributeValue("", "t").equals("s")) {
                                share = true;
                            }
                        } else if (startLocalPart.equals("v") && share) {
                            value = true;
                        }
                        break;
                    case XMLStreamConstants.END_ELEMENT:
                        String endLocalPart = reader.getName().getLocalPart();
                        if (endLocalPart.equals("c")) {
                            cell = false;
                            share = false;
                            value = false;
                        }
                        break;
                    case XMLStreamConstants.CHARACTERS:
                        if (value) {
                            int num = Integer.parseInt(reader.getText());
                            String[] position = p.matcher(cellPosition).replaceAll("$1,$2").split(",");
                            result.add(new CellPosition(file, sheetMap.get(sheetId), sheetId,
                                    this.getNumber(position[0]) - 1, Integer.parseInt(position[1]) - 1, cellPosition,
                                    stringList.get(num)));
                        }
                        break;
                    default:
                }
            }
        } catch (XMLStreamException | IOException e) {
            e.printStackTrace();
        } finally {
            try {
                reader.close();
            } catch (XMLStreamException ignored) {
            }
        }
        return result.stream();
    }

    private int getNumber(String row) {
        int result = 0;
        for (char c : row.toCharArray()) {
            result = result * 26 + (c - 'A' + 1);
        }
        return result;
    }
}
