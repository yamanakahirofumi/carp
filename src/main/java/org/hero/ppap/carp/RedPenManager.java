package org.hero.ppap.carp;

import cc.redpen.RedPen;
import cc.redpen.RedPenException;
import cc.redpen.model.Document;
import cc.redpen.tokenizer.RedPenTokenizer;
import cc.redpen.validator.ValidationError;

import java.io.File;
import java.util.stream.Stream;

public class RedPenManager {
    private final RedPen redPen;

    public RedPenManager(File file) throws RedPenException {
        this.redPen = new RedPen(file);
    }

    public RedPenManager(String path) throws RedPenException {
        this.redPen = new RedPen(path);
    }

    private RedPenTokenizer getTokenizer() {
        return this.redPen.getConfiguration().getTokenizer();
    }

    public Stream<ValidationError> validate(String sentence) {
        Document.DocumentBuilder builder = Document.builder(this.getTokenizer());
        builder.addSentence(sentence, 0);
        return this.redPen.validate(builder.build()).stream();
    }
}