package com.Spectfaker.Spectfaker;

import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.List;
import java.util.Random;

@Service
public class DocxService {

    private final Random random = new Random();

    // Размер шрифта — волна
    private double currentFontSize = 19.0;
    private int charsSinceLastShift = 0;
    private int nextShiftAfter = 0;

    // Межбуквенный интервал — волна
    private int currentSpacing = -28;
    private int charsInSpacingWave = 0;
    private int spacingWaveLength = 0;

    // Вертикальное смещение букв — волна
    private int currentVertShift = 0;
    private int charsInVertWave = 0;
    private int vertWaveLength = 0;

    public void processDocxFile(MultipartFile file, HttpServletResponse response) throws IOException {
        if (file.isEmpty()) {
            throw new IllegalArgumentException("Файл не выбран");
        }

        try (InputStream inputStream = file.getInputStream()) {
            XWPFDocument document = new XWPFDocument(inputStream);
            setupDocumentDefaults(document);

            StringBuilder fullTextBuilder = new StringBuilder();
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                for (XWPFRun run : paragraph.getRuns()) {
                    String text = run.getText(0);
                    if (text != null) fullTextBuilder.append(text).append(" ");
                }
                fullTextBuilder.append("\n");
            }

            String fullText = replaceUnsupportedCharacters(fullTextBuilder.toString().trim());

            int paragraphCount = document.getParagraphs().size();
            for (int i = paragraphCount - 1; i >= 0; i--) {
                document.removeBodyElement(i);
            }

            String[] lines = fullText.split("\\R");
            boolean firstLineOfParagraph = true;

            for (String line : lines) {
                if (line.isBlank()) {
                    XWPFParagraph blank = document.createParagraph();
                    applyLineSpacing(blank);
                    blank.setSpacingBefore(0);
                    blank.setSpacingAfter(0);
                    firstLineOfParagraph = true;
                    continue;
                }

                String lineWithIndent = firstLineOfParagraph
                        ? " ".repeat(random.nextInt(3) + 3) + line
                        : " ".repeat(random.nextInt(2)) + line;
                firstLineOfParagraph = false;

                XWPFParagraph newParagraph = document.createParagraph();
                applyLineSpacing(newParagraph);
                newParagraph.setSpacingBefore(0);
                newParagraph.setSpacingAfter(0);

                addLineToParagraph(newParagraph, lineWithIndent);
            }

            setupResponse(response, file.getOriginalFilename());
            writeDocumentToResponse(document, response);
        }
    }

    /**
     * Абсолютный межстрочный интервал, привязанный к сетке тетради.
     * 0.5 см = 28.35пт = 567 твипов. EXACT не даёт Word растягивать строку.
     */
    private void applyLineSpacing(XWPFParagraph paragraph) {
        int lineSpacingTwips = 567 + random.nextInt(3) - 1; // 566–568
        CTSpacing spacing = paragraph.getCTP().getPPr() != null
                ? (paragraph.getCTP().getPPr().isSetSpacing()
                ? paragraph.getCTP().getPPr().getSpacing()
                : paragraph.getCTP().getPPr().addNewSpacing())
                : paragraph.getCTP().addNewPPr().addNewSpacing();
        spacing.setLine(BigInteger.valueOf(lineSpacingTwips));
        spacing.setLineRule(STLineSpacingRule.EXACT);
    }

    private void addLineToParagraph(XWPFParagraph paragraph, String line) {
        String[] words = line.split("\\s+");
        resetWaves();

        for (int i = 0; i < words.length; i++) {
            for (char c : words[i].toCharArray()) {
                editSymbol(paragraph, c);
            }
            if (i < words.length - 1) {
                int spaceCount = random.nextInt(2) + 2;
                for (int j = 0; j < spaceCount; j++) {
                    editSymbol(paragraph, ' ');
                }
            }
        }
    }

    private void editSymbol(XWPFParagraph paragraph, char c) {
        XWPFRun run = paragraph.createRun();
        run.setText(String.valueOf(c), 0);
        run.setCharacterSpacing(getWaveSpacing());
        run.setTextPosition(getWaveVerticalShift()); // вертикальное смещение

        List<String> rusFonts = List.of(
                "PavelFont1 Regular", "PavelFont2 Regular",
                "PavelFont3 Regular", "PavelFont4 Regular",
                "PavelFont5 Regular", "PavelFont6 Regular",
                "PavelFont7 Regular", "PavelFont8 Regular"
        );
        List<String> engFonts = List.of(
                "PavelFontENG1 Regular", "PavelFontENG2 Regular"
        );
        String commonSymbols = "!\"'*()[]{},.+-'/0123456789:;<>?@";

        if ((c >= 'А' && c <= 'я') || c == 'ё' || c == 'Ё') {
            run.setFontFamily(rusFonts.get(random.nextInt(rusFonts.size())));
            run.setFontSize(getWaveFontSize());
        } else if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z')) {
            run.setFontFamily(engFonts.get(random.nextInt(engFonts.size())));
            run.setFontSize(getWaveFontSize());
        } else if (commonSymbols.indexOf(c) != -1) {
            run.setFontFamily(random.nextBoolean()
                    ? rusFonts.get(random.nextInt(rusFonts.size()))
                    : engFonts.get(random.nextInt(engFonts.size())));
            run.setFontSize(getWaveFontSize() - 1.5);
        } else {
            run.setFontFamily(engFonts.get(random.nextInt(engFonts.size())));
            run.setFontSize(getWaveFontSize());
        }
    }

    /**
     * Размер шрифта меняется плавно каждые 4–8 символов.
     * Высокая инерция (0.8) — волна медленная, без резких скачков.
     */
    private double getWaveFontSize() {
        if (charsSinceLastShift >= nextShiftAfter) {
            double target = 18.3 + random.nextDouble() * 1.9; // 18.3–20.2
            currentFontSize = currentFontSize * 0.8 + target * 0.2;
            charsSinceLastShift = 0;
            nextShiftAfter = random.nextInt(5) + 4;
        }
        charsSinceLastShift++;
        return currentFontSize + (random.nextDouble() * 0.3 - 0.15); // микро-дрожание
    }

    /**
     * Межбуквенный интервал: тесный, плавно гуляет.
     * Диапазон -38..-20 твипов/20.
     */
    private int getWaveSpacing() {
        if (charsInSpacingWave >= spacingWaveLength) {
            int target = -38 + random.nextInt(18); // -38 до -20
            currentSpacing = (int)(currentSpacing * 0.7 + target * 0.3);
            charsInSpacingWave = 0;
            spacingWaveLength = random.nextInt(6) + 3;
        }
        charsInSpacingWave++;
        return currentSpacing + random.nextInt(5) - 2; // микро-дрожание ±2
    }

    /**
     * Вертикальное смещение букв в полупунктах (+вверх, -вниз).
     * Диапазон -3..+3, волна медленная — имитирует неровность строки.
     */
    private int getWaveVerticalShift() {
        if (charsInVertWave >= vertWaveLength) {
            int target = random.nextInt(7) - 3; // -3 до +3
            currentVertShift = (int)(currentVertShift * 0.75 + target * 0.25);
            charsInVertWave = 0;
            vertWaveLength = random.nextInt(8) + 5; // каждые 5–12 символов
        }
        charsInVertWave++;
        return currentVertShift;
    }

    private void resetWaves() {
        charsSinceLastShift = 0;
        nextShiftAfter = 0;
        charsInSpacingWave = 0;
        spacingWaveLength = 0;
        charsInVertWave = 0;
        vertWaveLength = 0;
        currentFontSize = 19.0;
        currentSpacing = -28;
        currentVertShift = 0;
    }

    private String replaceUnsupportedCharacters(String text) {
        return text.replace("–", "-").replace("—", "-")
                .replace("‒", "-").replace("−", "-");
    }

    private void setupDocumentDefaults(XWPFDocument document) {
        CTSectPr sectPr = document.getDocument().getBody().isSetSectPr()
                ? document.getDocument().getBody().getSectPr()
                : document.getDocument().getBody().addNewSectPr();

        CTPageSz pageSize = sectPr.isSetPgSz() ? sectPr.getPgSz() : sectPr.addNewPgSz();
        pageSize.setW(BigInteger.valueOf(smToTWIPs(17)));
        pageSize.setH(BigInteger.valueOf(smToTWIPs(20.3)));
        pageSize.setOrient(STPageOrientation.PORTRAIT);

        CTPageMar pageMar = sectPr.isSetPgMar() ? sectPr.getPgMar() : sectPr.addNewPgMar();
        pageMar.setLeft(BigInteger.valueOf(smToTWIPs(2)));
        pageMar.setRight(BigInteger.valueOf(smToTWIPs(2)));
        pageMar.setTop(BigInteger.valueOf(smToTWIPs(0.7)));
        pageMar.setBottom(BigInteger.valueOf(smToTWIPs(0.6)));
    }

    public int smToTWIPs(double sm) {
        return (int)(sm * 567);
    }

    private void setupResponse(HttpServletResponse response, String originalFilename) {
        response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        response.setHeader("Content-Disposition", "attachment; filename=\"" + originalFilename + "\"");
    }

    private void writeDocumentToResponse(XWPFDocument document, HttpServletResponse response) throws IOException {
        try (OutputStream outputStream = response.getOutputStream()) {
            document.write(outputStream);
        }
    }
}