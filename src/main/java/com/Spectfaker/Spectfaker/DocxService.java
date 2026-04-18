package com.Spectfaker.Spectfaker;

import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;
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

    // Realistic handwriting variation ranges
    private static final double BASE_FONT_SIZE = 19.0;
    private static final double FONT_SIZE_VARIANCE = 0.8; // ±0.8pt per character
    private static final int SPACING_MIN = -35;           // twips/20, tighter
    private static final int SPACING_MAX = -18;           // twips/20, looser
    private static final double LINE_SPACING_MIN = 0.82;
    private static final double LINE_SPACING_MAX = 0.95;

    public void processDocxFile(MultipartFile file, HttpServletResponse response) throws IOException {
        if (file.isEmpty()) {
            throw new IllegalArgumentException("Файл не выбран");
        }

        try (InputStream inputStream = file.getInputStream()) {
            XWPFDocument document = new XWPFDocument(inputStream);
            setupDocumentDefaults(document);

            // Extract text
            StringBuilder fullTextBuilder = new StringBuilder();
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                for (XWPFRun run : paragraph.getRuns()) {
                    String text = run.getText(0);
                    if (text != null) {
                        fullTextBuilder.append(text).append(" ");
                    }
                }
                fullTextBuilder.append("\n");
            }

            String fullText = fullTextBuilder.toString().trim();
            fullText = replaceUnsupportedCharacters(fullText);

            // Remove existing paragraphs
            int paragraphCount = document.getParagraphs().size();
            for (int i = paragraphCount - 1; i >= 0; i--) {
                document.removeBodyElement(i);
            }

            String[] lines = fullText.split("\\R");
            for (String line : lines) {
                if (line.isBlank()) {
                    // Preserve empty lines as short blank paragraphs
                    XWPFParagraph blank = document.createParagraph();
                    blank.setSpacingAfter(0);
                    blank.setSpacingBefore(0);
                    continue;
                }

                String lineWithLeadingSpaces = addRandomLeadingSpaces(line);

                XWPFParagraph newParagraph = document.createParagraph();
                // Single consistent spacing setup — no double-assignment
                double lineSpacing = LINE_SPACING_MIN
                        + (LINE_SPACING_MAX - LINE_SPACING_MIN) * random.nextDouble();
                newParagraph.setSpacingBetween(lineSpacing, LineSpacingRule.AT_LEAST);
                newParagraph.setSpacingBefore(2);
                newParagraph.setSpacingAfter(0);

                addLineToParagraph(newParagraph, lineWithLeadingSpaces);
            }

            setupResponse(response, file.getOriginalFilename());
            writeDocumentToResponse(document, response);
        }
    }

    private String replaceUnsupportedCharacters(String text) {
        return text
                .replace("–", "-")
                .replace("—", "-")
                .replace("‒", "-")
                .replace("−", "-");
    }

    // 0–2 spaces: subtle indent, not a tab stop every line
    private String addRandomLeadingSpaces(String text) {
        int spacesCount = random.nextInt(3); // 0, 1, or 2
        return " ".repeat(spacesCount) + text;
    }

    private void addLineToParagraph(XWPFParagraph paragraph, String line) {
        String[] words = line.split("\\s+");
        for (int i = 0; i < words.length; i++) {
            String word = words[i];
            for (char c : word.toCharArray()) {
                editSymbol(paragraph, c);
            }

            if (i < words.length - 1) {
                // 1–2 spaces between words — enough to look hand-spaced, not typed
                int spaceCount = random.nextInt(2) + 1;
                for (int j = 0; j < spaceCount; j++) {
                    editSymbol(paragraph, ' ');
                }
            }
        }
    }

    private void editSymbol(XWPFParagraph paragraph, char c) {
        XWPFRun newSymbol = paragraph.createRun();
        newSymbol.setText(String.valueOf(c), 0);
        newSymbol.setCharacterSpacing(getRandomCharacterSpacing());

        // Slight font-size jitter — real handwriting is never perfectly uniform
        double sizeDelta = (random.nextDouble() * 2 - 1) * FONT_SIZE_VARIANCE;
        newSymbol.setFontSize(BASE_FONT_SIZE + sizeDelta);

        List<String> rusFontArray = List.of(
                "PavelFont1 Regular", "PavelFont2 Regular",
                "PavelFont3 Regular", "PavelFont4 Regular",
                "PavelFont5 Regular", "PavelFont6 Regular",
                "PavelFont7 Regular", "PavelFont8 Regular"
        );

        List<String> engFontArray = List.of(
                "PavelFontENG1 Regular", "PavelFontENG2 Regular"
        );

        String commonSymbols = "!\"'*()[]{},.+-'/0123456789:;<>?@";

        if ((c >= 'А' && c <= 'я') || c == 'ё' || c == 'Ё') {
            newSymbol.setFontFamily(rusFontArray.get(random.nextInt(rusFontArray.size())));
        } else if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z')) {
            newSymbol.setFontFamily(engFontArray.get(random.nextInt(engFontArray.size())));
        } else if (commonSymbols.indexOf(c) != -1) {
            // Common symbols: pick from either pool randomly
            boolean useRus = random.nextBoolean();
            String font = useRus
                    ? rusFontArray.get(random.nextInt(rusFontArray.size()))
                    : engFontArray.get(random.nextInt(engFontArray.size()));
            newSymbol.setFontFamily(font);
            newSymbol.setFontSize(17.5 + (random.nextDouble() * 2 - 1) * 0.5);
        } else {
            // Spaces, unknown chars
            newSymbol.setFontFamily(engFontArray.get(random.nextInt(engFontArray.size())));
        }
    }

    private int getRandomCharacterSpacing() {
        // Real variation in twips/20: tighter on some chars, looser on others
        return SPACING_MIN + random.nextInt(SPACING_MAX - SPACING_MIN + 1);
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
        return (int) (sm * 567);
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