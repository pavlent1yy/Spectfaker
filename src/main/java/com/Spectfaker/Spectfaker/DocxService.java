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
                    if (text != null) {
                        fullTextBuilder.append(text).append(" ");
                    }
                }
                fullTextBuilder.append("\n");
            }

            String fullText = fullTextBuilder.toString().trim();
            fullText = replaceUnsupportedCharacters(fullText); // заменяем тире на дефисы

            int paragraphCount = document.getParagraphs().size();
            for (int i = paragraphCount - 1; i >= 0; i--) {
                document.removeBodyElement(i);
            }

            String[] lines = fullText.split("\\R");
            for (String line : lines) {
                // Добавляем случайные пробелы в начало строки (от 1 до 4)
                String lineWithLeadingSpaces = addRandomLeadingSpaces(line);

                // Создаем мини-параграф для каждой строки
                XWPFParagraph newParagraph = document.createParagraph();
                newParagraph.setSpacingBetween(0.85 + (0.05 * random.nextDouble()));
                newParagraph.setSpacingBefore(0);
                newParagraph.setSpacingAfter(0);

                addLineToParagraph(newParagraph, lineWithLeadingSpaces);
            }

            setupResponse(response, file.getOriginalFilename());
            writeDocumentToResponse(document, response);
        }
    }

    // Заменяем тире на дефисы
    private String replaceUnsupportedCharacters(String text) {
        return text.replace("–", "-").replace("—", "-").replace("‒", "-").replace("−", "-");
    }

    // Пробелы в начале строки: от 1 до 4
    private String addRandomLeadingSpaces(String text) {
        int spacesCount = random.nextInt(4) + 1; // 1 - 4 пробела
        return " ".repeat(spacesCount) + text;
    }

    // Добавляем слова с пробелами (1-4 пробела между словами)
    private void addLineToParagraph(XWPFParagraph paragraph, String line) {
        String[] words = line.split("\\s+");
        paragraph.setSpacingBefore(2);
        paragraph.setSpacingBetween(0.65, LineSpacingRule.AT_LEAST);
        for (int i = 0; i < words.length; i++) {
            String word = words[i];
            for (char c : word.toCharArray()) {
                editSymbol(paragraph, c);
            }

            // Добавляем пробелы после слова (от 2 до 4)
            int spaceCount = random.nextInt(4) + 2; // 1 - 4 пробела
            for (int j = 0; j < spaceCount; j++) {
                editSymbol(paragraph, ' ');
            }
        }
    }

    private void editSymbol(XWPFParagraph paragraph, char c) {
        XWPFRun newSymbol = paragraph.createRun();
        newSymbol.setText(String.valueOf(c), 0);
        newSymbol.setCharacterSpacing(getRandomCharacterSpacing());
        newSymbol.setFontSize(19);

        List<String> rusFontArray = List.of(
                "PavelFont1 Regular", "PavelFont2 Regular",
                "PavelFont3 Regular", "PavelFont4 Regular",
                "PavelFont5 Regular", "PavelFont6 Regular",
                "PavelFont7 Regular", "PavelFont8 Regular"
        );

        List<String> engFontArray = List.of(
                "PavelFontENG1 Regular", "PavelFontENG2 Regular"
        );

        // Символы, которые встречаются в обоих шрифтах
        String commonSymbols = "!\"'*()[]{},.+-'/0123456789:;<>?@";

        // Выбор шрифта
        if ((c >= 'А' && c <= 'я') || c == 'ё' || c == 'Ё') {
            // Русские буквы — берем только из русского списка
            newSymbol.setFontFamily(rusFontArray.get(random.nextInt(rusFontArray.size())));
        } else if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z')) {
            // Английские буквы — берем только из английского списка
            newSymbol.setFontFamily(engFontArray.get(random.nextInt(engFontArray.size())));
        } else if (commonSymbols.indexOf(c) != -1) {
            // Общие символы — случайно выбираем русский или английский шрифт
            List<String> combinedFonts = List.of(
                    rusFontArray.get(random.nextInt(rusFontArray.size())),
                    engFontArray.get(random.nextInt(engFontArray.size()))
            );
            newSymbol.setFontFamily(combinedFonts.get(random.nextInt(combinedFonts.size())));
            newSymbol.setFontSize(17.5);
        } else {
            // Если это какой-то другой символ, например пробел, то оставляем стандартный шрифт
            newSymbol.setFontFamily(engFontArray.get(random.nextInt(engFontArray.size())));
        }
    }

    private void setupDocumentDefaults(XWPFDocument document) {
        CTSectPr sectPr = document.getDocument().getBody().isSetSectPr()
                ? document.getDocument().getBody().getSectPr()
                : document.getDocument().getBody().addNewSectPr();

        // Размер страницы: 170мм x 203мм
        CTPageSz pageSize = sectPr.isSetPgSz() ? sectPr.getPgSz() : sectPr.addNewPgSz();
        pageSize.setW(BigInteger.valueOf(smToTWIPs(17)));  // ширина
        pageSize.setH(BigInteger.valueOf(smToTWIPs(20.3))); // высота
        pageSize.setOrient(STPageOrientation.PORTRAIT);

        // Поля документа
        CTPageMar pageMar = sectPr.isSetPgMar() ? sectPr.getPgMar() : sectPr.addNewPgMar();
        pageMar.setLeft(BigInteger.valueOf(smToTWIPs(2)));    // 2 см слева
        pageMar.setRight(BigInteger.valueOf(smToTWIPs(2)));   // 2 см справа
        pageMar.setTop(BigInteger.valueOf(smToTWIPs(0.7)));   // 0.7 см сверху
        pageMar.setBottom(BigInteger.valueOf(smToTWIPs(0.6)));// 0.6 см снизу
    }
    

    private int getRandomCharacterSpacing() {
        return (int) (-1.4 * 20); // -28 в 1/20 пт
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
