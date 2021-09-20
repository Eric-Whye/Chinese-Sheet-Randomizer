import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Collections;
import java.util.StringTokenizer;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTabStop;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTabJc;


public class Main {

    public static void main(String[] args) {
        try {

            String path = "C:\\Users\\Eric\\OneDrive - University College Dublin\\Documents";
            File fp = new File(path + "\\Chinese Practice Sheet.docx");
            File fp2 = new File(path + "\\Chinese Practice Sheet for Print.docx");

            double SPACING = 1.4;
            double FONT_SIZE = 14;


            if (!fp.exists()) {
                System.out.println("File isn't there");
            }

            //Creating the word doc to print
            try {
                if (fp2.createNewFile()) {
                    System.out.println("File created in " + path);
                } else {
                    System.out.println("File already exists at " + path);
                }
            } catch (Exception ex) {
                System.out.println("Caught an Exception while creating Word File at " + path + ": " + ex);
            }


            //Reading in the word doc, putting all practice words into stringtokenizer, then into an arraylist, then randomized and passed over to output below
            FileInputStream fis = new FileInputStream(fp.getAbsolutePath());
            XWPFDocument document = new XWPFDocument(fis);
            XWPFWordExtractor extractor = new XWPFWordExtractor(document);
            String fileData = extractor.getText();


            assert fileData != null;
            fileData = fileData.replaceAll("\n", "\t");
            StringTokenizer tokens = new StringTokenizer(fileData, "\t");

            ArrayList<String> holder = new ArrayList<>();
            String token = null;
            int k = 0;
            while (tokens.hasMoreTokens()) {
                token = tokens.nextToken();
                if (token.trim().equals(""))
                    continue;
                if (k++ % 2 == 0) {
                    holder.add(token);
                }
            }

            Collections.shuffle(holder);


            //To output all the practice words with appropriate spacing and tab stops in random order
            FileOutputStream fs = new FileOutputStream(fp2.getAbsolutePath());
            XWPFDocument output = new XWPFDocument();
            int j = 0;

            XWPFParagraph paragraph = output.createParagraph();
            XWPFRun runx = paragraph.createRun();
            paragraph.setAlignment(ParagraphAlignment.LEFT);
            paragraph.setSpacingBetween(SPACING);
            runx.setFontSize(FONT_SIZE);
            try {
                for (int i = 0; i < holder.size() / 2; i++) {

                /*
                for (IBodyElement : document.getBodyElements){This for loop used to be the above for loop
                    if(element instanceof XWPFParagraph) {
                        XWPFParagraph paragraph = (XWPFParagraph)element;

                        if (paragraph.getStyleID() != null){
                            XWPFStyles styles = output.createStyles();
                            XWPFStyles stylesdoc2= document.getStyles();
                            styles.addStyle(stylesdoc2.getStyle(paragraph.getStyleID()));
                        }*/

                    //paragraph = output.createParagraph();
                    //runx = paragraph.createRun();

                    //paragraph.setAlignment(ParagraphAlignment.LEFT);
                    //runx.setFontSize(14);

                    runx.setText(holder.get(j++));

                    int twipsPerInch = 1280; //measurement unit for tab stop pos is twips (twentieth of an inch point)

                    CTTabStop tabStop = paragraph.getCTP().getPPr().addNewTabs().addNewTab();
                    tabStop.setVal(STTabJc.LEFT);
                    tabStop.setPos(BigInteger.valueOf(4 * twipsPerInch));

                    runx.addTab();
                    runx.setText(holder.get(j++));
                    runx.addBreak();
                    runx.addCarriageReturn();
                }
            } catch (Exception ex) {
                ex.printStackTrace();
            }

            output.write(fs);
            fs.close();
            System.out.println("done");
        } catch (IOException ex){ ex.printStackTrace();}
    }

    private static void writeTextToPdfFile(String path, String name) throws IOException {
        if (name.charAt(0) != '\\' && name.charAt(1) != '\\'){
            StringBuilder temp = new StringBuilder(name);
            temp.insert(0, "\\\\");
            name = temp.toString();
        }
        String FILE_PATH_NAME = path + name;

        try (PDDocument doc = new PDDocument()) {

            /*
             * Create a PDF Page:
             * PDF Page 1 ->
             */
            PDPage myPage = new PDPage();
            doc.addPage(myPage);

            try (PDPageContentStream cont = new PDPageContentStream(doc, myPage)) {

                cont.beginText();

                cont.setFont(PDType1Font.TIMES_BOLD, 12);
                cont.setLeading(15.5f);

                cont.newLineAtOffset(25, 700);
                String line1 = "Who We Are?";

                cont.showText(line1);
                cont.newLine();
                cont.newLine();

                cont.setFont(PDType1Font.TIMES_ROMAN, 12);
                String line2 = "- We are passionate engineers in software development. ";
                cont.showText(line2);
                cont.newLine();

                String line3 = "- We focus on how to do things that can both respect users and make money.";
                cont.showText(line3);
                cont.newLine();
                
                cont.endText();
            }

            doc.save(FILE_PATH_NAME);

            /*
             * Create a new PDF Page
             * -> PDF Page 2:
             */
            PDPage myPage2 = new PDPage();
            doc.addPage(myPage2);

            try (PDPageContentStream cont = new PDPageContentStream(doc, myPage2)) {

                cont.beginText();

                cont.setLeading(15.5f);

                cont.newLineAtOffset(25, 700);
                cont.setFont(PDType1Font.TIMES_ROMAN, 12);

                // line 1 = "What does grokonez mean?"
                cont.showText("What does ");

                cont.setFont(PDType1Font.TIMES_BOLD, 12);
                cont.showText("grokonez");

                cont.setFont(PDType1Font.TIMES_ROMAN, 12);
                cont.showText(" mean?");

                cont.newLine();
                cont.newLine();

                cont.setFont(PDType1Font.TIMES_ROMAN, 12);

                String line2 = "Well, ‘grokonez’ is derived from the words ‘grok’ and ‘konez’.";
                cont.showText(line2);
                cont.newLine();

                String line3 = "– ‘grok’ means understanding (something) intuitively or by empathy.";
                cont.showText(line3);
                cont.newLine();

                String line4 = "– ‘konez’ expresses ‘connect’ that represents the idea ‘connect the dots’, ‘connect everything’.";
                cont.showText(line4);
                cont.newLine();

                cont.endText();
            }

            doc.save(FILE_PATH_NAME);

            System.out.println("Done!");
        }
    }
}

