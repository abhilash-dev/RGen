package Chapters;

import org.apache.poi.xwpf.usermodel.*;

import javax.swing.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import static org.apache.poi.xwpf.usermodel.ParagraphAlignment.*;

/**
 * Created by Abhi on 11-06-2016.
 */
public class GenerateChapters {

    public static void createDoc(int fileCount) {
        try {
            //Blank Document
            XWPFDocument document = new XWPFDocument();
            //Write the Document in file system
            FileOutputStream out = new FileOutputStream(new File("Chapter - " + (fileCount + 1) + ".docx"));
            //create Paragraph
            XWPFParagraph paragraph = document.createParagraph();
            //styling the paragraph contents
            paragraph.setAlignment(CENTER);
            paragraph.setVerticalAlignment(TextAlignment.CENTER);

            XWPFRun run = paragraph.createRun();
            run.setFontSize(36);
            run.setFontFamily("Times New Roman");
            run.setBold(true);
            run.setItalic(true);

            addNewLine(7, run);

            run.setText("Chapter " + (fileCount + 1));
            run.addBreak();

            String chapterTitle = JOptionPane.showInputDialog("Enter the title for Chapter " + (fileCount + 1) + ": ");
            run.setText(chapterTitle);
            addNewLine(5, run);

            document.write(out);
            out.close();
            System.out.println("Chapter " + (fileCount + 1) + " written successully");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void addNewLine(int count, XWPFRun run) {
        for (int i = 0; i < count; i++) {
            run.addBreak();
        }
    }

    public static void main(String[] args) {
        int n = Integer.parseInt(JOptionPane.showInputDialog("Enter the no. of Chapters"));
        for (int i = 0; i < n; ++i) {
            createDoc(i);
        }
    }
}
