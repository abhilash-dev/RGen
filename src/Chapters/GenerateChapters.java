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

    public static void createDoc(String filename) {
        try {
            //Blank Document
            XWPFDocument document = new XWPFDocument();
            //Write the Document in file system
            FileOutputStream out = new FileOutputStream(new File(filename));
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

            run.setText("Chapter 1");
            run.addBreak();

            //run.setText("Introduction");
            String content = JOptionPane.showInputDialog("Enter the name of the Chapter: ");
            run.setText(content);
            addNewLine(5, run);

            document.write(out);
            out.close();
            System.out.println(filename + " written successully");
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
        // A simple GUI component to fetch filename
        String filename = JOptionPane.showInputDialog("Enter the name of the file : ");
        if (filename.isEmpty() || filename == null) {
            JOptionPane.showMessageDialog(null, "Filename cannot be empty", "Error", JOptionPane.OK_CANCEL_OPTION);
        } else if (filename.contains(".docx")) {
            createDoc(filename);
        } else {
            filename = filename + ".docx";
            createDoc(filename);
        }
    }
}
