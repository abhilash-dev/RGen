package Paragraph;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import javax.swing.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by Abhi on 10-06-2016.
 */
public class CreateParagraph {
    public static void main(String[] args) {
        // A simple GUI component to fetch filename
        String filename = JOptionPane.showInputDialog("Enter the name of the file : ");
        if (filename.isEmpty() || filename == null) {
            JOptionPane.showMessageDialog(null, "Filename cannot be empty", "Error", JOptionPane.OK_CANCEL_OPTION);
        } else if (filename.contains(".docx")) {
            try {
                //Blank Document
                XWPFDocument document = new XWPFDocument();
                //Write the Document in file system
                FileOutputStream out = new FileOutputStream(new File(filename));
                //create Paragraph
                XWPFParagraph paragraph = document.createParagraph();
                XWPFRun run=paragraph.createRun();
                /*run.setText("Jyothy Institute Of Technology (JIT) is one of the best private Bangalore engineering colleges, " +
                        "imparting quality technical education using most modern techniques. These days, when there are so " +
                        "many engg colleges in Karnataka, it is no mean feat that Jyothy Institute Of Technology has earned" +
                        " the reputation of being one of the best colleges in Bangalore for engineering, a centre of excellence" +
                        " where students are prepared for the competitive life ahead.");*/
                String content = JOptionPane.showInputDialog("Enter the content for paragraph: ");
                run.setText(content);
                document.write(out);
                out.close();
                System.out.println(filename + "written successully");
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            //filename.concat(".docx");
            filename = filename + ".docx";
            try {
                //Blank Document
                XWPFDocument document = new XWPFDocument();
                //Write the Document in file system
                FileOutputStream out = new FileOutputStream(new File(filename));
                //create Paragraph
                XWPFParagraph paragraph = document.createParagraph();
                XWPFRun run=paragraph.createRun();
                /*run.setText("Jyothy Institute Of Technology (JIT) is one of the best private Bangalore engineering colleges, " +
                        "imparting quality technical education using most modern techniques. These days, when there are so " +
                        "many engg colleges in Karnataka, it is no mean feat that Jyothy Institute Of Technology has earned" +
                        " the reputation of being one of the best colleges in Bangalore for engineering, a centre of excellence" +
                        " where students are prepared for the competitive life ahead.");*/
                String content = JOptionPane.showInputDialog("Enter the content for paragraph: ");
                run.setText(content);
                document.write(out);
                out.close();
                System.out.println(filename + " written successully");
            } catch (IOException e) {
                e.printStackTrace();
            }


        }
    }
}
