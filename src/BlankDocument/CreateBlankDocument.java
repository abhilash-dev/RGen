package BlankDocument;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import javax.swing.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by Abhi on 10-06-2016.
 */
public class CreateBlankDocument {
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
                document.write(out);
                out.close();
                System.out.println(filename + " written successully");
            } catch (IOException e) {
                e.printStackTrace();
            }


        }
    }
}