import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by Abhi on 10-06-2016.
 */
public class CreateBlankDocument {
    public CreateBlankDocument(){
        try {
            //Blank Document
            XWPFDocument document= new XWPFDocument();
            //Write the Document in file system
            FileOutputStream out = new FileOutputStream(new File("blank.docx"));
            document.write(out);
            out.close();
            System.out.println("blank.docx written successully");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
