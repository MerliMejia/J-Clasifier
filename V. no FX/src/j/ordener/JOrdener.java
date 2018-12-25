
package j.ordener;
import com.bulenkov.darcula.DarculaLaf;
import j.ordener.archivos.excel;
import java.io.File;
import javax.swing.*;
import javax.swing.plaf.basic.BasicLookAndFeel;

import org.apache.poi.ss.usermodel.Sheet;

public class JOrdener {


    public static void main(String[] args) {

        BasicLookAndFeel darcula = new DarculaLaf();
        try {
            UIManager.setLookAndFeel(darcula);
        } catch (UnsupportedLookAndFeelException e) {
            e.printStackTrace();
        }

        JFileChooser fc = new JFileChooser(System.getProperty("user.dir"));
        int aceptar = fc.showOpenDialog(null);
        if(aceptar == JFileChooser.APPROVE_OPTION)
        {
            excel.ordenarArchivo(fc.getSelectedFile());
            
            //System.out.println(hoja.getRow(7).getCell(0));
        }else
        {
            
        }
    }
    
}
