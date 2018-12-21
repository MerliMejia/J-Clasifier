
package j.ordener;
import j.ordener.archivos.excel;
import java.io.File;
import javax.swing.JFileChooser;
import org.apache.poi.ss.usermodel.Sheet;

public class JOrdener {


    public static void main(String[] args) {
        JFileChooser fc = new JFileChooser("D:\\Dropbox\\Proyectos Personales\\JAVA\\J-Ordener");
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
