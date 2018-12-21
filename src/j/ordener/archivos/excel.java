
package j.ordener.archivos;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class excel {
    
    public static void ordenarArchivo(File archivo)
    {
        Sheet hoja = ordenar(lector(archivo));
        //System.out.println(hoja.getPhysicalNumberOfRows());
        for(int i = 0; i < hoja.getPhysicalNumberOfRows(); i++)
        {
            for(int a = 0; a < hoja.getRow(i).getPhysicalNumberOfCells(); a++)
            {
                if(hoja.getRow(i).getCell(a).getCellType() == Cell.CELL_TYPE_STRING)
                {
                    System.out.println(hoja.getRow(i).getCell(a).getStringCellValue());
                }else if(hoja.getRow(i).getCell(a).getCellType() == Cell.CELL_TYPE_NUMERIC)
                {
                    System.out.println(hoja.getRow(i).getCell(a).getNumericCellValue());
                }
            }
        }
        
    }
    
    public static Sheet lector(File archivo)
    {
        try {
            Workbook workbook = new HSSFWorkbook(new FileInputStream(archivo));
            Sheet hoja = workbook.getSheetAt(0);
            return hoja;
        } catch (FileNotFoundException ex) {
            Logger.getLogger(excel.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(excel.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        return null;
    }
    
    public static Sheet ordenar(Sheet hoja)
    {
        Sheet hojaOrdenada = hoja;
        
        //System.out.println(hojaOrdenada.getPhysicalNumberOfRows());
        
       /*for(int i = 0; i < hojaOrdenada.getPhysicalNumberOfRows(); i++)
        {
            System.out.println(hojaOrdenada.getRow(i));
        }*/
       
        ArrayList<String> datos = new ArrayList<>();
        Map<Integer, ArrayList> mapa = new HashMap<>();
        
        
       
        for(int i = 8; i < 34; i++)
        {
            datos = new ArrayList<>();
            for(int a = 1; a < hoja.getRow(i).getPhysicalNumberOfCells(); a++)
            {
                if(hoja.getRow(i).getCell(a).getCellType() == Cell.CELL_TYPE_STRING)
                {
                    datos.add(hoja.getRow(i).getCell(a).getStringCellValue());
                }else
                {
                    datos.add(Double.toString(hoja.getRow(i).getCell(a).getNumericCellValue()));
                }
            }
            mapa.put(i, datos);
        }
        
        int[] indices = new int[]{
        
            8,
            9,
            11,
            19,
            12,
            18,
            22,
            25,
            21,
            17,
            23,
            26,
            10,
            13,
            14,
            20,
            15,
            33,
            30,
            16,
            24,
            29,
            28,
            27,
            31,
            32
        };
        
        for(int i = 8; i < 34; i++)
        {
            //System.out.println(mapa.get(i).size());
            for(int a = 0; a < mapa.get(i).size(); a++)
            {
                hojaOrdenada.getRow(i).getCell(a).setCellValue((String) mapa.get(indices[i - 8]).get(a));
                
            }
            
            
        }
        
        return hojaOrdenada;
    }
    
    
}
