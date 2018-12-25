
package j.ordener.archivos;

import java.io.*;
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

import javax.swing.*;

public class excel {
    static Workbook workbook;
    public static void ordenarArchivo(File archivo)
    {
        Sheet hojaxD = ordenar(lector(archivo), archivo);
        Sheet hoja = workbook.getSheetAt(0);

        //System.out.println(hoja.getPhysicalNumberOfRows());
        for(int i = 0; i < hoja.getPhysicalNumberOfRows(); i++)
        {
            for(int a = 0; a < hoja.getRow(i).getPhysicalNumberOfCells(); a++)
            {
                if(hoja.getRow(i).getCell(a).getCellType() == Cell.CELL_TYPE_STRING)
                {
                    //System.out.println(hoja.getRow(i).getCell(a).getStringCellValue());
                }else if(hoja.getRow(i).getCell(a).getCellType() == Cell.CELL_TYPE_NUMERIC)
                {
                    //System.out.println(hoja.getRow(i).getCell(a).getNumericCellValue());
                }
            }
        }

        FileOutputStream fileOut = null;
        JFileChooser fc = new JFileChooser(System.getProperty("user.dir"));
        fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

        int aceptar = fc.showOpenDialog(null);
        if(aceptar == JFileChooser.APPROVE_OPTION)
        {
            String nombre = fc.getSelectedFile().getPath() + "\\" + archivo.getName().replace(".xls", "");
            try {
                fileOut = new FileOutputStream(nombre + " DEPURADO.xls");
                workbook.write(fileOut);
                fileOut.close();
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }

            //System.out.println(hoja.getRow(7).getCell(0));
        }else
        {


        }




    }
    
    public static Sheet lector(File archivo)
    {
        try {
            workbook = new HSSFWorkbook(new FileInputStream(archivo));
            Sheet hoja = workbook.getSheetAt(0);
            return hoja;
        } catch (FileNotFoundException ex) {
            Logger.getLogger(excel.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(excel.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        return null;
    }
    
    public static Sheet ordenar(Sheet hoja, File archivo)
    {
        Sheet hojaOrdenada = hoja;
        
        //System.out.println(hojaOrdenada.getPhysicalNumberOfRows());
        
       /*for(int i = 0; i < hojaOrdenada.getPhysicalNumberOfRows(); i++)
        {
            System.out.println(hojaOrdenada.getRow(i));
        }*/
       
        ArrayList<String> datos = new ArrayList<>();
        Map<Integer, ArrayList> mapa = new HashMap<>();

        ArrayList<String> datos2 = new ArrayList<>();
        Map<Integer, ArrayList> mapa2 = new HashMap<>();
        
        
       
        for(int i = 8; i < 34; i++)
        {
            datos = new ArrayList<>();
            for(int a = 1; a < hoja.getRow(i).getPhysicalNumberOfCells(); a++)
            {

                hoja.getRow(i).getCell(a).setCellType(Cell.CELL_TYPE_STRING);
                datos.add(hoja.getRow(i).getCell(a).getStringCellValue());

            }
            mapa.put(i, datos);
        }

        for(int i = 36; i < 60; i++)
        {
            datos2 = new ArrayList<>();
            for(int a = 1; a < hoja.getRow(i).getPhysicalNumberOfCells(); a++)
            {
                hoja.getRow(i).getCell(a).setCellType(Cell.CELL_TYPE_STRING);
                datos2.add(hoja.getRow(i).getCell(a).getStringCellValue());
            }
            mapa2.put(i, datos2);
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

        int[] indices2 = new int[]{

                36,
                46,
                38,
                44,
                39,
                43,
                45,
                53,
                48,
                51,
                49,
                54,
                57,
                59,
                52,
                55,
                56,
                37,
                40,
                41,
                42,
                50,
                47,
                58,
                60
        };
        
        for(int i = 8; i < 34; i++)
        {
            //System.out.println(mapa.get(i).size());
            for(int a = 1; a < mapa.get(i).size(); a++)
            {
                System.out.println((String) mapa.get(indices[i - 8]).get(a-1));
                try
                {
                    hoja.getRow(i).getCell(a).setCellValue(Double.parseDouble((String) mapa.get(indices[i - 8]).get(a-1)));

                }catch (Exception e)
                {
                    hoja.getRow(i).getCell(a).setCellValue((String) mapa.get(indices[i - 8]).get(a-1));
                }
                
            }
            
            
        }

        for(int i = 36; i < 60; i++)
        {
            //System.out.println(mapa.get(i).size());
            for(int a = 1; a < mapa2.get(i).size(); a++)
            {
                System.out.println((String) mapa2.get(indices2[i - 36]).get(a-1));
                try
                {
                    hoja.getRow(i).getCell(a).setCellValue(Double.parseDouble((String) mapa2.get(indices2[i - 36]).get(a-1)));

                }catch (Exception e)
                {
                    hoja.getRow(i).getCell(a).setCellValue((String) mapa2.get(indices2[i - 36]).get(a-1));
                }


            }


        }


        return hoja;
    }
    
}
