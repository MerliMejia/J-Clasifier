package archivos;

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
import org.controlsfx.control.Notifications;

import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;

import javax.swing.*;

public class excel {
    static Workbook workbook;
    public static void ordenarArchivo(File archivo)
    {
        Sheet hojaxD = ordenar2(lector(archivo), archivo);
        Sheet hoja = workbook.getSheetAt(0);

        FileOutputStream fileOut = null;
        FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("Excel files (*.xls)", "*.xls");
        
        FileChooser fc = new FileChooser();
        fc.getExtensionFilters().add(extFilter);
        fc.setInitialDirectory(new File(System.getProperty("user.dir")));
        DirectoryChooser dc = new DirectoryChooser();
        dc.setInitialDirectory(new File(System.getProperty("user.dir")));
        File f = dc.showDialog(null);
        
        if(f != null)
        {
        	String nombre = f.getPath() + "\\" + archivo.getName().replace(".xls", "");
            try {
                fileOut = new FileOutputStream(nombre + " DEPURADO.xls");
                workbook.write(fileOut);
                fileOut.close();
                Notifications.create().title("ARCHIVO CLASIFICADO!").text("Todos los campos han sido calsificados por departamento!").showConfirm();
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
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
    
    /////////////////////////////////////////////////////////////////////////////////////////////
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
            System.out.println(mapa.get(i));
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
               // System.out.println((String) mapa.get(indices[i - 8]).get(a-1));
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
                //System.out.println((String) mapa2.get(indices2[i - 36]).get(a-1));
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
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    ////////////////////////////////////////////////////////////////////////////////////////////////
    

    public static Sheet ordenar2(Sheet hoja, File archivo)
    {
        Sheet hojaOrdenada = hoja;
        
        //System.out.println(hojaOrdenada.getPhysicalNumberOfRows());
        
       /*for(int i = 0; i < hojaOrdenada.getPhysicalNumberOfRows(); i++)
        {
            System.out.println(hojaOrdenada.getRow(i));
        }*/
       
        ArrayList<String> datos = new ArrayList<>();
        Map<String, ArrayList<String>> mapa1 = new HashMap<String, ArrayList<String>>();

        ArrayList<String> datos2 = new ArrayList<>();
        Map<String, ArrayList<String>> mapa2 = new HashMap<String, ArrayList<String>>();
        
        
       
        for(int i = 8; i < 34; i++)
        {
            datos = new ArrayList<>();
            for(int a = 1; a < hoja.getRow(i).getPhysicalNumberOfCells(); a++)
            {

                hoja.getRow(i).getCell(a).setCellType(Cell.CELL_TYPE_STRING);
                datos.add(hoja.getRow(i).getCell(a).getStringCellValue());
                //System.out.println(hoja.getRow(i).getCell(a).getStringCellValue());
            }
            //System.out.println(datos);
            mapa1.put(hoja.getRow(i).getCell(1).getStringCellValue(), datos);

        }

        for(int i = 36; i < 60; i++)
        {
            datos2 = new ArrayList<>();
            for(int a = 1; a < hoja.getRow(i).getPhysicalNumberOfCells(); a++)
            {
                hoja.getRow(i).getCell(a).setCellType(Cell.CELL_TYPE_STRING);
                datos2.add(hoja.getRow(i).getCell(a).getStringCellValue());
            }
            
            mapa2.put(hoja.getRow(i).getCell(1).getStringCellValue(), datos2);

        }
        
        String[] nombres = new String[] {
        		"      Pool",
        	      "      Cleanliness",
        	      "      Room",
        	      "      Food",
        	      "      Room Service",
        	      "      Breakfast",
        	      "      Dinner Buffet",
        	      "      Beach / Pool Restaurant",
        	      "      Lunch Buffet",
        	      "      Beverage",
        	      "      Bar",
        	      "      A La Carte Restaurants",
        	      "      Facilities",
        	      "      Reception-Guest Service",
        	      "      Porterage Service",
        	      "      Butler Service",
        	      "      Emotions",
        	      "      Concierge Service",
        	      "      SPA",
        	      "      Internet",
        	      "      Lobby Shop",
        	      "      Entertainment",
        	      "      Daytime",
        	      "      Bahia Principe Village",
        	      "      Beach",
        	      "      Evening"
        };
        
        String[] nombres2 = new String[] {
        		"      Cleanliness",
        	      "      Room",
        	      "      Pool",
        	      "      Food",
        	      "      Breakfast",
        	      "      Dinner Buffet",
        	      "      Beach / Pool Restaurant",
        	      "      Lunch Buffet",
        	      "      A La Carte Restaurants",
        	      "      Beverage",
        	      "      Bar",
        	      "      Entertainment",
        	      "      Bahia Principe Village",
        	      "      Beach",
        	      "      Evening",
        	      "      Daytime",
        	      "      Facilities",
        	      "      SPA",
        	      "      Reception-Guest Service",
        	      "      Concierge Service",
        	      "      Emotions",
        	      "      Porterage Service",
        	      "      Children",
        	      "      Lobby Shop",
        	      "      Internet",
        };
        
        ArrayList<ArrayList<String>> agregar = new ArrayList<ArrayList<String>>();
        ArrayList<ArrayList<String>> agregar2 = new ArrayList<ArrayList<String>>();
        
        int a = 0;
        while(agregar.size() < nombres.length)
        {
        	//System.out.println(nombres[a]);
        	agregar.add(mapa1.get(nombres[a]));
        	a++;
        }
        
        a = 0;
        
        while(agregar2.size() < nombres2.length)
        {
        	//System.out.println(nombres[a]);
        	agregar2.add(mapa2.get(nombres2[a]));
        	a++;
        }
        
        for(int i = 8; i < 34; i++)
        {
        	for(int x = 0; x < agregar.get(i - 8).size(); x++)
        	{
        		System.out.println(agregar.get(i-8).get(x));
        		try
                {
                    hoja.getRow(i).getCell(x + 1).setCellValue(Double.parseDouble((String) agregar.get(i-8).get(x)));

                }catch (Exception e)
                {
                    hoja.getRow(i).getCell(x + 1).setCellValue((String) agregar.get(i-8).get(x));
                }
        	}
        }
        
        for(int i = 36; i < 60; i++)
        {
        	for(int x = 0; x < agregar2.get(i - 36).size(); x++)
        	{
        		//System.out.println(agregar2.get(i-8).get(x));
        		try
                {
                    hoja.getRow(i).getCell(x + 1).setCellValue(Double.parseDouble((String) agregar2.get(i-36).get(x)));

                }catch (Exception e)
                {
                    hoja.getRow(i).getCell(x + 1).setCellValue((String) agregar2.get(i-36).get(x));
                }
        	}
        }
        
        /*for(int i = 8; i < 34; i++)
        {
            //System.out.println(mapa.get(i).size());
            for(int a = 1; a < mapa.get(i).size(); a++)
            {
               // System.out.println((String) mapa.get(indices[i - 8]).get(a-1));
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
                //System.out.println((String) mapa2.get(indices2[i - 36]).get(a-1));
                try
                {
                    hoja.getRow(i).getCell(a).setCellValue(Double.parseDouble((String) mapa2.get(indices2[i - 36]).get(a-1)));

                }catch (Exception e)
                {
                    hoja.getRow(i).getCell(a).setCellValue((String) mapa2.get(indices2[i - 36]).get(a-1));
                }


            }


        }*/


        return hoja;
    }
    
}