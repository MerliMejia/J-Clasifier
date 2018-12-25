package application;

import java.io.File;

import javafx.fxml.FXML;
import javafx.scene.control.Label;
import javafx.stage.FileChooser;
import javafx.stage.FileChooser.ExtensionFilter;
import archivos.excel;

public class controladorPrincipal {
	
	@FXML
	Label importarClick;
	
	@FXML
	public void clickImportar()
	{
		
		FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("Excel files (*.xls)", "*.xls");
		
		
		FileChooser fc = new FileChooser();
		fc.setInitialDirectory(new File(System.getProperty("user.dir")));
		fc.getExtensionFilters().add(extFilter);
		File archivo = fc.showOpenDialog(null);
		excel.ordenarArchivo(archivo);
	}
	
}
