package br.com.mab;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ManipulaXlsx {
	
	//private String caminhoArquivo = "src/br/com/mab/Balancete.xlsx";
	private String caminhoArquivo = "arquivos/teste.xlsx";
	
	public void lerArquivo() {
		
        FileInputStream arquivoEntrada = null;
		try {
			arquivoEntrada = new FileInputStream(new File(caminhoArquivo));
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

        // Criar uma instância de Workbook a partir do FileInputStream
        Workbook workbook = null;
		try {
			workbook = new XSSFWorkbook(arquivoEntrada);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

        // Obter a primeira planilha (índice 0)
        Sheet sheet = workbook.getSheetAt(0);

        // Iterar pelas linhas da planilha
        for (Row row : sheet) {
            // Iterar pelas células de cada linha
            for (Cell cell : row) {
                // Imprimir o conteúdo da célula
                System.out.print(cell.toString() + "\t");
            }
            System.out.println(); // Nova linha após cada linha da planilha
        }

        // Fechar o FileInputStream e liberar recursos
        try {
			arquivoEntrada.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}		
		
	}

}
