package poi;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.IOException;
import java.util.Iterator;

public class ExcelReader {

	public static final String caminho_planilha = "bin/tools/dados.xlsx";

    public static void main(String[] args) throws IOException, InvalidFormatException {


    	// Criando uma pasta de trabalho a partir de um arquivo do Excel (.xls ou .xlsx)
        Workbook workbook = WorkbookFactory.create(new File(caminho_planilha));

        // Recuperando o número de folhas na pasta de trabalho
        System.out.println("A planilha tem " + workbook.getNumberOfSheets() + " sheets \n");


        // 1. Você pode obter um sheetIterator e recuperar o nome de todas as sheets da planilha
        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        System.out.println("Recuperando o nome de todas as sheets utilizando o comando iterator: \n ");
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            System.out.println("=> " + sheet.getSheetName());
        }
        
        

        /*
           ================================================== ================
           Iterando sobre todas as linhas e colunas em uma planilha (várias maneiras)
           ================================================== ================
        */
        
        

        //Obtendo a planilha no índice zero
        Sheet sheet = workbook.getSheetAt(0);


        //Cria um DataFormatter para formatar e obter o valor de cada célula como String
        DataFormatter dataFormatter = new DataFormatter();


     // 1. Você pode obter um rowIterator e columnIterator e iterar sobre eles
        System.out.println("\n\n Iterando sobre linhas e colunas usando o iterator: \n");
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();


         // Agora vamos iterar sobre as colunas da linha atual
            Iterator<Cell> iteando_celula = row.cellIterator();

            while (iteando_celula.hasNext()) {
                Cell celula = iteando_celula.next();
                String valor_da_celula = dataFormatter.formatCellValue(celula);
                System.out.print(valor_da_celula + "\t");
            }
            System.out.println();
        }
        

     //Fechando a pasta de trabalho
        workbook.close();
    }
}
