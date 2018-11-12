package google;

import java.io.File;
import java.io.IOException;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class Acesso {

	public static void main(String[] args) throws InterruptedException, BiffException, IOException {
		
	  	System.setProperty("webdriver.chrome.driver", "bin/tools/chromedriver.exe");
	  	WebDriver driver = new ChromeDriver();
	  	
	  //Maximizo a janela do Browser
		driver.manage().window().maximize();
		
		driver.get("https://www.google.com.br/");
		
		//Deleto todos os cookies da pagina para não ter interferência nos testes
		driver.manage().deleteAllCookies();

		
		//Dados da planilha com a biblioteca JXL
		
		Workbook workbook = Workbook.getWorkbook(new File("bin/planilha/dados.xls"));
		
		//Em qual sheet começara a ler a planilha
		Sheet sheet = workbook.getSheet(0);  
		
		//Conte o número de linhas
		int rowCount = sheet.getRows();
		System.out.println("Existem " + rowCount + " linhas com dados na " + sheet.getName() + " porém a primeira linha é o cabeçalho.");

		
		//indica em qual linha deve começar a ler os dados
		int comece_ler_nesta_linha = 1;
		
		//comece o loop
		for(int i =  comece_ler_nesta_linha; i < rowCount; i++){
			
		    String textocelula = sheet.getCell(0, i).getContents();
	    
		    //Pegue o texto da planilha e digite na caixa de texto
		    driver.findElement(By.name("q")).sendKeys(textocelula);
		    
		    WebElement botaoPesquisar = driver.findElement(By.name("btnK"));
		    
		    
		    
		    //Este IF esta aqui porque o botão pesquisar muda de nome após a primeira pesquisa		    
		    if (botaoPesquisar.isDisplayed() )
		    {
		    //Clique no botão pesquisar
		    driver.findElement(By.name("btnK")).click();
		    }
		    else
		    {
			    driver.findElement(By.name("btnG")).click();
		    }
		    
		    //Limpe a caixa de texto, para a próxima pesquisa
		    driver.findElement(By.name("q")).clear();	    
		    
		 	    
			Thread.sleep(2000);
		}
		
		//Fecha a planilha
		workbook.close();
		
		//Encerra o driver
		driver.quit();

	}

}
