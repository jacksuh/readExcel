import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.ObjectInputStream.GetField;
import java.sql.Date;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.Period;
import java.time.format.DateTimeFormatter;
import java.util.Formatter;
import java.util.Iterator;
import java.util.concurrent.TimeUnit;

import javax.xml.crypto.Data;

import org.apache.poi.ss.format.CellDateFormatter;
import org.apache.poi.ss.formula.functions.Now;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFRow.CellIterator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class LendoXLSX {

	public static void main(String[] args) {

		
		try {
		File file = new File("D:\\Excel\\dados.xlsx");
			FileInputStream filePlanilha = new FileInputStream(file);
			
			
			//Cria um workbook com todas as abas
			try {
				XSSFWorkbook workbook = new XSSFWorkbook(filePlanilha);
				
				
				//primeira aba da planilha
				XSSFSheet sheet = workbook.getSheetAt(0);
				
				//retorna todas as linhas da planilha zero
				Iterator<Row> rowIterator = sheet.iterator();
				
				//Varre todas as linhas enquanto tiver dados leia
				while(rowIterator.hasNext()) {
					
					//recebe cada linha da planilha
					Row  row = rowIterator.next();
					
					
					
					//pegamos todas as celulas da linha
					Iterator<Cell> cellIterator = row.iterator();
					
					//varremos todas as celulas da linha atual
					while(cellIterator.hasNext()) {
						
						
						//criamos uma celula
						Cell cell = cellIterator.next();
						
						//LocalDate x = LocalDate.now();
						Date x = Date.valueOf(LocalDate.now());
						DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM/yyyy");
						//String y = x.format(formatter);
						
					
						switch(cell.getCellType()) {
						
						case Cell.CELL_TYPE_STRING:
						 System.out.println("Nome:" + cell.getStringCellValue());
						 break;
						 
						case Cell.CELL_TYPE_NUMERIC:
						
							
							long diffemMil = Math.abs(x.getTime() - cell.getDateCellValue().getTime()); 
							
							long diff = TimeUnit.DAYS.convert(diffemMil, TimeUnit.MILLISECONDS);
								
							int dias = (int) (diffemMil / (1000 * 60 * 60 * 24));
							
							
							int tresmes = 90;
							int seismes = 180;
							
							if(dias >= tresmes && dias < seismes) {
								
								System.out.println("Não acessa o sistema a mais de 3 meses efetuar o bloqueio.");
							} 
							
							if (dias >= seismes) {
								
								System.out.println("Não acessa o sistema a mais de 6 meses efetuar o cancelamento.");
							
							}else{
								System.out.println("Dentro do prazo");
							}
						
							//String a = cell.getDateCellValue().toLocaleString().toString();
							
							
							//System.out.println("diff" + diff);
							System.out.println("Dias: " +  dias);
							//LocalDate w = LocalDate.from(a);
		
							
							//Period periodo = Period.between(x, w);
						
							break;
							
							}
						
					}
				}
				
		
			} catch (IOException e) {
				e.printStackTrace();
			}
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
		
	}
	
	
	

}
