package ExcelWrite;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;



public class ExcelRead 
{

   public static void main(String[] args) throws IOException
		{
			ArrayList<Integer> a=new ArrayList<Integer>();
			{
				try
				{
					File f=new File("C:\\Users\\dasam\\OneDrive\\Desktop\\Book1.xlsx");
					FileInputStream file=new FileInputStream(f);
					XSSFWorkbook wo=new XSSFWorkbook(file);
					XSSFSheet sheet=wo.getSheetAt(0);

					Iterator<Row> itr=sheet.iterator();
					while(itr.hasNext())
					{
						Row row=itr.next();
						Iterator<Cell> cellitr=row.cellIterator();
						while(cellitr.hasNext())
						{
							Cell cell=cellitr.next();
							
							switch(cell.getCellType())
							{
							case STRING:
								System.out.println(cell.getStringCellValue()+"\t\t\t");
								break;
							case BOOLEAN:
								System.out.println(cell.getBooleanCellValue()+"\t\t\t");
								break;
							case FORMULA:
								System.out.println(cell.getDateCellValue()+"\t\t\t");
								break;
							case NUMERIC:
								System.out.println(cell.getNumericCellValue()+"\t\t\t");
								a.add((int)cell.getNumericCellValue());
								break;
								default:
									break;
									
							}
						}
						wo.close();
						System.out.println(" ");
					}
				}
				
						catch(Exception e)
							{
							e.printStackTrace();
							
							}
				System.out.println(a);
				System.out.println("Total amount for the customer");
				int p=(a.get(0)*a.get(1)*a.get(2));
				System.out.println(p);
			}
		}
	}


	


