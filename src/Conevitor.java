
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.util.Iterator;
import java.util.Properties;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class Conevitor {
	static String dataDirName = "J:\\Workspace\\XLSToCSVConeVitor\\data\\";
	//static final String testInputFileName = "test_xls.xls";
	//static final String testOutputFileName = "test_csv.csv";

	static void xlsToText(File inputFile, String delimiter) {
		
		File outputFile = new File(getOutputFile(inputFile.getAbsolutePath(), delimiter));
		System.out.println(outputFile.getAbsolutePath());
		// For storing data into CSV files
		StringBuffer data = new StringBuffer();

		try {
			FileOutputStream fos = new FileOutputStream(outputFile);
			// Get the workbook object for XLSX file
			/*XSSFWorkbook wBook = new XSSFWorkbook(
					new FileInputStream(inputFile));
			// Get first sheet from the workbook
			XSSFSheet sheet = wBook.getSheetAt(0);*/
			
			Workbook wBook = new HSSFWorkbook(new FileInputStream(inputFile));
			// Get first sheet from the workbook
			Sheet sheet = wBook.getSheetAt(0);
			
			Row row;
			Cell cell;
			// Iterate through each rows from first sheet
			Iterator<Row> rowIterator = sheet.iterator();

			while (rowIterator.hasNext()) {
				row = rowIterator.next();

				// For each row, iterate through each columns
				Iterator<Cell> cellIterator = row.cellIterator();
				while (cellIterator.hasNext()) {

					cell = cellIterator.next();

					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_BOOLEAN:
						data.append(cell.getBooleanCellValue() + delimiter);

						break;
					case Cell.CELL_TYPE_NUMERIC:
						data.append(cell.getNumericCellValue() + delimiter);

						break;
					case Cell.CELL_TYPE_STRING:
						data.append(cell.getStringCellValue() + delimiter);
						break;

					case Cell.CELL_TYPE_BLANK:
						data.append("" + delimiter);
						break;
					default:
						data.append(cell + delimiter);

					}
				}
				data.append("\r\n");
			}

			fos.write(data.toString().getBytes());
			fos.close();

		} catch (Exception ioe) {
			ioe.printStackTrace();
		}
	}
	
	private static String getOutputFile(String inputFileName, String delimiter) {
		if(inputFileName.lastIndexOf('.')>0)
        {
           // get last index for '.' char
           int lastIndex = inputFileName.lastIndexOf('.');
           
           // get extension
           String extn = inputFileName.substring(lastIndex);
           String fileName = inputFileName.substring(0,lastIndex);
           // match path name extension
           if(extn.equals(".xls"))
           {
              if(delimiter.equals(",")) {
            	  return (fileName+".csv");
              } else if(delimiter.equals("|")) {
            	  return (fileName+".txt");
              }
           }
        }
		return null;
	}

	private static File[] loadXLSFiles(File dataDir) { 

		 FilenameFilter fileNameFilter = new FilenameFilter() {
           @Override
           public boolean accept(File dir, String name) {
              if(name.lastIndexOf('.')>0)
              {
                 // get last index for '.' char
                 int lastIndex = name.lastIndexOf('.');
                 
                 // get extension
                 String str = name.substring(lastIndex);
                 
                 // match path name extension
                 if(str.equals(".xls"))
                 {
                    return true;
                 }
              }
              return false;
           }
        };
        
        return  dataDir.listFiles(fileNameFilter);
	}

	// testing the application

	public static void main(String[] args) throws Exception {
		
		dataDirName = loadDir();
		File dataDir = new File(dataDirName);
		if(dataDir.isDirectory()) {		
			
			
			File[] fList = loadXLSFiles(dataDir);
			System.out.println("dataDir::"+dataDir);
			
			for(File inputFile : fList){
				// generate CSV
				String delimiter = ",";
				xlsToText(inputFile, delimiter);
				// generate TXT
				delimiter = "|";				
				xlsToText(inputFile, delimiter);				
			}
		}	
		
		System.out.println("Convertion Complete!");
	}

	private static String loadDir()  throws Exception {
		// TODO Auto-generated method stub
		Properties propFile = new Properties();
		propFile.load(new FileInputStream("dir.properties"));
		return propFile.getProperty("DIR");
	}
}
