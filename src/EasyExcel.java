/*
	A Java-based Excel using Apache POI
	Author: Bird Liu (Liu Xin)
	License: GNU GPLv2

	GitHub:
		https://www.github.com/JavaHomework
*/

import java.io.FileOutputStream;
import java.util.Date;
import java.util.*;
import java.io.*;
import org.apache.poi.ss.examples.ToCSV;
import org.FRing.ToExcel;

public class EasyExcel {
	public static void help() {
			System.out.println("5Ring(R) EasyExcel(CSV)");
			System.out.println("	An open-source CSV reader based on Java with .xls support.");
			System.out.println("	The Java homework project of Lanzhou University.");
			System.out.println("");
			System.out.println("Usage:");
			System.out.println("  --start-gui: Run EasyExcel(CSV) in GUI mode");
			System.out.println("  --show-csv filename: Show a CSV file");
			System.out.println("  --show-xls filename: Show a XLS file");
			System.out.println("  --convert-xls filename: Convert a XLS file to CSV file to edit it as text");
			System.out.println("  --convert-csv filename: Convert a CSV file to XLS file to open it by Excel");
			System.out.println("  --version: Version Information of EasyExcel(CSV)");
			System.out.println("  --designer: Show Designer of EasyExcel(CSV)");
			System.out.println("  --help: Show help document");
	}
	public static void main(String[] args) {
		try {
			if (args[0].equals("--start-gui"))
				new MainWindow();
			else if (args[0].equals("--version")) {
				System.out.println("5Ring(R) EasyExcel");
				System.out.println("V1.0 Beta (CSV) build 20151221 on OpenJDK");
			}
			else if (args[0].equals("--designer")) {
				System.out.println("Designer: Liu Xin");
				System.out.println("ID: 320130938311");
				System.out.println("Major: 2013 Information Security");
			}
			else if (args[0].equals("--show-xls")) {
				try {
					ToCSV source = new ToCSV();
					File f = new File(args[1]);
					source.convertExcelToCSV(f.toString(), ".");
					String tempName[] = f.getName().split("\\.");
					File f_t = new File(tempName[0] + ".csv");
					//System.out.println(f_t.toString());
					InputStreamReader read = new InputStreamReader (new FileInputStream(f_t), "UTF-8");
					BufferedReader bufferedReader = new BufferedReader(read);
					String tempStr;
					int col = 0; int row = 0;
					System.out.println("===== [i] means No.i colume =====");
					while((tempStr = bufferedReader.readLine()) != null){
						String[] tempStrBox = tempStr.split(",");
						while (col < tempStrBox.length) {
							System.out.print( "[" + col + "]" + tempStrBox[col] + " ");
							col++;
						}
						System.out.println("");
						col = 0;
					}
					f_t.delete();
					read.close();	
				}
				catch (Exception ecp) {
					System.out.println("Error: Fail to open file.");					
				}	
			}
			else if (args[0].equals("--show-csv")) {
				try {
					File f_t = new File(args[1]);
				//System.out.println(f_t.toString());
					InputStreamReader read = new InputStreamReader (new FileInputStream(f_t), "UTF-8");
					BufferedReader bufferedReader = new BufferedReader(read);
					String tempStr;
					int col = 0; int row = 0;
					System.out.println("===== [i] means No.i colume =====");
					while((tempStr = bufferedReader.readLine()) != null){
						String[] tempStrBox = tempStr.split(",");
						while (col < tempStrBox.length) {
							System.out.print( "[" + col + "]" + tempStrBox[col] + " ");
							col++;
						}
						System.out.println("");
						col = 0;
					}
					read.close();
				}	
				catch (Exception ecp) {
					System.out.println("Error: Fail to open file.");
				}		
			}
			else if (args[0].equals("--convert-xls")) {
				try {
					ToCSV source = new ToCSV();
					source.convertExcelToCSV(args[1], ".");
					System.out.println("File converted.");
				}
				catch (Exception ecp) {
					System.out.println("Error: Fail to convert.");
				}
			}
			else if (args[0].equals("--convert-csv")) {
				try {
					ToExcel source = new ToExcel();
					String[] temp = args[1].split("\\.");
					source.convertCSVToExcel(args[1], temp[0] + ".xls");
					System.out.println("File converted.");
				}
				catch (Exception ecp) {
					System.out.println("Error: Fail to convert.");
				}
			}
			else {
				help();
			}
		}
		catch (Exception ecp) {
			help();
		}
	}
}