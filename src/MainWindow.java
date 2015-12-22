/*
	A Java-based Excel using Apache POI
	Author: Bird Liu (Liu Xin)
	License: GNU GPLv2

	GitHub:
		https://www.github.com/JavaHomework
*/

import java.io.FileOutputStream;
import java.util.Date;
import javax.swing.*;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.TableModel;
import java.util.*;
import java.awt.*;
import java.awt.event.*;
import java.io.*;
import org.apache.poi.ss.examples.ToCSV;
import org.FRing.ToExcel;

class Cell implements Serializable {
	private int x, y, z;
	private String value;
	private Color color;

	public void setValue (String value) { this.value = value; }
	public void setX(int x) { this.x = x; }
	public void setY(int y) { this.y = y; }
	public void setColor(Color color) { this.color = color; }

	public String getValue() { return value; }
	public int getX() { return x; }
	public int getY() { return y; }
	public Color getColor() { return color; }
}

public class MainWindow {
	JFrame frame;
	Container con;
	JTable table;
	JMenuBar menubar;
	JMenu fileMenu, editMenu, aboutMenu;
	JMenuItem open, openxls, save, saveAs, newFile, quit;
	JMenuItem copy, cut, paste, search, color;
	JMenuItem about;
	JDialog dialog;
	Cell cell[] = new Cell[1000];
	MyActionListener myActionListener = new MyActionListener();
	Dimension screensize = Toolkit.getDefaultToolkit().getScreenSize();
	int screenWidth = (int) screensize.getWidth();
	int screenHeight = (int) screensize.getHeight();
	String clipboard = new String();

	public MainWindow() {
		/* Draw Frame */
		frame = new JFrame("EasyExcel(CSV) v1.0 Beta");
		con = frame.getContentPane();
		table = new JTable(50, 20);
		JScrollPane scrollpane = new JScrollPane(table);
		con.add(scrollpane);
		menubar = new JMenuBar();
		frame.setJMenuBar(menubar); // Add Menus to panel

		/* File Menu Items */
		fileMenu = new JMenu("File");
		open = new JMenuItem("Open ..");
		openxls = new JMenuItem("Open XLS..");
		save = new JMenuItem("Save .");
		saveAs = new JMenuItem("Save XLS..");
		newFile = new JMenuItem("New File");
		quit = new JMenuItem("Quit");
		save.addActionListener(myActionListener);
		saveAs.addActionListener(myActionListener);
		newFile.addActionListener(myActionListener);
		open.addActionListener(myActionListener);
		openxls.addActionListener(myActionListener);
		quit.addActionListener(myActionListener);
		fileMenu.add(open);
		fileMenu.add(openxls);
		fileMenu.add(save);
		fileMenu.add(saveAs);
		fileMenu.add(newFile);
		fileMenu.add(newFile);
		fileMenu.add(quit);
		fileMenu.addSeparator(); 
		/* Edit Menu Items */
		editMenu = new JMenu("Edit");
		copy = new JMenuItem("Copy");
		cut = new JMenuItem("Cut");
		paste = new JMenuItem("Paste");
		search = new JMenuItem("Search");
		color = new JMenuItem("Color");
		color.addActionListener(myActionListener);
		copy.addActionListener(myActionListener);
		cut.addActionListener(myActionListener);
		paste.addActionListener(myActionListener);
		search.addActionListener(myActionListener);
		editMenu.add(copy);
		editMenu.add(cut);
		editMenu.add(paste);
		editMenu.add(search);
		editMenu.add(color);
		/* About Menu Items */
		aboutMenu = new JMenu("About");
		about = new JMenuItem("About EasyExcel");
		about.addActionListener(myActionListener);
		aboutMenu.add(about);
		/* Add to panel */
		menubar.add(fileMenu);
		menubar.add(editMenu);
		menubar.add(aboutMenu);

		frame.setBounds(screenWidth / 2 - 1028 / 2, screenHeight / 2 - 526 / 2, 1028, 526);
		frame.setVisible(true);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

		NewFile();
	}

	private class MyActionListener implements ActionListener {
		public void actionPerformed(ActionEvent ae) {
			if (ae.getSource() == newFile) {
				NewFile();
			}
			if (ae.getSource() == save) {
				Save();
			}
			if (ae.getSource() == saveAs) {
				SaveAs();
			}
			if (ae.getSource() == open) {
				Open();
			}
			if (ae.getSource() == openxls) {
				OpenXls();
			}
			if (ae.getSource() == quit) {
				System.exit(0);
			}
			if (ae.getSource() == copy) {
				Copy();
			}
			if (ae.getSource() == cut) {
				Cut();
			}
			if (ae.getSource() == paste) {
				Paste();
			}
			if (ae.getSource() == search) {
				Search();
			}
			if (ae.getSource() == color) {
				ColorChange();
			}
			if (ae.getSource() == about) {
				About();
			}
		}
	}

	public void Open() {
		try {
			JFileChooser dlg = new JFileChooser();
			dlg.setDialogTitle("Open CSV File");
			dlg.setFileFilter(new javax.swing.filechooser.FileFilter() {
            	public boolean accept(File f) {
                	if (f.getName().endsWith(".csv") || f.isDirectory()) {
                    	return true;
                	}
                	return false;
            	}
            	public String getDescription() {
                	return "Comma-Separated Values Files (*.csv)";
            	}
        	});			
			int result = dlg.showOpenDialog(frame); 
			if (result == JFileChooser.APPROVE_OPTION) {
				File f = dlg.getSelectedFile();
				InputStreamReader read = new InputStreamReader (new FileInputStream(f), "UTF-8");
				BufferedReader bufferedReader = new BufferedReader(read);
				NewFile();
				int col = 0; int row = 0;
				Cell temp = new Cell();
				String tempStr;
				while((tempStr = bufferedReader.readLine()) != null){
					String[] strBox = tempStr.split(",");
					while (col < strBox.length) {
						temp.setValue(strBox[col]);
						temp.setX(row); temp.setY(col);
						table.setValueAt(temp.getValue(), temp.getX(), temp.getY());
						col = col + 1;
					}
					col = 0;
					row = row + 1;
				}
				read.close();
			}
		} 
		catch (Exception ecp) {
			System.out.println("Error");
		}
	}

	public void OpenXls() {
		try {
			ToCSV source = new ToCSV();
			JFileChooser dlg = new JFileChooser();
			dlg.setFileFilter(new javax.swing.filechooser.FileFilter() {
            	public boolean accept(File f) {
                	if (f.getName().endsWith(".xls") || f.isDirectory()) {
                    	return true;
                	}
                	return false;
            	}
            	public String getDescription() {
                	return "Standard Microsoft Excel Files (*.xls)";
            	}
        	});
			dlg.setDialogTitle("Open XLS File");
			int result = dlg.showOpenDialog(frame); 
			if (result == JFileChooser.APPROVE_OPTION) {
				File f = dlg.getSelectedFile();
				source.convertExcelToCSV(f.toString(), ".");
				String tempName[] = f.getName().split("\\.");
				File f_t = new File(tempName[0] + ".csv");
				//System.out.println(f_t.toString());
				InputStreamReader read = new InputStreamReader (new FileInputStream(f_t), "UTF-8");
				BufferedReader bufferedReader = new BufferedReader(read);
				NewFile();
				int col = 0; int row = 0;
				Cell temp = new Cell();
				String tempStr;
				while((tempStr = bufferedReader.readLine()) != null){
					String[] strBox = tempStr.split(",");
					while (col < strBox.length) {
						temp.setValue(strBox[col]);
						temp.setX(row); temp.setY(col);
						table.setValueAt(temp.getValue(), temp.getX(), temp.getY());
						col = col + 1;
					}
					col = 0;
					row = row + 1;
				}
				read.close();
				f_t.delete();
			}
		} 
		catch (Exception ecp) {
			System.out.println("Error");
		}
	}

	public void Save() {
		try {
			JFileChooser dlg = new JFileChooser();
			dlg.setDialogTitle("Save CSV File");
			dlg.setFileFilter(new javax.swing.filechooser.FileFilter() {
            	public boolean accept(File f) {
                	if (f.getName().endsWith(".csv") || f.isDirectory()) {
                    	return true;
                	}
                	return false;
            	}
            	public String getDescription() {
                	return "Comma-Separated Values Files (*.csv)";
            	}
        	});				
			int result = dlg.showSaveDialog(frame);  // 打"开保存文件"对话框
			if (result == JFileChooser.APPROVE_OPTION) {
				File f = dlg.getSelectedFile();
				BufferedWriter writer = new BufferedWriter(new FileWriter(f, false));
				int row = 0; int col = 0;
				int final_row = 0;
				int flag = 0;
			
				for (int i = 49; i >= 0; i--) {
					for (int j = 0; j < 20; j++) {
						if (table.getValueAt(i, j).toString() != "") {
							final_row = i;
							flag = 1;
							//System.out.println(final_row);
							//System.out.println(table.getValueAt(i, j));
							break;
						}
					}
					if (flag == 1) {
						flag = 0;
						break;
					}
				}
			
				for (int i = 0; i <= final_row; i++) {
					String temp = new String();
					for (int j = 0; j < 20; j++) {
						String value = new String();
						if (table.getValueAt(i, j) == null) {
							value = "";	
						}
						else 
							value = table.getValueAt(i, j).toString();
						temp = temp + value + ",";
					}
					writer.write(temp + "\n");
				}
				writer.close();
				final_row = 0;
			} 
		} 
		catch (Exception ecp) {
			System.out.println("Error");
		}
	}

	public void SaveAs() {
		try {
			JFileChooser dlg = new JFileChooser();
			dlg.setDialogTitle("Save XLS File");
			dlg.setFileFilter(new javax.swing.filechooser.FileFilter() {
            	public boolean accept(File f) {
                	if (f.getName().endsWith(".xls") || f.isDirectory()) {
                    	return true;
                	}
                	return false;
            	}
            	public String getDescription() {
                	return "Standard Microsoft Excel Files (*.xls)";
            	}
        	});				
			int result = dlg.showSaveDialog(frame);  // 打"开保存文件"对话框
			if (result == JFileChooser.APPROVE_OPTION) {
				File f = dlg.getSelectedFile();
				BufferedWriter writer = new BufferedWriter(new FileWriter(f, false));
				int row = 0; int col = 0;
				int final_row = 0;
				int flag = 0;
			
				for (int i = 49; i >= 0; i--) {
					for (int j = 0; j < 20; j++) {
						if (table.getValueAt(i, j).toString() != "") {
							final_row = i;
							flag = 1;
							//System.out.println(final_row);
							//System.out.println(table.getValueAt(i, j));
							break;
						}
					}
					if (flag == 1) {
						flag = 0;
						break;
					}
				}
			
				for (int i = 0; i <= final_row; i++) {
					String temp = new String();
					for (int j = 0; j < 20; j++) {
						String value = new String();
						if (table.getValueAt(i, j) == null) {
							value = "";	
						}
						else 
							value = table.getValueAt(i, j).toString();
						temp = temp + value + ",";
					}
					writer.write(temp + "\n");
				}
				writer.close();
				final_row = 0;
				ToExcel source = new ToExcel();
				source.convertCSVToExcel(f.toString(), f.toString());
			} 
		} 
		catch (Exception ecp) {
			System.out.println("Error");
		}
	}

	public void NewFile() {
		for (int i = 0; i < 50; i++) {
			for (int j = 0; j < 20; j++) {
				table.setValueAt("", i, j);
			}
		}
	}

	public void About() {
		dialog = new JDialog(frame, "About EasyExcel", false);
		dialog.setLayout(new GridLayout(8, 10));
		dialog.getContentPane().add(new JLabel(" EasyExcel(CSV) is a Java-based software like Excel."));
		dialog.getContentPane().add(new JLabel(""));
		dialog.getContentPane().add(new JLabel(" It uses CSV format to work, and also support opening XLS files."));
		dialog.getContentPane().add(new JLabel(""));
		dialog.getContentPane().add(new JLabel(" Designer: Bird Liu (Liu Xin)"));
		dialog.getContentPane().add(new JLabel(" Student ID: 320130938311"));
		dialog.getContentPane().add(new JLabel(" Major: 2013 Information Security"));
		dialog.setBounds(screenWidth / 2 - 200 / 2, screenHeight / 2 - 100 / 2, 500, 100);
		dialog.setVisible(true);
	}

	public void Copy() {
		clipboard = table.getValueAt(table.getSelectedRow(), table.getSelectedColumn()).toString();
		System.out.println(clipboard);
	}

	public void Cut() {
		clipboard = table.getValueAt(table.getSelectedRow(), table.getSelectedColumn()).toString();
		table.setValueAt("", table.getSelectedRow(), table.getSelectedColumn());
		System.out.println(clipboard);		
	}

	public void Paste() {
		table.setValueAt(clipboard, table.getSelectedRow(), table.getSelectedColumn());
		System.out.println(clipboard);		
	}

	public void Search() {
		dialog = new JDialog(frame, "Search", false);
		dialog.setLayout(new FlowLayout());
		dialog.getContentPane().add(new JLabel("Target："));
		JTextField target = new JTextField(10);
		dialog.getContentPane().add(target);
		JButton yes = new JButton("确定");
		JButton cancel = new JButton("取消");
		dialog.add(yes);
		dialog.add(cancel);
		dialog.setBounds(screenWidth / 2 - 200 / 2, screenHeight / 2 - 100 / 2, 200, 100);
		dialog.setVisible(true);
		yes.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				for (int i = 0; i < 50; i++) {
					for (int j = 0; j < 20; j++) {
						if (table.getValueAt(i, j).toString().equals(target.getText())) {
						table.setRowSelectionInterval(i,i);
						table.setColumnSelectionInterval(j,j);
						}
					}
				}
				dialog.setVisible(false);
			}
		});
		cancel.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				dialog.setVisible(false);
			}
		});
	}

	public void ColorChange() {
		JColorChooser cc = new JColorChooser();
		dialog = new JDialog(frame, "Change Color", false);
		dialog.setLayout(new FlowLayout());
		dialog.getContentPane().add(cc);
		dialog.getContentPane().add(new JButton("OK"));
		dialog.setBounds(screenWidth / 2 - 700 / 2, screenHeight / 2 - 400 / 2, 700, 400);
		dialog.setVisible(true);
	}
}