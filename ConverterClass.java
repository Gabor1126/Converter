package Code;

import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.Button;

import java.awt.Component;
import java.awt.Dimension;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.UIManager;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.eclipse.swt.SWT;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;
import org.eclipse.swt.graphics.Rectangle;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.Monitor;
import org.eclipse.swt.widgets.Text;

public class ConverterClass {

	protected Shell shell;
	private File fileReadPlace, fileWritePlace;
	private Text txtFileName;

	/**
	 * Launch the application.
	 * 
	 * @param args
	 */
	public static void main(String[] args) {
		try {
			ConverterClass window = new ConverterClass();
			window.open();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * Open the window.
	 */
	public void open() {
		Display display = Display.getDefault();
		createContents();
		Monitor primary = display.getPrimaryMonitor();
		Rectangle bounds = primary.getBounds();
		Rectangle rect = shell.getBounds();

		int x = bounds.x + (bounds.width - rect.width) / 2;
		int y = bounds.y + (bounds.height - rect.height) / 2;

		shell.setLocation(x, y);
		shell.open();
		shell.layout();
		while (!shell.isDisposed()) {
			if (!display.readAndDispatch()) {
				display.sleep();
			}
		}
	}

	/**
	 * Create contents of the window.
	 */
	protected void createContents() {
		shell = new Shell();
		shell.setSize(450, 300);
		shell.setText("Converter");

		Label lblVlasszaKiA = new Label(shell, SWT.NONE);
		lblVlasszaKiA.setBounds(37, 24, 192, 15);
		lblVlasszaKiA.setText("Válassza ki a beolvasandómappát:");

		Label lblVlasszaKiA_1 = new Label(shell, SWT.NONE);
		lblVlasszaKiA_1.setBounds(37, 91, 159, 15);
		lblVlasszaKiA_1.setText("Válassza ki a cél mappát:");

		Label lblReadPlace = new Label(shell, SWT.NONE);
		lblReadPlace.setBounds(37, 45, 339, 40);

		Label lblWritePlace = new Label(shell, SWT.NONE);
		lblWritePlace.setBounds(37, 112, 339, 50);
		Label lblAdjaMegA = new Label(shell, SWT.NONE);
		lblAdjaMegA.setBounds(37, 175, 192, 15);
		lblAdjaMegA.setText("Adja meg a létrehozandó fájl nevét:");

		txtFileName = new Text(shell, SWT.BORDER);
		txtFileName.setBounds(235, 172, 115, 21);

		Label lblxls = new Label(shell, SWT.NONE);
		lblxls.setBounds(353, 175, 55, 15);
		lblxls.setText(".xls");

		Button btnDif = new Button(shell, SWT.NONE);
		btnDif.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {

				JFileChooser chooser = new JFileChooser();
				chooser.setCurrentDirectory(new java.io.File("D:\\"));
				chooser.setDialogTitle("choosertitle");
				chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				chooser.setAcceptAllFileFilterUsed(false);

				if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
					fileReadPlace = chooser.getSelectedFile();
					lblReadPlace.setText("A beolvandó mappa:\n" + fileReadPlace.toString());
				}

			}
		});
		btnDif.setBounds(231, 19, 119, 25);
		btnDif.setText("DIF-ek mappája");

		Button btnExcel = new Button(shell, SWT.NONE);
		btnExcel.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				JFileChooser chooser = new JFileChooser();
				chooser.setCurrentDirectory(new java.io.File("D:\\"));
				chooser.setDialogTitle("choosertitle");
				chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				chooser.setAcceptAllFileFilterUsed(false);

				if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
					fileWritePlace = chooser.getSelectedFile();
					lblWritePlace.setText("A beolvandó mappa:\n" + fileWritePlace.toString());
				}

			}
		});
		btnExcel.setBounds(231, 86, 119, 25);
		btnExcel.setText("Excel mappája");

		Button btnExecute = new Button(shell, SWT.NONE);
		btnExecute.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {

				if (fileReadPlace != null && fileWritePlace != null && txtFileName.getText() != "") {
					Workbook wb = new HSSFWorkbook();
					CreationHelper helper = wb.getCreationHelper();
					Sheet sheet = wb.createSheet("new sheet");

					System.setProperty("file.encoding", "UTF-8");
					Field charset = null;
					try {
						charset = Charset.class.getDeclaredField("defaultCharset");
						charset.setAccessible(true);
					} catch (NoSuchFieldException ex) {
						ex.printStackTrace();
					} catch (SecurityException ex) {
						ex.printStackTrace();
					}
					try {
						charset.set(null, null);
					} catch (IllegalArgumentException ex) {
						ex.printStackTrace();
					} catch (IllegalAccessException ex) {
						ex.printStackTrace();
					}

					String[] types = new String[] { "NUMBER\\({1}[0-9]++,{1}[0-9]++\\){1}",
							"TIMESTAMP\\({1}[0-9]++\\){1}", "VARCHAR[0-9]++\\({1}[0-9]++\\){1}", "INTEGER", "DATE",
							"UNIQ" };

					Matcher tableMatcher1 = null, tableMatcher2 = null, rowsMatcher = null, bodyMatcher = null,
							oumMatcher = null;
					Pattern tablePattern1 = null, tablePattern2 = null, rowsPattern = null, bodyPatter = null,
							oumPattern = null;
					String[] parts = null;
					String line;
					int r = 0;
					String table = null;
					String[] rowsCut = new String[] { "#INDEX.*", "#CONSTRAINT.*" };

					File[] files = fileReadPlace.listFiles();

					for (File file : files) {
						boolean firstRow = true;

						Scanner sc = null;
						try {
							sc = new Scanner(file);
						} catch (FileNotFoundException e2) {
							e2.printStackTrace();
						}

						while (sc.hasNext()) {
							line = sc.nextLine();

							if (firstRow == true) {
								String tableCash = null;
								tablePattern1 = Pattern.compile("TABLE");
								tableMatcher1 = tablePattern1.matcher(line);
								if (tableMatcher1.find()) {
									tableCash = line.replace(tableMatcher1.group(), "");
								}
								tablePattern2 = Pattern.compile("[A-Z](.*)");
								tableMatcher2 = tablePattern2.matcher(tableCash);
								if (tableMatcher2.find()) {
									table = tableCash.replace(tableMatcher2.group(), "");
								}

								firstRow = false;
							} else {
								String typeValue;
								Row row = sheet.createRow((short) r);
								line = line.replace("FIELD", "");

								for (String rowCut : rowsCut) {
									rowsPattern = Pattern.compile(rowCut);
									rowsMatcher = rowsPattern.matcher(line);
									if (rowsMatcher.find()) {
										line = line.replace(rowsMatcher.group(), "");
									}
								}
								if (!line.equals("")) {
									r++;
									row.createCell(0).setCellValue(helper.createRichTextString(table));
								}

								for (String type : types) {
									bodyPatter = Pattern.compile(type);
									bodyMatcher = bodyPatter.matcher(line);
									if (bodyMatcher.find()) {
										parts = line.split(type);
										typeValue = bodyMatcher.group();
										row.createCell(1).setCellValue(helper.createRichTextString(parts[0]));
										row.createCell(2).setCellValue(helper.createRichTextString(typeValue));
										row.createCell(3).setCellValue(helper.createRichTextString(parts[1]));
										break;
									} else {
										row.createCell(1).setCellValue(helper.createRichTextString(line));
									}
								}

								oumPattern = Pattern.compile("ÜOM(.*)");
								oumMatcher = oumPattern.matcher(line);
								if (oumMatcher.find()) {
									row.createCell(4).setCellValue(helper.createRichTextString(oumMatcher.group()));
								}

							}

						}

						FileOutputStream fileOut = null;
						try {
							fileOut = new FileOutputStream(
									fileWritePlace.toString() + "/" + txtFileName.getText() + ".xls");
							wb.write(fileOut);
							fileOut.close();
						} catch (FileNotFoundException e1) {
							e1.printStackTrace();
						} catch (IOException e1) {
							e1.printStackTrace();
						}
					}
					JOptionPane.showMessageDialog(null, "Sikeres konvertálás!");
				} else {
					JOptionPane.showMessageDialog(null,
							"Nem választotta ki az egyik mappát vagy nem adott meg fájl nevet!");
				}
			}
		});
		btnExecute.setBounds(98, 213, 75, 25);
		btnExecute.setText("Végrehajt");

		Button btnClose = new Button(shell, SWT.NONE);
		btnClose.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
				shell.close();

			}
		});

		btnClose.setBounds(249, 213, 75, 25);
		btnClose.setText("Bezár");

	}
}
