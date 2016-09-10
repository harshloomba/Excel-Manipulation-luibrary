using System;
using System.Collections.Generic;
using System.IO;
using ikvm.extensions;
using java.io;
using java.util;
using org.apache.poi.ss.usermodel;
using org.apache.poi.xssf.usermodel;
using log4net;
using log4net.Config;

namespace PantheonProject
{
	// Input of set names 
	// file name as paramenter 
	class PSIManipulation
	{
		private static readonly ILog logger = LogManager.GetLogger(typeof(DatabaseConnection));

		private static Dictionary<String, List<CodesDescription>> dataSource = new Dictionary<String, List<CodesDescription>>();

		public static void readCodesDataIntoDataBase(String s)
		{
			//dynamic property file will be read. input set(search), file path
			string home = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
			java.io.File ExcelFileToRead = new java.io.File(home + java.io.File.separator + "Documents" + java.io.File.separator + "Pantheon" + java.io.File.separator + "source1.xlsx");
		
			FileInputStream inputStream = new FileInputStream(ExcelFileToRead);

			XSSFWorkbook wb = new XSSFWorkbook(inputStream);
			XSSFSheet sheet = wb.getSheetAt(0);
			XSSFRow row;
			XSSFCell cell;

			Iterator rows = sheet.rowIterator();
			//int count = 0;
			while (rows.hasNext())
			{
				row = (XSSFRow)rows.next();
				Iterator cells = row.cellIterator();

				if (cells.hasNext())
				{
					cell = (XSSFCell)cells.next();
					if (cell.getStringCellValue().Equals(s))
					{
						XSSFCell code = (XSSFCell)cells.next();
						XSSFCell description = (XSSFCell)cells.next();

						CodesDescription ob = new CodesDescription();
						ob.setCodes(code.getStringCellValue());
						ob.setDescription(description.getStringCellValue());

						if (dataSource.ContainsKey(s))
						{
							dataSource[s].Add(ob);
						}
						else {
							List<CodesDescription> temp = new List<CodesDescription>();
							temp.Add(ob);
							dataSource.Add(s, temp);
						}
					}
				}
			}
		}


		public static void updateCodeData(int j, String s)
		{
			int rowCount = -1;

			string home = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
			java.io.File ExcelFileToRead = new java.io.File(home + java.io.File.separator + "Documents" + java.io.File.separator + "Pantheon" + java.io.File.separator + "PSI03Excel.xlsx");
			FileInputStream inputStream = new FileInputStream(ExcelFileToRead);
			XSSFWorkbook wb = new XSSFWorkbook(inputStream);
				XSSFSheet sheet = wb.getSheetAt(j);
				Dictionary<String, List<CodesDescriptionWithPos>> tempMap = new Dictionary<String, List<CodesDescriptionWithPos>>();
				XSSFRow row;
				XSSFCell cell;
				int firstRemovePoint = 0;
				CellStyle st = null;
				for (int m = 0; m < sheet.getPhysicalNumberOfRows(); m++)
				{
					rowCount++;
					row = sheet.getRow(rowCount);
					if (row == null)
						break;
					System.Console.WriteLine("excel row number: " + row.getRowNum() + "rowcount: " + rowCount);
					Iterator cells = row.cellIterator();
					cell = (XSSFCell)cells.next();

					if (cell.getStringCellValue().contains(s))
					{
						rowCount = rowCount + 2;
						List<CodesDescription> temp = dataSource[s];
						int len = temp.Count / 2;

						for (int i = 0; i < 2 * len; i = i + 2)
						{
							CodesDescription t1 = temp[i];
							CodesDescription t2 = temp[i + 1];
							System.Console.WriteLine(t1.toString() + " " + t2.toString());
							//System.out.println("excel row number: "+row.getRowNum()+"rowcount: "+rowCount);
							row = sheet.getRow(rowCount);
							if (row == null)
							{
								if (tempMap[s] == null)
								{
									List<CodesDescriptionWithPos> tempOb = new List<CodesDescriptionWithPos>();
									CodesDescriptionWithPos tem = new CodesDescriptionWithPos();
									tem.setCodes(t1.getCodes());
									tem.setDescription(t1.getDescription());
									tem.setPos(rowCount);
									tempOb.Add(tem);
									CodesDescriptionWithPos tem1 = new CodesDescriptionWithPos();
									tem1.setCodes(t2.getCodes());
									tem1.setDescription(t2.getDescription());
									tem1.setPos(rowCount);
									tempOb.Add(tem1);
									tempMap.Add(s, tempOb);
								}
								else {
									CodesDescriptionWithPos tem = new CodesDescriptionWithPos();
									tem.setCodes(t1.getCodes());
									tem.setDescription(t1.getDescription());
									tem.setPos(rowCount);
									tempMap[s].Add(tem);
									CodesDescriptionWithPos tem1 = new CodesDescriptionWithPos();
									tem1.setCodes(t2.getCodes());
									tem1.setDescription(t2.getDescription());
									tem1.setPos(rowCount);

									tempMap[s].Add(tem1);
								}
								continue;
							}

							//cells = row.cellIterator();
							//cell=(XSSFCell)cells.next();
							cell = row.getCell(0);

							//st=cell.getCellStyle();
							cell.setCellValue(t1.getCodes());
							cell.setCellStyle(st);

							//cell=(XSSFCell)cells.next();
							cell = row.getCell(1);
							//st=cell.getCellStyle();
							cell.setCellValue(t1.getDescription());
							cell.setCellStyle(st);

							row.getCell(3).setCellValue("");
							row.getCell(4).setCellValue("");

							//cells.next();

							//cell=(XSSFCell)cells.next();
							cell = row.getCell(3);
							//st=cell.getCellStyle();
							cell.setCellValue(t2.getCodes());
							cell.setCellStyle(st);

							//cell=(XSSFCell)cells.next();
							cell = row.getCell(4);
							//st=cell.getCellStyle();
							cell.setCellValue(t2.getDescription());
							cell.setCellStyle(st);
							st = cell.getCellStyle();
							rowCount++;
						}

						for (int i = 2 * len; i < temp.Count; i++)
						{

							CodesDescription t1 = temp[i];
							row = sheet.getRow(rowCount);
							System.Console.WriteLine(t1.toString());
							System.Console.WriteLine("rowCount: " + rowCount);
							if (row == null)
							{
								if (tempMap[s] == null)
								{
									List<CodesDescriptionWithPos> tempOb = new List<CodesDescriptionWithPos>();
									CodesDescriptionWithPos tem = new CodesDescriptionWithPos();
									tem.setCodes(t1.getCodes());
									tem.setDescription(t1.getDescription());
									tem.setPos(rowCount);
									tempOb.Add(tem);
									tempMap.Add(s, tempOb);
								}
								else {
									CodesDescriptionWithPos tem = new CodesDescriptionWithPos();
									tem.setCodes(t1.getCodes());
									tem.setDescription(t1.getDescription());
									tem.setPos(rowCount);
									tempMap[s].Add(tem);
								}
								continue;
							}

							cell = row.getCell(0);
							//st=cell.getCellStyle();
							cell.setCellValue(t1.getCodes());
							cell.setCellStyle(st);

							cell = row.getCell(1);
							//st=cell.getCellStyle();
							cell.setCellValue(t1.getDescription());
							cell.setCellStyle(st);
							st = cell.getCellStyle();

							row.getCell(3).setCellValue("");
							row.getCell(4).setCellValue("");
							rowCount++;
						}

						if (tempMap.Count == 0 && sheet.getRow(rowCount) != null)
						{
							firstRemovePoint = rowCount;
							//call delete method to modify the excel
							delete(sheet, firstRemovePoint);
						}

						if (tempMap.Count != 0)
						{
							createRows(tempMap, sheet, st, s);
						}


					}

				}

				//earlier I had delete and create function which I have separated in 2 different methods	

				FileOutputStream fileOut = new FileOutputStream(ExcelFileToRead);
				wb.write(fileOut);
				fileOut.flush();
				fileOut.close();
			
		}

		public static void delete(XSSFSheet sheet, int firstRemovePoint)
		{
			//delete if extra cells are there
			while (sheet.getRow(firstRemovePoint) != null)
			{
				removeRow(sheet, firstRemovePoint);
			}
		}

		public static void createRows(Dictionary<String, List<CodesDescriptionWithPos>> tempMap, XSSFSheet sheet, CellStyle st,String s)
		{
			XSSFRow row = null;
			XSSFCell cell = null;

			foreach (String c in tempMap.Keys)
				logger.Info("key in hashmap: "+c);
				
			List<CodesDescriptionWithPos> cachedValues = tempMap[s];
			int cellCount = 0;
			int rowCountTemp = 0;
			for (int i = 0; i < cachedValues.Count; i++)
			{
				CodesDescriptionWithPos t = cachedValues[i];
			
				if (rowCountTemp == 0)
					rowCountTemp = t.getPos();

				if (cellCount > 4)
				{
					cellCount = 0;
					rowCountTemp++;
				}
				row = sheet.getRow(rowCountTemp);
				if (row == null)
					row = sheet.createRow(rowCountTemp);
				row.createCell(cellCount);
				cell = row.getCell(cellCount);
				cell.setCellStyle(st);
				cell.setCellValue(t.getCodes());
				cellCount++;
				row.createCell(cellCount);
				cell = row.getCell(cellCount);
				cell.setCellStyle(st);
				cell.setCellValue(t.getDescription());
				cellCount = cellCount + 2;

				if (sheet.getRow(rowCountTemp + 1) != null)
				{
					sheet.shiftRows(rowCountTemp + 1, sheet.getPhysicalNumberOfRows(), +1);
				}
			}
		}

		public static void removeRow(XSSFSheet sheet, int rowIndex)
		{
			int lastRowNum = sheet.getPhysicalNumberOfRows();
			if (rowIndex >= 0 && rowIndex < lastRowNum)
			{
				sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
			}
			if (rowIndex == lastRowNum)
			{
				XSSFRow removingRow = sheet.getRow(rowIndex);
				if (removingRow != null)
				{
					sheet.removeRow(removingRow);
				}
			}

		}
	}
}
