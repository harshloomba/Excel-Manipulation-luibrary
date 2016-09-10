using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using java.io;
using org.apache.poi.ss.usermodel;
using org.apache.poi.xssf.usermodel;

namespace PantheonProject
{
	public class test
	{
		public static void testMethod()
		{
			
			InputStream ExcelFileToRead = new FileInputStream("/Users/harshloomba/Documents/workspace/PantheonProject/source.xlsx");
			XSSFWorkbook hssfwb = new XSSFWorkbook(ExcelFileToRead);


			Sheet sheet = hssfwb.getSheetAt(0);
			Row row = sheet.getRow(4);

			for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++)
			{
				System.Console.Write(row.getCell(i)+" ");
			}
			row.getCell(0).setCellValue("Sample");
			//sheet.CreateRow(row.LastCellNum);
			Cell cell = row.createCell(row.getLastCellNum());
			Cell cell1 = row.createCell(row.getLastCellNum()+1);
			cell.setCellValue("test");
			cell1.setCellValue("test1");
			System.Console.WriteLine("after adding");
			for (int i = 0; i < row.getLastCellNum()+1; i++)
			{
				System.Console.WriteLine(row.getCell(i));
			}

			FileOutputStream fileOut = new FileOutputStream("/Users/harshloomba/Documents/workspace/PantheonProject/source3.xlsx");
			hssfwb.write(fileOut);
			fileOut.flush();
			fileOut.close();
		}
	}
}

