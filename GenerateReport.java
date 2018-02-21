package myclasses;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;

public class GenerateReport {

	public static Workbook book = null;
	public static Sheet sheet = null;
	public static String filePath = "D:\\working\\POC\\RunManager.xls";
	public static String reportsLoc = "D:\\working\\POC\\Reports.html";
	public static StringBuffer htmlTable = null;
	public static int stepsCount = 0;
	private static LinkedHashMap<String, String> testSuites = new LinkedHashMap<String, String>();
	
	public static void buildTestSuiteSummaryReport()
	{
		try
		{
			FileInputStream fis = new FileInputStream(new File(filePath));
			WorkbookSettings ws = new WorkbookSettings();
			ws.setSuppressWarnings(true);			
			book = Workbook.getWorkbook(fis,ws);
			sheet = book.getSheet(0);
			htmlTable = new StringBuffer();
			
			int totalRows = sheet.getRows();
			int totalCols = sheet.getColumns();
			int totalSuites = 0;
			int totalPass = 0;
			int totalFail = 0;
			int totalTimeTaken = 0; 
			int sno = 0;
			
			htmlTable.append("<html><head><title>::API Automation::Reports</title><script type='text/javascript'>function showTestSuiteSummary(id)"
					+ "{var e = document.getElementById(id); if (e.style.display == 'none'){e.style.display='block';}"
					+ "else{e.style.display='none'}}function showHideTestCaseSummary(divId){var dvId = document.getElementById(divId);"
					+ "if (dvId.style.display == 'none'){dvId.style.display='block';}else{dvId.style.display='none';}}"
					+ "function showTestCaseReport(testCaseId, testSteps, tsStatus){"
					+ "var steps = new Array(); steps = testSteps.split('|');" 
					+ "var tStatus = new Array(); tStatus = tsStatus.split('|');"
					+ "win=window.open('','_blank', 'width=700, height=200, top=100, left=200, resizeable, scrollbars');"
					+ "win.document.write('<title>::API Automation::Testcase Results</title>"
					+ "<table style=color:black;font-family:verdana;font-size:12px; border=1 "
					+ "cellpadding=4 cellspacing=0><tr style=background-color:coral><td align=center>' + testCaseId + '</td>');"
					+ "for (stp=0; stp<steps.length-1; stp++){win.document.write('<td align=center>' + steps[stp] + '</td>');}"
					+ "win.document.write('</tr><tr>');"
					+ "for (idx=1; idx < tStatus.length-1; idx++)"
					+ "{if (tStatus[idx] == 'FAIL'){"
					+ "win.document.write('<td align=center><img src=D:/working/POC/fail.png alt=Fail style=width:15px;height:15px;>');}"
					+ "else if (tStatus[idx] == 'PASS'){"
					+ "win.document.write('<td align=center><img src=D:/working/POC/pass.png alt=Pass style=width:15px;height:15px;>');}"
					+ "else{win.document.write('<td>&nbsp;</td>');}}"
					+ "win.document.write('</tr></table><p align=center style=font-family:verdana;font-size:12px;>"
					+ "<a href=# style=text-decoration:none;color:blue title=Close Window onClick=javascript:window.close()>"
					+ "Close Window</a></p>');win.document.close();}"
					+ "</script></head><body>");
			htmlTable.append("<h2 style='font-family:verdana;font-size:18px;'>Report Summary</h2>"
					+ "<table border='1' id='testSuiteTable' cellpadding='4' cellspacing='0'>");
			htmlTable.append("<thead><tr style='background-color:darkblue;color:white;font-family:verdana;font-size:14px;'>");
			htmlTable.append("<th>Sno</th><th># Test Suites</th>");
			htmlTable.append("<th># Passed</th><th># Failed</th>");
			htmlTable.append("<th>Total Time Taken</th></tr></thead>"
					+ "<tbody style='background-color:white;color:black;font-family:verdana;font-size:12px;'>");
			
			// This loop will give the total testsuites, totalpass/fail count and totalTimeTaken
			for (int baseRow = 1; baseRow < totalRows; baseRow++)
			{
				// Check if the flag is 'Y'
				if (sheet.getCell(1,baseRow).getContents().equalsIgnoreCase("Y"))
				{
					totalSuites++;

					// Read the row data where the 'Y' is enabled
					for (int col = 0; col < totalCols; col++)
					{
						String strStatus = sheet.getCell(col,baseRow).getContents();
						if (strStatus.equalsIgnoreCase("pass"))
						{
							totalPass++;
						}
						else if (strStatus.equalsIgnoreCase("fail"))
						{
							totalFail++;
						}
					}

					// Store TestSuite Name and Execution Time
					testSuites.put(sheet.getCell(0,baseRow).getContents(), sheet.getCell(3,baseRow).getContents());
					
					totalTimeTaken += Integer.parseInt(sheet.getCell(3,baseRow).getContents()); // Execution Time
				}
			}

			// Add a table row in the HTML table along with appropriate data 
			htmlTable.append("<tr>");
			htmlTable.append("<td align='center'>1</td>");
			htmlTable.append("<td align='center'>"
					+ "<a href='#' title='click here' style='text-decoration:none;color:blue' onclick=showTestSuiteSummary('suiteDiv')>" 
					+ totalSuites + "</a></td>");
			htmlTable.append("<td align='center'>" + totalPass + "</td>");
			htmlTable.append("<td align='center'>" + totalFail + "</td>");
			htmlTable.append("<td align='center'>" + totalTimeTaken + " (seconds)</td>");
			htmlTable.append("</tr></tbody></table><br>");
			htmlTable.append("<div id='suiteDiv' style='display:none'>"
					+ "<h3 style='font-family:verdana;font-size:18px'>Testsuite Summary Report</h3>"
					+ "<table border='1' cellpadding='4' cellspacing='0'>");
			
			htmlTable.append("<tr style='background-color:darkblue;color:white;font-family:verdana;font-size:14px;'>"
					+ "<th align='center'>Sno</th><th align='center'>Test Suite Name</th>");
			htmlTable.append("<th align='center'># TestCases</th><th align='center'># Passed</th>");
			htmlTable.append("<th align='center'># Failed</th><th align='center'>Time Taken (in seconds)</th></tr>");
			
			// Build 'Test Suite' Summary table
			for (Map.Entry<String, String> entry : testSuites.entrySet())
			{
				String suiteName = "";
				int totalTestCases = 0;
				String totalTime = "";
				totalPass = 0;
				totalFail = 0;
				
				if (entry !=null)
				{
					suiteName = entry.getKey();
					totalTime = entry.getValue();
					sno++;
					
					// Test case sheet
					Sheet testSheet = book.getSheet(suiteName);
					
					int rows = testSheet.getRows();
					int cols = testSheet.getColumns();
					
					// This loop reads the test case sheet which has the flag 'Y'
					for (int row = 1; row < rows; row++)
					{
						// Pick the test cases whose flag is 'Y'
						if (testSheet.getCell(1, row).getContents().equalsIgnoreCase("Y"))
						{
							totalTestCases++;
							
							// Read the row data where the 'Y' is enabled
							for (int col = 0; col < cols; col++)
							{
								String strValue = testSheet.getCell(col,row).getContents();
								
								if (strValue.equalsIgnoreCase("pass"))
								{
									totalPass++;
								}
								else if (strValue.equalsIgnoreCase("fail"))
								{
									totalFail++;
								}
							}
						}
					}
					
					// Add row
					htmlTable.append("<tr style='background-color:white;color:black;font-family:verdana;font-size:12px;'>");
					htmlTable.append("<td align='center'>" + sno + "</td>");
					htmlTable.append("<td align='center'><a href='#' style='text-decoration:none;color:blue' title='click " + suiteName 
							+ "' onclick=showHideTestCaseSummary('" + suiteName + "')>" + suiteName + "</a></td>");
					htmlTable.append("<td align='center'>" + totalTestCases + "</td>");
					htmlTable.append("<td align='center'>" + totalPass + "</td>");
					htmlTable.append("<td align='center'>" + totalFail + "</td>");
					htmlTable.append("<td align='center'>" + totalTime + "</td>");
					htmlTable.append("</tr>");
				}
			}
			
			htmlTable.append("</table></div><br><br>"); // TestSuite table ends here
			
			// Build Testcase Summary Report
			for (Map.Entry<String, String> entry : testSuites.entrySet())
			{
				if (entry != null)
				{
					Sheet tcSheet = book.getSheet(entry.getKey());
					
					// Table for Testcase Summary report for each Testsuite
					htmlTable.append("<div id='" + entry.getKey() + "' style='display:none'><table border='1' cellpadding='4' cellspacing='0'>");
					htmlTable.append("<tr style='background-color:skyblue;color:black;font-family:verdana;font-size:14px'>"
							+ "<th colspan='5' align='left'>Test Suite Name : " + entry.getKey() + "</th><tr>");
					
					htmlTable.append("<tr style='background-color:darkblue;color:white;font-family:verdana;font-size:14px'>");
					htmlTable.append("<th align='center'>Sno</th><th align='center'>TestCase</th>");
					htmlTable.append("<th align='center'>Steps</th><th align='center'>Status</th>");
					htmlTable.append("<th align='center'>Detail</th></tr>");
					
					int rows = tcSheet.getRows();
					int cols = tcSheet.getColumns();
					List<String> strSteps = null;
					
					for (int row = 1; row < rows; row++)
					{
						strSteps = new ArrayList<String>();
						String tcId = "";
						
						if (tcSheet.getCell(1,row).getContents().equalsIgnoreCase("Y"))
						{
							tcId = tcSheet.getCell(0,row).getContents();
							htmlTable.append("<tr style='background-color:white;color:black;font-family:verdana;font-size:12px'>");
							htmlTable.append("<td align='center'>" + row + "</td>");
							htmlTable.append("<td align='center'>" + tcId + "</td>");
						
							for (int col = 0; col < cols; col++)
							{
								String strValue = tcSheet.getCell(col,row).getContents();
							
								if (strValue.contains("Request"))
								{
									strSteps.add(strValue);
								}
							}
						
							stepsCount = strSteps.size();
							String strStep = "";
							Iterator itr = strSteps.iterator();
							while (itr.hasNext())
							{
								strStep += itr.next() + "|";
							}
							htmlTable.append("<td align='center'>" + strStep + "</td>");
							if (tcSheet.getCell(2,row).getContents().equalsIgnoreCase("PASS"))
							{
								htmlTable.append("<td align='center' style='color:green;font-weight:bold'>" + tcSheet.getCell(2,row).getContents() + "</td>");
							}
							else if (tcSheet.getCell(2,row).getContents().equalsIgnoreCase("FAIL"))
							{
								htmlTable.append("<td align='center' style='color:red;font-weight:bold'>" + tcSheet.getCell(2,row).getContents() + "</td>");
							}
							htmlTable.append("<td align='center'><a href='#' title='click here'" 
									+ " style='text-decoration:none;color:blue' onclick=\"showTestCaseReport('" 
									+ tcId + "','" + strStep + "','" + getTestStepsByCaseId(tcId) + "')\">Detail</a></td></tr>");
						}
					}
					htmlTable.append("</table></div>");
				}
			}
			
			htmlTable.append("</body></html>");
			fis.close();
			book.close();
			
			FileWriter writer = new FileWriter(new File(reportsLoc));
			writer.write(htmlTable.toString());
			writer.flush();
			writer.close();
			testSuites.clear();
			htmlTable = null;
			
		}
		catch(Exception ex)
		{
			ex.printStackTrace();
		}
	}
	
	public static String getTestStepsByCaseId(String tstId)
	{
		//String[] steps = new String[stepsCount];
		String steps = "";
		
		try
		{
			WorkbookSettings ws = new WorkbookSettings();
			ws.setSuppressWarnings(true);
			FileInputStream fis = new FileInputStream(new File(filePath));
			Workbook myBook = Workbook.getWorkbook(fis,ws);
			Sheet mySheet = myBook.getSheet("TestSuiteSummary");
			
			int rows = mySheet.getRows();
			int cols = mySheet.getColumns();

			for (int r = 0; r < rows; r++)
			{
				String tstCaseId = mySheet.getCell(1, r).getContents();

				if (tstCaseId.equalsIgnoreCase(tstId))
				{
					for (int c = 0; c < cols; c++)
					{
						steps += mySheet.getCell(c, r).getContents() + "|";
					}
					break;
				}
			}
			fis.close();
			myBook.close();
		}
		catch(Exception ex)
		{
			ex.printStackTrace();
		}
		return steps;
	}
	
	public static void main(String args[])
	{
		buildTestSuiteSummaryReport();
	}
}