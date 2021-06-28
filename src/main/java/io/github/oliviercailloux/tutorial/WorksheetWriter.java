package io.github.oliviercailloux.tutorial;

import static com.google.common.base.Preconditions.checkArgument;
import static com.google.common.base.Preconditions.checkNotNull;

import java.util.Iterator;
import java.util.List;

import com.google.gson.JsonPrimitive;
import com.microsoft.graph.http.CustomRequest;
import com.microsoft.graph.models.WorkbookRange;
import com.microsoft.graph.models.WorkbookRangeFill;
import com.microsoft.graph.models.WorkbookRangeFont;
import com.microsoft.graph.models.WorkbookRangeFormat;
import com.microsoft.graph.models.WorkbookWorksheet;
import com.microsoft.graph.models.WorkbookWorksheetRangeParameterSet;
import com.microsoft.graph.requests.GraphServiceClient;

/**
 * This class gathers basic methods that help writing an Excel online line
 */

public class WorksheetWriter {

	/**
	 * This function check if the given worksheet exist or not
	 * 
	 * @param fileId        - The fileId identifies the workbook where we want to
	 *                      create the worksheet
	 * @param worksheetName - The name of the worksheet we want to create
	 * @param graphClient   - The Microsoft's GraphServiceClient will allow to send
	 *                      the request to the Microsoft Graph API
	 * 
	 * @return <code>true</code> if the sheet exist or <code>false</code> if it
	 *         doesn't
	 * 
	 */

	public static boolean checkExistingSheet(String fileId, String workSheetName, GraphServiceClient graphClient) {
		checkNotNull(fileId);
		checkNotNull(workSheetName);
		checkNotNull(graphClient);

		boolean sheetExist = false;

		List<WorkbookWorksheet> listWorksheetExist = graphClient.me().drive().items(fileId).workbook().worksheets()
				.buildRequest().get().getCurrentPage();
		Iterator<WorkbookWorksheet> it = listWorksheetExist.iterator();

		while (it.hasNext() && !sheetExist) {
			if (it.next().name.equals(workSheetName)) {
				sheetExist = true;
			}
		}

		return sheetExist;
	}

	/**
	 * This method creates an empty Worksheet in the workbook identified by his
	 * fileId
	 * 
	 * @param fileId        - The fileId identifies the workbook where we want to
	 *                      create the worksheet
	 * @param worksheetName - The name of the worksheet we want to create
	 * @param graphClient   - The Microsoft's IGraphServiceClient will allow to send
	 *                      the request to the Microsoft Graph API
	 * 
	 * @see io.github.oliviercailloux.teach_spreadsheets.online.read.WorksheetReader#getFileId(String,
	 *      GraphServiceClient) To get the fileId of the desired workbook
	 */
	public static void createNewWorkSheet(String fileId, String workSheetName, GraphServiceClient graphClient) {

		boolean sheetExist = checkExistingSheet(fileId, workSheetName, graphClient);

		if (!sheetExist) {
			throw new IllegalStateException("This worksheet doesn't exist ");
		}

		graphClient.me().drive().items(fileId).workbook().worksheets(workSheetName).buildRequest().get();

	}

	/**
	 * This method allow to write in a worksheet using the row and the column of the
	 * cell (example : A1 => (row=0,clomn=0)
	 * 
	 * @param fileId        - The fileId identifies the workbook where the
	 *                      worksheet/ * is located
	 * @param worksheetName - The name of the worksheet we want to write into
	 * @param row           - the row of the cell which we want to read - the range
	 *                      of the value of row and column are >= 0
	 * @param column        - the column of the cell which we want to read - In
	 *                      excel the column's name are letter, here the letter is
	 *                      is translated by its alphabetical rank starting with 0
	 *                      (ex : A -> 0, B->1 ... ZA-> 26)
	 * @param text          - String we want to write
	 * @param graphClient   - The Microsoft's GraphServiceClient will allow to send
	 *                      the request to the Microsoft Graph API
	 *
	 * @see io.github.oliviercailloux.teach_spreadsheets.online.read.WorksheetReader#getFileId(String,
	 *      GraphServiceClient) To get the fileId of the desired workbook
	 */
	public static void writeStringIntheSheet(String fileId, String worksheetName, int row, int column, String text,
			GraphServiceClient graphClient) {
		checkArgument(row >= 0, column >= 0);
		boolean sheetExist = checkExistingSheet(fileId, worksheetName, graphClient);
		if (!sheetExist) {
			throw new IllegalStateException("This worksheet doesn't exist ");
		}
		String url = graphClient.me().drive().items(fileId).workbook().worksheets(worksheetName).buildRequest()
				.getRequestUrl().toString();

		WorkbookRange wR = new WorkbookRange();
		wR.values = new JsonPrimitive(text);
		CustomRequest<WorkbookRange> request1 = new CustomRequest<>(
				url + "/microsoft.graph.cell(row=" + row + ",column=" + column + ")", graphClient, null,
				WorkbookRange.class);
		request1.patch(wR);

	}

	/**
	 * This method allow to change the color of a cell using the row and the column
	 * of the cell (example : a1 =>(row=0,column=0))
	 * 
	 * @param fileId        - The fileId identifies the workbook where the
	 *                      worksheet/ * is located
	 * @param worksheetName - The name of the worksheet we want to write into
	 * @param row           - the row of the cell which we want to read - the range
	 *                      of the value of row and column are >= 0
	 * @param column        - the column of the cell which we want to read - In
	 *                      excel the column's name are letter, here the letter is
	 *                      is translated by its alphabetical rank starting with 0
	 *                      (ex : A -> 0, B->1 ... ZA-> 26)
	 * @param color         - color of the cell of the form #RRGGBB (e.g. 'FFA500')
	 *                      or as a named HTML color (e.g. 'orange')
	 * @param graphClient   - The Microsoft's GraphServiceClient will allow to send
	 *                      the request to the Microsoft Graph API
	 *
	 * @see io.github.oliviercailloux.teach_spreadsheets.online.read.WorksheetReader#getFileId(String,
	 *      GraphServiceClient) To get the fileId of the desired workbook
	 */
	public static void changeCellColor(String fileId, String worksheetName, int row, int column, String color,
			GraphServiceClient graphClient) {
		checkArgument(row >= 0, column >= 0);
		boolean sheetExist = checkExistingSheet(fileId, worksheetName, graphClient);
		if (!sheetExist) {
			throw new IllegalStateException("This worksheet doesn't exist ");
		}
		WorkbookRangeFill workbookRangeFill = new WorkbookRangeFill();
		workbookRangeFill.color = color;

		String url = graphClient.me().drive().items(fileId).workbook().worksheets(worksheetName)
				.range(WorkbookWorksheetRangeParameterSet.newBuilder().withAddress("").build()).format().fill()
				.buildRequest().getRequestUrl().toString();
		if (!url.contains("microsoft.graph.range")) {
			throw new IllegalStateException("Error with MS Graph Url");
		}
		String urlRequest = url.replace("microsoft.graph.range",
				"microsoft.graph.cell(row=" + row + ",column=" + column + ")");

		CustomRequest<WorkbookRangeFill> request = new CustomRequest<>(urlRequest, graphClient, null,
				WorkbookRangeFill.class);
		request.patch(workbookRangeFill);
	}

	/**
	 * This method allow to change the font of a cell using the row and the column
	 * of the cell (example : a1 =>(row=0,column=0))
	 * 
	 * @param fileId        - The fileId identifies the workbook where the worksheet
	 *                      is located
	 * @param worksheetName - The name of the worksheet we want to write into
	 * 
	 * @param graphClient   - The Microsoft's GraphServiceClient will allow to send
	 *                      the request to the Microsoft Graph API
	 * 
	 * @param row           - the row of the cell which we want to change its font :
	 *                      the range of the value of row and column are >= 0
	 * @param column        - the column of the cell which we want to its font : In
	 *                      excel the column's name are letter, here the letter is
	 *                      is translated by its alphabetical rank starting with 0
	 *                      (ex : A -> 0, B->1 ... ZA-> 26)
	 * @param color         - color of the cell of the form #RRGGBB (e.g. 'FFA500')
	 *                      or as a named HTML color (e.g. 'orange')
	 * @param bold
	 * @param size          - size of the cell
	 * @param name          - police name
	 * 
	 * @see <a href=
	 *      "https://docs.microsoft.com/fr-fr/graph/api/resources/rangefont?view=graph-rest-1.0">
	 *      The Microsoft doc that helped to implement this function online </a>
	 * 
	 */
	public void setFont(String fileId, String worksheetName, GraphServiceClient graphClient, int row, int column,
			Boolean bold, String color, Double size, String name) {
		checkArgument(row >= 0, column >= 0);
		WorkbookRangeFont workbookRangeFont = new WorkbookRangeFont();
		workbookRangeFont.bold = bold;
		workbookRangeFont.color = color;
		workbookRangeFont.size = size;
		workbookRangeFont.name = name;
		String url = graphClient.me().drive().items(fileId).workbook().worksheets(worksheetName)
				.range(WorkbookWorksheetRangeParameterSet.newBuilder().withAddress("").build()).format().font()
				.buildRequest().getRequestUrl().toString();
		if (!url.contains("microsoft.graph.range")) {
			throw new IllegalStateException("Error with MS Graph Url");
		}
		String urlRequest = url.replace("microsoft.graph.range",
				"microsoft.graph.cell(row=" + row + ",column=" + column + ")");

		CustomRequest<WorkbookRangeFont> request = new CustomRequest<>(urlRequest, graphClient, null,
				WorkbookRangeFont.class);
		request.patch(workbookRangeFont);

	}

	/**
	 * This method allow to change the format of a cell using the row and the column
	 * of the cell (example : a1 =>(row=0,column=0))
	 * 
	 * @param fileId              - The fileId identifies the workbook where the
	 *                            worksheet/ * is located
	 * @param worksheetName       - The name of the worksheet we want to write into
	 * 
	 * @param graphClient         - The Microsoft's GraphServiceClient will allow to
	 *                            send the request to the Microsoft Graph API
	 * 
	 * @param row                 - the row of the cell which we want to change its
	 *                            format : the range of the value of row and column
	 *                            are >= 0
	 * @param column              - the column of the cell which we want to change
	 *                            its format : In excel the column's name are
	 *                            letter, here the letter is translated by its
	 *                            alphabetical rank starting with 0 (ex : A -> 0,
	 *                            B->1 ... ZA-> 26)
	 * @param columnWidth         - Width of the column
	 * @param alignmentHorizontal -text horizontal aligmnent
	 * @param alignmentVertical   - text vertical aligmnent
	 * 
	 * @see <a href=
	 *      "https://docs.microsoft.com/fr-fr/graph/api/resources/rangeformat?view=graph-rest-1.0">
	 *      The Microsoft doc that helped to implement this function online </a>
	 *
	 * 
	 */
	public void setFormat(String fileId, String worksheetName, GraphServiceClient graphClient, int row, int column,
			double columnWidth, String alignmentHorizontal, String alignmentVertical) {
		checkArgument(row >= 0, column >= 0);
		WorkbookRangeFormat workbookRangeFormat = new WorkbookRangeFormat();
		workbookRangeFormat.columnWidth = columnWidth;
		workbookRangeFormat.horizontalAlignment = alignmentHorizontal;
		workbookRangeFormat.verticalAlignment = alignmentVertical;
		String url = graphClient.me().drive().items(fileId).workbook().worksheets(worksheetName)
				.range(WorkbookWorksheetRangeParameterSet.newBuilder().withAddress("").build()).format().buildRequest()
				.getRequestUrl().toString();
		if (!url.contains("microsoft.graph.range")) {
			throw new IllegalStateException("Error with MS Graph Url");
		}
		String urlRequest = url.replace("microsoft.graph.range",
				"microsoft.graph.cell(row=" + row + ",column=" + column + ")");

		CustomRequest<WorkbookRangeFormat> request = new CustomRequest<>(urlRequest, graphClient, null,
				WorkbookRangeFormat.class);
		request.patch(workbookRangeFormat);
	}

}