package io.github.oliviercailloux.tutorial;

import static com.google.common.base.Preconditions.checkArgument;
import static com.google.common.base.Preconditions.checkNotNull;

import com.google.gson.Gson;
import com.google.gson.JsonArray;
import com.google.gson.JsonElement;
import com.google.gson.JsonObject;
import com.google.gson.JsonParser;
import com.microsoft.graph.models.DriveSearchParameterSet;
import com.microsoft.graph.models.WorkbookWorksheetCellParameterSet;
import com.microsoft.graph.requests.DriveSearchCollectionPage;
import com.microsoft.graph.requests.GraphServiceClient;

public class WorksheetReader {

	/**
	 * @param fileName
	 * @param graphClient
	 * @return FileId
	 */
	public static String getFileId(String fileName, GraphServiceClient graphClient) {
		checkNotNull(fileName);
		checkNotNull(graphClient);
		DriveSearchCollectionPage search = graphClient.me().drive()
				.search(DriveSearchParameterSet.newBuilder().withQ(fileName).build()).buildRequest().get();
		Gson gson = new Gson();
		String item = gson.toJson(search);
		JsonObject itemJsonObect = (JsonObject) JsonParser.parseString(item);
		JsonArray content = (JsonArray) (itemJsonObect.get("pageContents"));
		if (content.size() != 1) {
			throw new IllegalStateException("size of page contents is not equal to 1");
		}
		String fileId = content.get(0).getAsJsonObject().get("id").getAsString();
		return fileId;

	}

	/**
	 * Read the value of worksheet cell, identified by his row and column
	 * 
	 * @param fileId
	 * @param worksheetName
	 * @param graphClient
	 * @param row           - the row of the cell which we want to read - the range
	 *                      of the value of row and column are >= 0
	 * @param column        - the column of the cell which we want to read - In
	 *                      excel the column's name are letter, here the letter is
	 *                      is translated by its alphabetical rank starting with 0
	 *                      (ex : A -> 0, B->1 ... ZA-> 26)
	 * @return - the value of the cell
	 */
	public static String getCellValue(String fileId, String worksheetName, GraphServiceClient graphClient, int row,
			int column) {
		checkArgument(row >= 0, column >= 0);

		JsonElement contentCell = graphClient.me().drive().items(fileId).workbook().worksheets(worksheetName)
				.cell(WorkbookWorksheetCellParameterSet.newBuilder().withRow(row).withColumn(column).build())
				.buildRequest().get().values;

		return contentCell.getAsString();
	}

}